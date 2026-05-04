"""
Excel Report Automation – Web App Backend
Flask server: handles multi-user concurrent uploads, sheet detection,
metrics computation, and output generation.
"""

import os, sys, uuid, json, shutil, zipfile, re, traceback
from datetime import datetime, timedelta
from pathlib import Path
from threading import Lock

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl
from lxml import etree

# ── App setup ─────────────────────────────────────────────────────────────────
APP_DIR      = Path(__file__).parent          # folder containing app.py and index.html
BASE_DIR     = APP_DIR.parent                 # outer excel_webapp folder
TEMPLATE_DIR = APP_DIR / "template"
SESSIONS_DIR = APP_DIR / "sessions"
SESSIONS_DIR.mkdir(exist_ok=True)

app = Flask(__name__, static_folder=str(APP_DIR), static_url_path="")
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max upload
CORS(app)

# ── Sheet signature: required columns (lowercase, partial match) ──────────────
SHEET_SIGNATURES = {
    "Bank_Statement_R": {
        "required": ["date", "amount"],
        "optional": ["status", "description"],
        "date_col": "Date",
        "priority_cols": ["date", "amount", "status"],
    },
    "Pay_06_R": {
        "required": ["textbox16", "check_amt"],
        "optional": [],
        "date_col": "textbox16",
        "priority_cols": ["textbox16", "check_amt"],
    },
    "Pay_41_R": {
        "required": ["trans_date", "payment", "financial_class"],
        "optional": ["clinic", "proc_code", "rendering_phy"],
        "date_col": "Trans_Date",
        "priority_cols": ["trans_date", "payment", "financial_class"],
    },
    "Age_24_R": {
        "required": ["financial_class", "textbox18"],
        "optional": ["currentamt", "over30amt", "over60amt"],
        "date_col": None,
        "priority_cols": ["financial_class", "textbox18", "currentamt"],
    },
    "Fin_25_R": {
        "required": ["textbox2", "total_charge"],
        "optional": ["pat_num"],
        "date_col": "Textbox2",
        "priority_cols": ["textbox2", "total_charge", "pat_num"],
    },
    "Waystar_R": {
        "required": ["trans date", "charges", "claim id"],
        "optional": ["patient name", "remit amount"],
        "date_col": "Trans Date",
        "priority_cols": ["trans date", "charges", "claim id"],
        # Waystar rows have NO data in Last Event Message
        "data_absent_col": "last event message",
    },
    "Rejected": {
        "required": ["trans date", "charges", "claim id", "last event message"],
        "optional": ["patient name", "last note"],
        "date_col": "Trans Date",
        "priority_cols": ["trans date", "charges", "claim id", "last event message"],
        # Rejected rows ALWAYS have data in Last Event Message
        "data_present_col": "last event message",
    },
    "Cnt_27": {
        "required": ["svc_date", "withhold_code"],
        "optional": ["rendering_phy", "clinic"],
        "date_col": "Svc_Date",
        "priority_cols": ["svc_date", "withhold_code"],
    },
}

REQUIRED_SHEETS = list(SHEET_SIGNATURES.keys())

# ══════════════════════════════════════════════════════════════════════════════
#  SHEET AUTO-DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def detect_header_row(filepath, sheet_name, max_scan=8, required_cols=None):
    try:
        raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=max_scan)
    except Exception:
        return 0
    required_cols = [c.strip().lower() for c in (required_cols or [])]
    for row_idx in range(len(raw)):
        row_vals = [str(v).strip() if pd.notna(v) else '' for v in raw.iloc[row_idx]]
        if not any(row_vals):
            continue
        first_val = row_vals[0]
        if not first_val:
            continue
        try:
            float(first_val.replace(',','').replace('$','').replace('(','-').replace(')',''))
            continue
        except ValueError:
            pass
        if first_val.startswith('='):
            continue
        if any(k in first_val.lower() for k in ('sum','total','grand')):
            continue
        if required_cols:
            row_lower = [v.lower() for v in row_vals]
            if all(any(req in cell for cell in row_lower) for req in required_cols):
                return row_idx
            continue
        return row_idx
    return 0


def identify_sheet_type(filepath, sheet_name):
    """
    Given a file and sheet name, determine which raw data type it matches.
    Returns (canonical_name, score, header_row) or (None, 0, 0)
    """
    best_match = None
    best_score = 0
    best_hdr   = 0

    for canonical, sig in SHEET_SIGNATURES.items():
        hdr = detect_header_row(filepath, sheet_name, max_scan=10,
                                required_cols=sig["required"])
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=hdr, nrows=3)
            df.columns = df.columns.str.strip().str.lower()
        except Exception:
            continue

        score = 0
        for req in sig["required"]:
            if any(req in col for col in df.columns):
                score += 10
        for opt in sig["optional"]:
            if any(opt in col for col in df.columns):
                score += 2

        if score >= len(sig["required"]) * 10:  # all required matched
            # Data-presence check: if signature requires a column to have actual
            # non-null data, verify it using a larger sample to avoid false negatives
            if "data_present_col" in sig:
                try:
                    df_check = pd.read_excel(filepath, sheet_name=sheet_name,
                                             header=hdr, nrows=200)
                    df_check.columns = df_check.columns.str.strip().str.lower()
                    target = sig["data_present_col"].lower()
                    matched_cols = [c for c in df_check.columns if target in c]
                    if not matched_cols or df_check[matched_cols[0]].notna().sum() == 0:
                        continue  # Column exists but is empty — not this sheet type
                except Exception:
                    pass

            # Data-absence check: if signature expects a column to be empty,
            # skip this canonical type when the column has actual data
            if "data_absent_col" in sig:
                try:
                    df_check = pd.read_excel(filepath, sheet_name=sheet_name,
                                             header=hdr, nrows=200)
                    df_check.columns = df_check.columns.str.strip().str.lower()
                    target = sig["data_absent_col"].lower()
                    matched_cols = [c for c in df_check.columns if target in c]
                    if matched_cols and df_check[matched_cols[0]].notna().sum() > 0:
                        continue  # Column has data — this is Rejected, not Waystar_R
                except Exception:
                    pass

            if score > best_score:
                best_score  = score
                best_match  = canonical
                best_hdr    = hdr

    return best_match, best_score, best_hdr


def filename_forced_canonical(filepath):
    """
    Determine if the filename explicitly identifies a Waystar or Rejected file.
    Returns 'Waystar_R', 'Rejected', or None.

    Rules (case-insensitive, match anywhere in the base filename):
      - Contains 'claimsscreen' AND 'all'  → Waystar_R
      - Contains 'claimsscreen' AND 'rejected' → Rejected
    """
    name = Path(filepath).stem.lower()
    if "claimsscreen" in name:
        if "all" in name:
            return "Waystar_R"
        if "rejected" in name:
            return "Rejected"
    return None


def scan_uploaded_file(filepath):
    """
    Scan all sheets in an uploaded Excel file, return detection results.
    Returns dict: { canonical_sheet_name: { 'source_sheet': str, 'header_row': int, 'confidence': int } }

    For files whose names explicitly identify them as Waystar or Rejected
    (e.g. 'Claimsscreen_all.xlsx' and 'Claimsscreen_rejected.xlsx'), the
    filename takes priority over column-based detection so that two files
    sharing the same column layout are never confused.
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names
    except Exception as e:
        return {}, str(e)

    # ── Filename-based forced identification ──────────────────────────────────
    forced_canonical = filename_forced_canonical(filepath)
    if forced_canonical:
        # Use the first sheet; detect its header row with the appropriate sig
        sheet = sheet_names[0]
        sig = SHEET_SIGNATURES[forced_canonical]
        hdr = detect_header_row(filepath, sheet, max_scan=10,
                                required_cols=sig["required"])
        return {
            forced_canonical: {
                "source_sheet": sheet,
                "header_row": hdr,
                "confidence": 999,   # highest confidence – name-based
                "original_name": sheet,
            }
        }, None

    # ── Column-based detection (all other files) ──────────────────────────────
    detected = {}
    used_sources = set()

    for sheet in sheet_names:
        canonical, score, hdr = identify_sheet_type(filepath, sheet)
        if canonical and canonical not in detected:
            detected[canonical] = {
                "source_sheet": sheet,
                "header_row": hdr,
                "confidence": score,
                "original_name": sheet,
            }
            used_sources.add(sheet)

    return detected, None


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LOADING (session-aware)
# ══════════════════════════════════════════════════════════════════════════════

XLSX_NS     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EXCEL_EPOCH = datetime(1899, 12, 30)

DME_COMM_PROCS = {
    "A9150","A7003","A6449","A4450","A6250","A6450","A4615","A6454","A4467",
    "A6448","A6216","A4314","A4452","A6222","A4566","A4550","E0114","E0110",
    "L3809","L1833","L4361","L3908","L1830","Q4049","L3660","L0642","L4350",
    "L1902","L4386","L3260","L0641","A4565","Q4018","L0120","L1820","L4360",
    "L3925","Q4020","L0174","Q4024","Q4012","L3670","Q4006","Q4042","L3980",
    "L3807","Q4022","Q4046","L3984","L1810","L3923","L0172",
}


def load_raw_data_from_session(session_dir, sheet_map):
    """
    Load raw data using the detected sheet mapping.
    sheet_map: { canonical_name: { 'file': path, 'source_sheet': name, 'header_row': int } }
    """
    data = {}
    errors = []

    def load(canonical):
        info = sheet_map.get(canonical)
        if not info:
            errors.append(f"Missing sheet: {canonical}")
            return None
        try:
            hdr = info.get("header_row", 0)
            df = pd.read_excel(info["file"], sheet_name=info["source_sheet"],
                               header=hdr, nrows=50000)
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            errors.append(f"Error loading {canonical}: {e}")
            return None

    # Bank_Statement_R
    bank = load("Bank_Statement_R")
    if bank is not None:
        bank['Date']   = pd.to_datetime(bank['Date'],   errors='coerce')
        bank['Amount'] = pd.to_numeric(bank['Amount'],  errors='coerce')
        bank['Status'] = bank.get('Status', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        if 'Status' not in bank.columns:
            bank['Status'] = ''
        bank = bank.dropna(subset=['Date'])
        data['bank'] = bank

    # Pay_06_R
    pay06 = load("Pay_06_R")
    if pay06 is not None:
        pay06['textbox16'] = pd.to_datetime(pay06['textbox16'], errors='coerce')
        pay06['Check_Amt'] = pd.to_numeric(pay06['Check_Amt'],  errors='coerce')
        data['pay06'] = pay06

    # Pay_41_R
    pay41 = load("Pay_41_R")
    if pay41 is not None:
        pay41['Trans_Date'] = pd.to_datetime(pay41['Trans_Date'], errors='coerce')
        pay41['Payment']    = pd.to_numeric(pay41['Payment'],     errors='coerce')
        if 'Proc_Code' in pay41.columns:
            pay41['Proc_Code'] = pay41['Proc_Code'].astype(str).str.strip().str.upper()
        else:
            pay41['Proc_Code'] = ''
        if 'Financial_Class' not in pay41.columns:
            pay41['Financial_Class'] = ''
        data['pay41'] = pay41

    # Age_24_R
    age24 = load("Age_24_R")
    if age24 is not None:
        for c in ['textbox18','CurrentAmt','Over30Amt','Over60Amt',
                  'Over90Amt','Over120Amt','Over150Amt']:
            if c in age24.columns:
                age24[c] = pd.to_numeric(age24[c], errors='coerce')
        data['age24'] = age24

    # Fin_25_R
    fin25 = load("Fin_25_R")
    if fin25 is not None:
        fin25['Textbox2']     = pd.to_datetime(fin25['Textbox2'],    errors='coerce')
        fin25['Total_Charge'] = pd.to_numeric(fin25['Total_Charge'], errors='coerce')
        data['fin25'] = fin25

    # Waystar_R
    waystar = load("Waystar_R")
    if waystar is not None:
        if 'Trans Date' not in waystar.columns:
            # try lowercase fallback
            waystar.columns = [c if c != 'trans date' else 'Trans Date'
                               for c in waystar.columns.str.lower()]
        waystar['Trans Date'] = pd.to_datetime(waystar['Trans Date'], errors='coerce')
        waystar['Charges']    = pd.to_numeric(waystar['Charges'],     errors='coerce')
        data['waystar'] = waystar

    # Rejected
    rejected = load("Rejected")
    if rejected is not None:
        rejected['Trans Date'] = pd.to_datetime(rejected['Trans Date'], errors='coerce')
        lem_cols = [c for c in rejected.columns if 'Last Event' in c]
        rejected['Last_Event_Message'] = rejected[lem_cols[0]] if lem_cols else ''
        rejected = rejected.dropna(subset=['Trans Date'])
        data['rejected'] = rejected

    # Cnt_27
    cnt27 = load("Cnt_27")
    if cnt27 is not None:
        cnt27['Svc_Date'] = pd.to_datetime(cnt27['Svc_Date'], errors='coerce')
        data['cnt27'] = cnt27

    return data, errors


# ══════════════════════════════════════════════════════════════════════════════
#  METRICS
# ══════════════════════════════════════════════════════════════════════════════

def safe_sum(series):
    numeric = pd.to_numeric(series, errors='coerce')
    nan_mask = numeric.isna()
    if nan_mask.any():
        s = series.astype(str).str.strip()
        acct = s.str.match(r'^\([\d,\.]+\)$')
        if acct.any():
            cleaned = s.where(~acct, '-' + s.str[1:-1])
            cleaned = cleaned.str.replace(',','',regex=False).str.replace('$','',regex=False)
            fixed = pd.to_numeric(cleaned, errors='coerce')
            numeric = numeric.where(~acct, fixed)
    return float(numeric.sum())

def safe_count(series):   return int(series.count())
def safe_distinct(series): return int(series.nunique(dropna=True))

def filter_date(df, col, td):
    return df[df[col].dt.date == td]

def filter_mtd(df, col, ms, td):
    return df[(df[col].dt.date >= ms) & (df[col].dt.date <= td)]


def compute_metrics(data, target_date):
    td = target_date.date()
    ms = target_date.replace(day=1).date()
    m  = {}

    bank = data.get('bank', pd.DataFrame())
    pay06 = data.get('pay06', pd.DataFrame())
    pay41 = data.get('pay41', pd.DataFrame())
    age24 = data.get('age24', pd.DataFrame())
    fin25 = data.get('fin25', pd.DataFrame())
    waystar = data.get('waystar', pd.DataFrame())
    rejected = data.get('rejected', pd.DataFrame())
    cnt27 = data.get('cnt27', pd.DataFrame())

    if not bank.empty:
        bd = filter_date(bank,'Date',td)
        m['daily_deposits'] = safe_sum(bd['Amount'])
        m['mtd_deposits']   = safe_sum(filter_mtd(bank,'Date',ms,td)['Amount'])
        xfer = bank['Status'] == 'No Action Required - Transfer'
        m['daily_transfer'] = safe_sum(bank[xfer & (bank['Date'].dt.date==td)]['Amount'])
        m['mtd_transfer']   = safe_sum(bank[xfer & (bank['Date'].dt.date>=ms) & (bank['Date'].dt.date<=td)]['Amount'])
    else:
        m.update({'daily_deposits':0,'mtd_deposits':0,'daily_transfer':0,'mtd_transfer':0})

    if not pay06.empty:
        p06d = filter_date(pay06,'textbox16',td)
        m['daily_pay06'] = safe_sum(p06d['Check_Amt'])
        m['mtd_pay06']   = safe_sum(filter_mtd(pay06,'textbox16',ms,td)['Check_Amt'])
    else:
        m.update({'daily_pay06':0,'mtd_pay06':0})

    if not pay41.empty:
        pat_m = pay41['Financial_Class'].str.contains(r'\bPAT\b|Self\s*Pay',case=False,na=False,regex=True)
        m['daily_pat41'] = safe_sum(filter_date(pay41[pat_m],'Trans_Date',td)['Payment'])
        m['mtd_pat41']   = safe_sum(filter_mtd(pay41[pat_m],'Trans_Date',ms,td)['Payment'])

        dme_m   = pay41['Financial_Class'].str.contains('DME',case=False,na=False)
        dme_c_m = dme_m & pay41['Proc_Code'].isin(DME_COMM_PROCS)
        dme_o_m = dme_m & ~dme_c_m
        m['daily_dme_comm_pat41'] = safe_sum(filter_date(pay41[dme_c_m],'Trans_Date',td)['Payment'])
        m['mtd_dme_comm_pat41']   = safe_sum(filter_mtd(pay41[dme_c_m],'Trans_Date',ms,td)['Payment'])
        m['daily_dme_pat41'] = safe_sum(filter_date(pay41[dme_o_m],'Trans_Date',td)['Payment'])
        m['mtd_dme_pat41']   = safe_sum(filter_mtd(pay41[dme_o_m],'Trans_Date',ms,td)['Payment'])

        eps_m = pay41['Financial_Class'].str.contains(r'EMP|EPS',case=False,na=False,regex=True)
        m['daily_eps_pat41'] = safe_sum(filter_date(pay41[eps_m],'Trans_Date',td)['Payment'])
        m['mtd_eps_pat41']   = safe_sum(filter_mtd(pay41[eps_m],'Trans_Date',ms,td)['Payment'])

        wc_m = pay41['Financial_Class'].str.contains(r'WC|Auto|MVA|Work\w*Comp',case=False,na=False,regex=True)
        m['daily_wc_auto_pat41'] = safe_sum(filter_date(pay41[wc_m],'Trans_Date',td)['Payment'])
        m['mtd_wc_auto_pat41']   = safe_sum(filter_mtd(pay41[wc_m],'Trans_Date',ms,td)['Payment'])
    else:
        for k in ['daily_pat41','mtd_pat41','daily_dme_comm_pat41','mtd_dme_comm_pat41',
                  'daily_dme_pat41','mtd_dme_pat41','daily_eps_pat41','mtd_eps_pat41',
                  'daily_wc_auto_pat41','mtd_wc_auto_pat41']:
            m[k] = 0

    if not age24.empty:
        # Step 1: Pre-filter — keep only rows where textbox18 > 0.
        # This removes fully zero or missing rows before any AR calculation.
        if 'textbox18' in age24.columns:
            age24 = age24[age24['textbox18'] > 0]

        # Step 2: Drop rows where textbox18 is NaN (non-numeric / blank).
        # Include ALL remaining rows regardless of sign — negative values are valid AR credits
        # and must be included so they correctly offset positive balances.
        # (Previously `textbox18 > 0` silently dropped negative rows, causing
        #  totals like 120 + 20 + (−20) to return 140 instead of the correct 120.)
        ad = age24.dropna(subset=['textbox18']) if 'textbox18' in age24.columns else age24
        fc = ad['Financial_Class'].fillna('') if 'Financial_Class' in ad.columns else pd.Series([''] * len(ad), index=ad.index)

        # Exclude 11-DMERC financial class from ALL AR calculations:
        # Total AR, Total EPS AR, Total Insurance AR, Insurance buckets (CURRENT/31-60/61-90/
        # 91-120/121-150/151+), % AR > 120, Total Patient AR, Patient buckets, % AR > 120
        dmerc_m = fc.str.contains(r'11[\s\-]*DMERC', case=False, na=False, regex=True)
        ad = ad[~dmerc_m]  # drop 11-DMERC rows for all AR metrics
        # Refresh fc after filtering
        fc = ad['Financial_Class'].fillna('') if 'Financial_Class' in ad.columns else pd.Series([''] * len(ad), index=ad.index)

        eps_m = fc.str.contains(r'\bEPS\b',case=False,na=False,regex=True)
        emp_m = fc.str.contains(r'\bEMP\b',case=False,na=False,regex=True)
        sp_m  = fc.str.contains(r'Self\s*Pay|\bPAT\b',case=False,na=False,regex=True)
        wc_m  = fc.str.contains(r'\bWC\b|Auto|MVA|Work[\s\-]?Comp',case=False,na=False,regex=True)
        excl  = ~(eps_m | emp_m | sp_m | wc_m)

        # Total EPS AR: sum of textbox18 for EPS rows (negatives included, 11-DMERC excluded)
        m['total_eps'] = safe_sum(ad[eps_m]['textbox18']) if 'textbox18' in ad.columns else 0

        # Insurance and Patient AR buckets: each age-band column sums
        # positives AND negatives so credits properly reduce the total.
        # 11-DMERC rows are already excluded from `ad` above.
        for bucket, mask in [('ins', excl), ('pat', sp_m)]:
            for col_sfx, src_col in [('current','CurrentAmt'),('31_60','Over30Amt'),
                                      ('61_90','Over60Amt'),('91_120','Over90Amt'),
                                      ('121_150','Over120Amt'),('151plus','Over150Amt')]:
                m[f'{bucket}_{col_sfx}'] = safe_sum(ad[mask][src_col]) if src_col in ad.columns else 0

        # Total Insurance AR and Total Patient AR are the net sum of all age bands
        # (including negative values, excluding 11-DMERC), matching the real AR balance.
        ti = sum(m.get(f'ins_{s}',0) for s in ['current','31_60','61_90','91_120','121_150','151plus'])
        tp = sum(m.get(f'pat_{s}',0) for s in ['current','31_60','61_90','91_120','121_150','151plus'])
        m['total_ins']   = ti
        m['total_pat']   = tp
        m['pct_ins_120'] = ((m['ins_121_150']+m['ins_151plus'])/ti) if ti else 0
        m['pct_pat_120'] = ((m['pat_121_150']+m['pat_151plus'])/tp) if tp else 0
        # Total AR = EPS AR + Insurance AR + Patient AR (all net of negatives, 11-DMERC excluded)
        m['total_ar']    = m['total_eps'] + ti + tp
    else:
        for k in ['total_eps','total_ins','total_pat','total_ar','pct_ins_120','pct_pat_120',
                  'ins_current','ins_31_60','ins_61_90','ins_91_120','ins_121_150','ins_151plus',
                  'pat_current','pat_31_60','pat_61_90','pat_91_120','pat_121_150','pat_151plus']:
            m[k] = 0

    if not fin25.empty:
        f25d = filter_date(fin25,'Textbox2',td)
        m['daily_claim_count']   = safe_count(f25d['Pat_Num']) if 'Pat_Num' in f25d.columns else len(f25d)
        m['mtd_claim_count']     = safe_count(filter_mtd(fin25,'Textbox2',ms,td)['Pat_Num']) if 'Pat_Num' in fin25.columns else 0
        m['daily_claim_charges'] = safe_sum(f25d['Total_Charge'])
        m['mtd_claim_charges']   = safe_sum(filter_mtd(fin25,'Textbox2',ms,td)['Total_Charge'])
    else:
        m.update({'daily_claim_count':0,'mtd_claim_count':0,'daily_claim_charges':0,'mtd_claim_charges':0})

    if not waystar.empty:
        wsd = filter_date(waystar,'Trans Date',td)
        m['daily_waystar_amt'] = safe_sum(wsd['Charges'])
        m['mtd_waystar_amt']   = safe_sum(filter_mtd(waystar,'Trans Date',ms,td)['Charges'])
        m['daily_waystar_cnt'] = safe_distinct(wsd['Claim ID']) if 'Claim ID' in wsd.columns else 0
        m['mtd_waystar_cnt']   = safe_distinct(filter_mtd(waystar,'Trans Date',ms,td)['Claim ID']) if 'Claim ID' in waystar.columns else 0
    else:
        m.update({'daily_waystar_amt':0,'mtd_waystar_amt':0,'daily_waystar_cnt':0,'mtd_waystar_cnt':0})

    if not rejected.empty:
        rejd = filter_date(rejected, 'Trans Date', td)
        # Return list of Last_Event_Message values for rows matching the target date
        m['rejected'] = rejd['Last_Event_Message'].dropna().tolist() if not rejd.empty else []
        m['rejected_count'] = len(rejd)
    else:
        m['rejected'] = []
        m['rejected_count'] = 0

    if not cnt27.empty:
        # Waiting for Coding: any of these exact codes
        WAITING_CODING_CODES = {
            'Diag Pending', 'E-M Code', 'CPT Code', 'ClinicalReview',
            'J-Code Hold', 'Exam Rec Incomp', 'Coding Review',
        }
        wc_mask = (cnt27['Withhold_Code'].isin(WAITING_CODING_CODES)
                   if 'Withhold_Code' in cnt27.columns
                   else pd.Series([False] * len(cnt27)))
        m['waiting_coding'] = safe_count(cnt27[wc_mask]['Withhold_Code']) if 'Withhold_Code' in cnt27.columns else 0

        # Cred Issues: Withhold_Code contains 'Cred' anywhere (start, middle, end),
        # case-insensitive
        cred_mask = (cnt27['Withhold_Code'].str.contains('Cred', case=False, na=False)
                     if 'Withhold_Code' in cnt27.columns
                     else pd.Series([False] * len(cnt27)))
        m['cred_issues'] = safe_count(cnt27[cred_mask]['Withhold_Code']) if 'Withhold_Code' in cnt27.columns else 0

        # Pending Charts: everything that is NOT Waiting for Coding AND NOT Cred Issues
        pend = (cnt27['Withhold_Code'].notna() & ~wc_mask & ~cred_mask
                if 'Withhold_Code' in cnt27.columns
                else pd.Series([False] * len(cnt27)))
        m['pending_charts'] = safe_count(cnt27[pend]['Withhold_Code']) if 'Withhold_Code' in cnt27.columns else 0
    else:
        m.update({'waiting_coding': 0, 'pending_charts': 0, 'cred_issues': 0})

    return m


# ══════════════════════════════════════════════════════════════════════════════
#  SUMMARY XML PATCH (from original script, unchanged logic)
# ══════════════════════════════════════════════════════════════════════════════

SUMMARY_ROWS = [
    (5,False,'daily_deposits'),(6,True,'mtd_deposits'),
    (7,False,'daily_pay06'),(8,True,'mtd_pay06'),
    (9,False,'daily_transfer'),(10,True,'mtd_transfer'),
    (11,False,'daily_pat41'),(12,True,'mtd_pat41'),
    (13,False,'daily_dme_pat41'),(14,True,'mtd_dme_pat41'),
    (15,False,'daily_dme_comm_pat41'),(16,True,'mtd_dme_comm_pat41'),
    (17,False,'daily_eps_pat41'),(18,True,'mtd_eps_pat41'),
    (19,False,'daily_wc_auto_pat41'),(20,True,'mtd_wc_auto_pat41'),
    (21,False,'total_ar'),(22,False,'total_eps'),(23,False,'total_ins'),
    (24,False,'ins_current'),(25,False,'ins_31_60'),(26,False,'ins_61_90'),
    (27,False,'ins_91_120'),(28,False,'ins_121_150'),(29,False,'ins_151plus'),
    (30,False,'pct_ins_120'),(31,False,'total_pat'),(32,False,'pat_current'),
    (33,False,'pat_31_60'),(34,False,'pat_61_90'),(35,False,'pat_91_120'),
    (36,False,'pat_121_150'),(37,False,'pat_151plus'),(38,False,'pct_pat_120'),
    (39,False,'daily_claim_count'),(40,True,'mtd_claim_count'),
    (41,False,'daily_claim_charges'),(42,True,'mtd_claim_charges'),
    (43,False,'daily_waystar_amt'),(44,False,'daily_waystar_cnt'),(45,True,'mtd_waystar_cnt'),
    (46,False,'rejected_count'),(47,False,'waiting_coding'),
    (48,False,'pending_charts'),(49,False,'cred_issues'),
]

# Rows that must ONLY be written for the current (latest/target) date column.
# All other date columns must be left blank for these rows.
# Covers: Total AR, Total EPS AR, Total Insurance AR, all Insurance age buckets,
# % AR > 120 (Insurance), Total Patient AR, all Patient age buckets,
# % AR > 120 (Patient), Waiting for Coding, Pending Charts, Provider/Credentialling Issues.
CURRENT_DATE_ONLY_ROWS = {
    21,  # Total AR (Age-24)
    22,  # Total EPS AR
    23,  # Total Insurance AR
    24,  # Insurance CURRENT
    25,  # Insurance 31-60
    26,  # Insurance 61-90
    27,  # Insurance 91-120
    28,  # Insurance 121-150
    29,  # Insurance 151+
    30,  # % AR > 120 (Insurance)
    31,  # Total Patient AR (Age-24)
    32,  # Patient CURRENT
    33,  # Patient 31-60
    34,  # Patient 61-90
    35,  # Patient 91-120
    36,  # Patient 121-150
    37,  # Patient 151+
    38,  # % AR > 120 (Patient)
    47,  # Waiting for Coding (CNT-27)
    48,  # Pending Charts (CNT-27)
    49,  # Provider / Credentialling Issues (CNT-27)
}

def _e(tag, **attrs):
    return etree.Element(f"{{{XLSX_NS}}}{tag}", **attrs)

def col_to_letter(n):
    result = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def letter_to_col(s):
    result = 0
    for ch in s.upper():
        result = result * 26 + (ord(ch) - 64)
    return result

def cref(row, col):
    return f"{col_to_letter(col)}{row}"

def num_cell(ref, val):
    c = _e('c', r=ref)
    v = _e('v'); v.text = str(round(float(val), 10)); c.append(v)
    return c

def str_cell(ref, text):
    c = _e('c', r=ref, t='inlineStr')
    is_ = _e('is'); t = _e('t'); t.text = str(text) if text else ''
    is_.append(t); c.append(is_); return c

def find_target_col(xml_bytes, target_date):
    serial = (target_date - EXCEL_EPOCH).days
    tree = etree.fromstring(xml_bytes)
    sd = tree.find(f'{{{XLSX_NS}}}sheetData')
    for row in sd.findall(f'{{{XLSX_NS}}}row'):
        if int(row.get('r')) != 3:
            continue
        for c in row.findall(f'{{{XLSX_NS}}}c'):
            v = c.find(f'{{{XLSX_NS}}}v')
            if v is None or v.text is None:
                continue
            try:
                if int(float(v.text)) == serial:
                    m = re.match(r'([A-Z]+)', c.get('r',''))
                    if m:
                        return letter_to_col(m.group(1))
            except (ValueError, TypeError):
                pass
    return None

def update_summary_xml_multi(xml_bytes, col_metrics, current_col=None):
    """
    Write computed metrics into the Summary sheet XML.

    col_metrics : { col_index: metrics_dict }
    current_col : the column index that represents today / the latest date.
                  Rows listed in CURRENT_DATE_ONLY_ROWS are written ONLY into
                  this column; all other date columns are left untouched for
                  those rows.  If current_col is None it is inferred as the
                  maximum column index present in col_metrics.
    """
    tree = etree.fromstring(xml_bytes)
    sd   = tree.find(f'{{{XLSX_NS}}}sheetData')
    row_map = {int(r.get('r')): r for r in sd.findall(f'{{{XLSX_NS}}}row')}

    # Determine which column is "current" (latest date)
    if current_col is None and col_metrics:
        current_col = max(col_metrics.keys())

    for target_col, metrics in col_metrics.items():
        for (rn, is_formula, key) in SUMMARY_ROWS:
            # Skip CURRENT_DATE_ONLY_ROWS for every column except the current date
            if rn in CURRENT_DATE_ONLY_ROWS and target_col != current_col:
                continue
            val = metrics.get(key)
            if val is None:
                continue
            ref    = cref(rn, target_col)
            row_el = row_map.get(rn)
            if row_el is None:
                continue
            existing = next((c for c in row_el.findall(f'{{{XLSX_NS}}}c') if c.get('r')==ref), None)
            if isinstance(val, list):
                val = '; '.join(str(v) for v in val if v) if val else ''
            if isinstance(val, str):
                if existing is not None:
                    row_el.remove(existing)
                row_el.append(str_cell(ref, val))
                continue
            num_val = float(val)
            if existing is None:
                row_el.append(num_cell(ref, num_val))
            elif is_formula:
                v_el = existing.find(f'{{{XLSX_NS}}}v')
                if v_el is None:
                    v_el = _e('v'); existing.append(v_el)
                v_el.text = str(round(num_val, 10))
            else:
                for child in list(existing):
                    existing.remove(child)
                existing.attrib.pop('t', None)
                v_el = _e('v'); v_el.text = str(round(num_val, 10))
                existing.append(v_el)
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


# ══════════════════════════════════════════════════════════════════════════════
#  STYLED SHEET BUILDERS (openpyxl)
# ══════════════════════════════════════════════════════════════════════════════

_THIN   = Side(style='thin', color='B0C4DE')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
HDR_FILL = PatternFill('solid', fgColor='1F3864')
HDR_FONT = Font(bold=True, color='FFFFFF', size=10)
TOT_FILL = PatternFill('solid', fgColor='2E75B6')
TOT_FONT = Font(bold=True, color='FFFFFF', size=10)
ALT_FILL = PatternFill('solid', fgColor='D6E4F0')
WHT_FILL = PatternFill('solid', fgColor='FFFFFF')
BODY_FONT = Font(size=10)
C_RIGHT  = Alignment(horizontal='right', vertical='center')
C_LEFT   = Alignment(horizontal='left',  vertical='center')
C_CENTER = Alignment(horizontal='center',vertical='center')
DATE_FMT  = 'M/D/YYYY'
MONEY_FMT = '#,##0.00'

def _hdr(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.fill=HDR_FILL; c.font=HDR_FONT; c.border=_BORDER; c.alignment=C_CENTER

def _data(ws, row, col, value, fmt=None, is_alt=False, align=C_RIGHT):
    c = ws.cell(row=row, column=col, value=value)
    c.fill=ALT_FILL if is_alt else WHT_FILL; c.font=BODY_FONT
    c.border=_BORDER; c.alignment=align
    if fmt: c.number_format=fmt

def _total(ws, row, col, value, fmt=None, align=C_RIGHT):
    c = ws.cell(row=row, column=col, value=value)
    c.fill=TOT_FILL; c.font=TOT_FONT; c.border=_BORDER; c.alignment=align
    if fmt: c.number_format=fmt

def build_dcf_sheet(wb, bank_df):
    ws = wb.create_sheet('DAILY CASH FLOW')
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 20
    _hdr(ws,1,1,'Deposit Date'); _hdr(ws,1,2,'Daily Cash Flow')
    b = bank_df.dropna(subset=['Date'])
    daily = b.groupby(b['Date'].dt.date)['Amount'].sum().reset_index().rename(columns={'Date':'d','Amount':'amt'}).sort_values('d')
    xfer_total = safe_sum(b[b['Status']=='No Action Required - Transfer']['Amount'])
    amounts = []
    for ri,(_, r) in enumerate(daily.iterrows(), 2):
        alt=(ri%2==0)
        _data(ws,ri,1,r['d'],DATE_FMT,alt,C_LEFT)
        _data(ws,ri,2,round(float(r['amt']),2),MONEY_FMT,alt)
        amounts.append(float(r['amt']))
    nr=len(amounts)+2
    mtd_t=round(sum(amounts),2); avg_v=round(mtd_t/len(amounts),2) if amounts else 0
    for i,(label,val) in enumerate([('MTD $',mtd_t),('Avg',avg_v),('Internal Transfer',round(xfer_total,2))]):
        _total(ws,nr+i,1,label,align=C_LEFT); _total(ws,nr+i,2,val,MONEY_FMT)

def build_rejections_sheet(wb, rejected_df, date_start, date_end):
    ws = wb.create_sheet('REJECTIONS')
    ws.column_dimensions['A'].width=16; ws.column_dimensions['B'].width=70
    _hdr(ws,1,1,'Date'); _hdr(ws,1,2,'Rejection Reason')
    ds = date_start.date() if isinstance(date_start, datetime) else date_start
    de = date_end.date()   if isinstance(date_end,   datetime) else date_end
    dr = rejected_df[
        (rejected_df['Trans Date'].dt.date >= ds) &
        (rejected_df['Trans Date'].dt.date <= de)
    ].sort_values('Trans Date')
    for ri,(_, r) in enumerate(dr.iterrows(), 2):
        alt=(ri%2==0)
        d=r['Trans Date']
        reason=r.get('Last_Event_Message','')
        _data(ws,ri,1,d.date() if pd.notna(d) else None,DATE_FMT,alt,C_LEFT)
        _data(ws,ri,2,str(reason) if pd.notna(reason) else '',None,alt,Alignment(wrap_text=True,vertical='center'))

def build_kpi_sheet(wb, pay41_df, cnt27_df, target_date):
    ws = wb.create_sheet('KPI')
    td=target_date.date(); ms=target_date.replace(day=1).date()
    p41=pay41_df.dropna(subset=['Trans_Date','Clinic']) if 'Clinic' in pay41_df.columns else pd.DataFrame()
    if not p41.empty:
        p41_mtd=p41[(p41['Trans_Date'].dt.date>=ms)&(p41['Trans_Date'].dt.date<=td)]
        pivot=(p41_mtd.groupby([p41_mtd['Trans_Date'].dt.date,'Clinic'])['Payment'].sum().unstack(fill_value=0).reset_index().sort_values('Trans_Date'))
        clinics=[c for c in pivot.columns if c!='Trans_Date']
        n_cols=len(clinics)+2
        ws.column_dimensions['A'].width=13
        for ci in range(2,n_cols+1): ws.column_dimensions[gcl(ci)].width=15
        for ci,val in enumerate(['Trans Date']+clinics+['TOTAL'],1): _hdr(ws,1,ci,val)
        col_totals=[0.0]*len(clinics)
        for ri,(_,r) in enumerate(pivot.iterrows(),2):
            alt=(ri%2==0)
            _data(ws,ri,1,r['Trans_Date'],DATE_FMT,alt,C_LEFT)
            rt=0.0
            for ci,clinic in enumerate(clinics):
                v=round(float(r.get(clinic,0)),2)
                _data(ws,ri,ci+2,v,MONEY_FMT,alt); col_totals[ci]+=v; rt+=v
            _data(ws,ri,n_cols,round(rt,2),MONEY_FMT,alt)
        trn=len(pivot)+2; _total(ws,trn,1,'TOTAL',align=C_LEFT)
        gt=0.0
        for ci,ct in enumerate(col_totals): _total(ws,trn,ci+2,round(ct,2),MONEY_FMT); gt+=ct
        _total(ws,trn,n_cols,round(gt,2),MONEY_FMT)
        t2_hdr=trn+2
    else:
        t2_hdr=3
    for ci,val in enumerate(['Rendering Physician','Clinic','CredIssue Count'],1): _hdr(ws,t2_hdr,ci,val)
    ws.column_dimensions['A'].width=max(ws.column_dimensions['A'].width,28)
    ws.column_dimensions['B'].width=max(ws.column_dimensions['B'].width,20)
    ws.column_dimensions['C'].width=max(ws.column_dimensions['C'].width,16)
    if not cnt27_df.empty and 'Withhold_Code' in cnt27_df.columns:
        cred=cnt27_df[cnt27_df['Withhold_Code'].str.contains('Cred', case=False, na=False)]
        if cred.empty:
            ws.cell(row=t2_hdr+1,column=1).value='No CredIssue records found'
        else:
            grp=(cred.groupby(['Rendering_Phy','Clinic']).size().reset_index(name='Count').sort_values(['Rendering_Phy','Clinic']))
            for ri,(_,r) in enumerate(grp.iterrows(),t2_hdr+1):
                alt=(ri%2==0)
                _data(ws,ri,1,str(r['Rendering_Phy']),None,alt,C_LEFT)
                _data(ws,ri,2,str(r['Clinic']),None,alt,C_LEFT)
                _data(ws,ri,3,int(r['Count']),None,alt)


def extend_summary_columns(ws, target_date):
    from datetime import date, timedelta
    td = target_date.date()
    col_by_date={}; last_hc_idx=None; last_hc_col=None; last_hc_date=None
    last_col_idx=0; prev_mtd_col=None
    row3_cells=list(ws[3])
    for i,cell in enumerate(row3_cells):
        if cell.column>last_col_idx: last_col_idx=cell.column
        val=cell.value
        if isinstance(val, datetime):
            d=val.date(); col_by_date[d]=cell.column
            last_hc_idx=i; last_hc_col=cell.column; last_hc_date=d
        elif isinstance(val,str) and 'MTD' in val.upper():
            prev_mtd_col=cell.column
    if last_hc_date is None: return
    n_formula=0
    for i in range(last_hc_idx+1, len(row3_cells)):
        v=row3_cells[i].value
        if isinstance(v,str) and v.startswith("=") and "+1" in v: n_formula+=1
        else: break
    last_existing_date=last_hc_date+timedelta(days=n_formula)
    for k in range(1,n_formula+1): col_by_date[last_hc_date+timedelta(days=k)]=last_hc_col+k
    if td<=last_existing_date: return
    plan=[]; next_col=last_col_idx+1; cursor=last_existing_date+timedelta(days=1)
    current_month=last_existing_date.month; current_year=last_existing_date.year
    while cursor<=td:
        if cursor.month!=current_month or cursor.year!=current_year:
            plan.append((next_col,'mtd',date(current_year,current_month,1))); next_col+=1
            current_month=cursor.month; current_year=cursor.year
        plan.append((next_col,'date',cursor)); next_col+=1; cursor+=timedelta(days=1)
    if not plan: return
    def month_cols(year,month):
        fc,lc=None,None
        for d,c in col_by_date.items():
            if d.year==year and d.month==month:
                fc=c if fc is None else min(fc,c); lc=c if lc is None else max(lc,c)
        for (pc,pk,pv) in plan:
            if pk=='date' and pv and pv.year==year and pv.month==month:
                fc=pc if fc is None else min(fc,pc); lc=pc if lc is None else max(lc,pc)
        return fc,lc
    DAILY_ROWS={5,7,9,11,13,15,17,19,39,41,43}
    RUNNING_MTD_ROWS={6,8,10,12,14,16,18,20,40,42,45}
    ALL_DATA_ROWS=set(range(5,50))
    running_mtd_base={r:prev_mtd_col for r in RUNNING_MTD_ROWS}
    from openpyxl.utils import get_column_letter
    for (col_idx,kind,val) in plan:
        cl=get_column_letter(col_idx); prev_cl=get_column_letter(col_idx-1)
        if kind=='date':
            d=val
            ws.cell(row=3,column=col_idx).value=datetime(d.year,d.month,d.day) if d.day==1 else f'={prev_cl}3+1'
            ws.cell(row=3,column=col_idx).number_format='M/D/YYYY'
            for row in RUNNING_MTD_ROWS:
                base=running_mtd_base[row]
                ws.cell(row=row,column=col_idx).value=(f'={cl}{row-1}+{get_column_letter(base)}{row}' if base else f'={cl}{row-1}')
            col_by_date[d]=col_idx
        elif kind=='mtd':
            year,month=val.year,val.month
            label=datetime(year,month,1).strftime("%b\\'%y")+' MTD'
            ws.cell(row=3,column=col_idx).value=label
            fc,lc=month_cols(year,month)
            lc_cl=get_column_letter(lc) if lc else cl; fc_cl=get_column_letter(fc) if fc else cl
            for row in ALL_DATA_ROWS:
                c=ws.cell(row=row,column=col_idx)
                if row in DAILY_ROWS:       c.value=f'=SUM({fc_cl}{row}:{lc_cl}{row})'
                elif row in RUNNING_MTD_ROWS: c.value=f'={lc_cl}{row}'
                elif row==30:               c.value=f'=SUM({cl}28:{cl}29)/{cl}23'
                elif row==38:               c.value=f'=SUM({cl}36:{cl}37)/{cl}31'
                else:                       c.value=f'={lc_cl}{row}'
            for r in RUNNING_MTD_ROWS: running_mtd_base[r]=col_idx


# ══════════════════════════════════════════════════════════════════════════════
#  BUILD FINAL OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def build_output(template_path, output_path, col_metrics_map, data, latest_date, date_start=None, date_end=None):
    with zipfile.ZipFile(template_path, 'r') as z:
        summary_xml = z.read("xl/worksheets/sheet1.xml")
    # Determine the current (latest) date column so CURRENT_DATE_ONLY_ROWS
    # are written only once (into the latest date column).
    current_col = max(col_metrics_map.keys()) if col_metrics_map else None
    new_summary = update_summary_xml_multi(summary_xml, col_metrics_map, current_col=current_col)
    tmp = output_path + '.tmp.xlsx'
    skip = {'xl/worksheets/sheet1.xml','xl/calcChain.xml'}
    with zipfile.ZipFile(template_path,'r') as zin, \
         zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename not in skip:
                zout.writestr(item, zin.read(item.filename))
        zout.writestr('xl/worksheets/sheet1.xml', new_summary)
    wb = openpyxl.load_workbook(tmp)
    if 'Summary' in wb.sheetnames:
        extend_summary_columns(wb['Summary'], latest_date)
    RAW_SHEETS=['Bank_Statement_R','Pay_06_R','Pay_41_R','Age_24_R',
                'Fin_25_R','Waystar_R','Rejected','Cnt_27']
    for name in RAW_SHEETS:
        if name in wb.sheetnames: del wb[name]
    for name in ['DAILY CASH FLOW','REJECTIONS','KPI']:
        if name in wb.sheetnames: del wb[name]
    build_dcf_sheet(wb, data['bank'] if 'bank' in data else pd.DataFrame(columns=['Date','Amount','Status']))
    if 'rejected' in data:
        ds = date_start if date_start else latest_date
        de = date_end   if date_end   else latest_date
        build_rejections_sheet(wb, data['rejected'], ds, de)
    if 'pay41' in data and 'cnt27' in data:
        build_kpi_sheet(wb, data['pay41'], data['cnt27'], latest_date)
    wb.save(output_path)
    try: os.remove(tmp)
    except: pass


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

sessions = {}
sessions_lock = Lock()

def get_session(session_id):
    with sessions_lock:
        if session_id not in sessions:
            sessions[session_id] = {
                "id": session_id,
                "dir": str(SESSIONS_DIR / session_id),
                "files": {},       # canonical_name → { file, source_sheet, header_row }
                "status": "idle",
            }
            os.makedirs(sessions[session_id]["dir"], exist_ok=True)
        return sessions[session_id]

def cleanup_old_sessions():
    """Remove sessions older than 2 hours."""
    import time
    cutoff = time.time() - 7200
    with sessions_lock:
        dead = [sid for sid, s in sessions.items()
                if os.path.getmtime(s["dir"]) < cutoff if os.path.exists(s["dir"])]
        for sid in dead:
            try: shutil.rmtree(sessions[sid]["dir"])
            except: pass
            del sessions[sid]


# ══════════════════════════════════════════════════════════════════════════════
#  API ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return send_from_directory(str(APP_DIR), "index.html")

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Maximum upload size is 100 MB."}), 413

@app.route("/api/session", methods=["POST"])
def create_session():
    sid = str(uuid.uuid4())
    get_session(sid)
    return jsonify({"session_id": sid})

@app.route("/api/session/<sid>/upload", methods=["POST"])
def upload_file(sid):
    """Upload one or more Excel files; auto-detect sheets inside each."""
    session = get_session(sid)
    if "files" not in request.files:
        return jsonify({"error": "No files provided"}), 400

    uploaded_files = request.files.getlist("files")
    results = []

    for f in uploaded_files:
        if not f.filename:
            continue
        safe_name = f"upload_{uuid.uuid4().hex}_{f.filename}"
        dest = os.path.join(session["dir"], safe_name)
        f.save(dest)

        # Detect all sheets in this file
        detected, err = scan_uploaded_file(dest)
        if err:
            results.append({"filename": f.filename, "error": err, "detected": {}})
            continue

        # Merge detections into session
        for canonical, info in detected.items():
            session["files"][canonical] = {
                "file": dest,
                "source_sheet": info["source_sheet"],
                "header_row": info["header_row"],
                "original_filename": f.filename,
                "original_sheet": info["original_name"],
                "confidence": info["confidence"],
            }

        results.append({
            "filename": f.filename,
            "detected": {k: {"source_sheet": v["source_sheet"], "confidence": v["confidence"]}
                         for k, v in detected.items()},
            "error": None,
        })

    return jsonify({
        "results": results,
        "session_sheets": {
            k: {"source_sheet": v["source_sheet"],
                "original_filename": v["original_filename"],
                "confidence": v["confidence"]}
            for k, v in session["files"].items()
        }
    })

@app.route("/api/session/<sid>/sheets", methods=["GET"])
def get_sheets(sid):
    session = get_session(sid)
    return jsonify({
        "sheets": {
            k: {"source_sheet": v["source_sheet"],
                "original_filename": v["original_filename"],
                "confidence": v["confidence"]}
            for k, v in session["files"].items()
        },
        "required": REQUIRED_SHEETS,
        "missing": [s for s in REQUIRED_SHEETS if s not in session["files"]],
    })

@app.route("/api/session/<sid>/detect", methods=["POST"])
def detect_sheets(sid):
    """Manually trigger re-detection or override a mapping."""
    session = get_session(sid)
    body = request.get_json() or {}
    overrides = body.get("overrides", {})  # { canonical_name: { file_idx, sheet_name } }
    # Re-scan existing files
    for canonical, override in overrides.items():
        fname = override.get("filename")
        sname = override.get("sheet_name")
        # Find matching uploaded file
        for fpath in os.listdir(session["dir"]):
            if fname in fpath:
                full = os.path.join(session["dir"], fpath)
                hdr = detect_header_row(full, sname, max_scan=10)
                session["files"][canonical] = {
                    "file": full,
                    "source_sheet": sname,
                    "header_row": hdr,
                    "original_filename": fname,
                    "original_sheet": sname,
                    "confidence": 999,
                }
                break
    return jsonify({"ok": True, "sheets": list(session["files"].keys())})

@app.route("/api/template/status", methods=["GET"])
def template_status():
    """Check if template file exists."""
    templates = list(TEMPLATE_DIR.glob("*.xlsx"))
    if templates:
        return jsonify({"found": True, "name": templates[0].name})
    return jsonify({"found": False})

@app.route("/api/session/<sid>/process", methods=["POST"])
def process(sid):
    """Run the full report generation."""
    session = get_session(sid)
    body = request.get_json() or {}

    date_start = body.get("date_start", "")
    date_end   = body.get("date_end", "")

    # Parse dates
    def parse_dt(s):
        for fmt in ["%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"]:
            try: return datetime.strptime(s.strip(), fmt)
            except: pass
        raise ValueError(f"Cannot parse date: {s}")

    try:
        dt_start = parse_dt(date_start)
        dt_end   = parse_dt(date_end) if date_end else dt_start
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if dt_start > dt_end:
        dt_start, dt_end = dt_end, dt_start

    # Generate date list
    dates = []
    cur = dt_start
    while cur <= dt_end:
        dates.append(cur)
        cur += timedelta(days=1)

    # Check template
    templates = list(TEMPLATE_DIR.glob("*.xlsx"))
    if not templates:
        return jsonify({"error": "No template file found in template/ folder"}), 500
    template_path = str(templates[0])

    # Load data
    try:
        data, load_errors = load_raw_data_from_session(session["dir"], session["files"])
    except Exception as e:
        return jsonify({"error": f"Data load failed: {e}", "trace": traceback.format_exc()}), 500

    if load_errors:
        # Non-fatal warnings
        pass

    # Find columns and compute metrics
    with zipfile.ZipFile(template_path, 'r') as z:
        s1_xml = z.read("xl/worksheets/sheet1.xml")

    col_metrics_map = {}
    skipped = []
    warnings = list(load_errors)

    for td in dates:
        tc = find_target_col(s1_xml, td)
        if tc is None:
            skipped.append(td.strftime("%m/%d/%Y"))
            continue
        try:
            metrics = compute_metrics(data, td)
            col_metrics_map[tc] = metrics
        except Exception as e:
            warnings.append(f"{td:%m/%d/%Y}: metric error – {e}")

    if not col_metrics_map:
        return jsonify({
            "error": "None of the requested dates were found in the Summary template.",
            "skipped": skipped,
        }), 400

    # Build output
    latest_date = max(dates)
    out_filename = f"Report_{dt_start:%m%d%Y}"
    if dt_end != dt_start:
        out_filename += f"_{dt_end:%m%d%Y}"
    out_filename += f"_{uuid.uuid4().hex[:6]}.xlsx"
    out_path = os.path.join(session["dir"], out_filename)

    try:
        build_output(template_path, out_path, col_metrics_map, data, latest_date,
                     date_start=dt_start, date_end=dt_end)
    except Exception as e:
        return jsonify({"error": f"Build failed: {e}", "trace": traceback.format_exc()}), 500

    return jsonify({
        "ok": True,
        "output_file": out_filename,
        "download_url": f"/api/session/{sid}/download/{out_filename}",
        "dates_processed": len(col_metrics_map),
        "dates_skipped": skipped,
        "warnings": warnings,
    })

@app.route("/api/session/<sid>/download/<filename>")
def download(sid, filename):
    session = get_session(sid)
    path = os.path.join(session["dir"], filename)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/session/<sid>/reset", methods=["POST"])
def reset_session(sid):
    with sessions_lock:
        if sid in sessions:
            try: shutil.rmtree(sessions[sid]["dir"])
            except: pass
            del sessions[sid]
    get_session(sid)
    return jsonify({"ok": True})


if __name__ == "__main__":
    print("=" * 60)
    print("  Excel Report Automation – Web Server")
    print(f"  Template dir : {TEMPLATE_DIR}")
    print(f"  Sessions dir : {SESSIONS_DIR}")
    print("  Place your Summary template .xlsx in the template/ folder")
    print("=" * 60)
    app.run(debug=False, host="0.0.0.0", port=5050, threaded=True)
