# Excel Report Automation – Web App

## Folder Structure

```
excel_webapp/
├── backend/
│   └── app.py              # Flask backend
├── frontend/
│   └── index.html          # Web UI
├── template/
│   └── <your_template>.xlsx  ← PUT YOUR SUMMARY TEMPLATE HERE
├── sessions/               # Auto-created, temp user files
└── requirements.txt
```

## Setup

1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Place your Summary template**
   Copy your Summary Excel template (with the pre-built Summary sheet) into the `template/` folder.
   The app will always use it as the base — users never need to upload it.

3. **Start the server**
   ```bash
   python backend/app.py
   ```
   Open http://localhost:5050 in your browser.

## Multiple Users

The app is session-based — each browser tab gets a unique session ID stored in `sessionStorage`.
All uploaded files are isolated to that session's folder under `sessions/`.
Sessions auto-expire after 2 hours.

## How Sheet Detection Works

Users can upload files with any sheet name. The app identifies each sheet by scanning its columns
and matching against known signatures:

| Target Sheet    | Key Columns Checked                              |
|-----------------|--------------------------------------------------|
| Bank_Statement_R| Date, Amount, Status                             |
| Pay_06_R        | textbox16, Check_Amt                             |
| Pay_41_R        | Trans_Date, Payment, Financial_Class             |
| Age_24_R        | Financial_Class, textbox18                       |
| Fin_25_R        | Textbox2, Total_Charge                           |
| Waystar_R       | Trans Date, Charges, Claim ID                    |
| Rejected        | Trans Date, Charges, Last Event                  |
| Cnt_27          | Svc_Date, Withhold_Code                          |

## Usage Modes

- **Single combined file**: Upload one .xlsx with all raw sheets → all detected automatically
- **Separate files**: Upload 1-8 individual files, each with one sheet → each detected independently
- **Mixed**: Upload some combined, some separate — duplicates are resolved by confidence score

## Date Range

Use the date range picker to process multiple days at once:
- All calendar days from start → end are processed
- Each day gets its own column updated in the Summary
- REJECTIONS and KPI sheets reflect the latest date in the range
