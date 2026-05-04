import multiprocessing

workers = multiprocessing.cpu_count() * 2 + 1
worker_class = "sync"
threads = 2
timeout = 120
bind = "0.0.0.0:10000"
max_requests = 1000
max_requests_jitter = 100
