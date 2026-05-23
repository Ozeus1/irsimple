import multiprocessing
import os


bind = os.environ.get("IRSIMPLE_BIND", "127.0.0.1:5050")
workers = int(os.environ.get("IRSIMPLE_WORKERS", max(2, multiprocessing.cpu_count() // 2)))
threads = int(os.environ.get("IRSIMPLE_THREADS", "2"))
timeout = int(os.environ.get("IRSIMPLE_TIMEOUT", "180"))
accesslog = "-"
errorlog = "-"
capture_output = True
