# gunicorn.conf.py — Production config for ~1000 concurrent users
worker_class = "eventlet"
workers = 1
threads = 4
bind = "0.0.0.0:5000"
timeout = 30
keepalive = 5
graceful_timeout = 30
accesslog = "-"
errorlog  = "-"
loglevel  = "info"
access_log_format = '%(h)s "%(r)s" %(s)s %(b)s %(D)sms'
proc_name = "deeksharambh2026"
reload = False
max_requests = 1000
max_requests_jitter = 50
