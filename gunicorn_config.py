# Gunicorn configuration file
import multiprocessing
import os

# Server socket
bind = f"0.0.0.0:{int(os.getenv('PORT', 8000))}"
backlog = 2048

# Worker processes
workers = multiprocessing.cpu_count() * 2 + 1
worker_class = 'uvicorn.workers.UvicornWorker'
worker_connections = 1000
timeout = 30
keepalive = 2

# Logging
accesslog = '-'
errorlog = '-'
loglevel = 'info'