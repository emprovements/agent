import os, sys, time
import requests
import psutil

cpu = psutil.cpu_percent(interval=1)
ram = psutil.virtual_memory()

payload = {'cpu': cpu, 'ram': ram.percent}
print "sending"
r = requests.post("http://172.16.3.52:8000/monitor/receive/", params=payload)
