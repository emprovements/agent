import os, sys, time
import requests
import psutil

cpu = putil.cpu_percent(interval=1)
ram = psutil.virtual_memory()
mport psutl
payload = {'cpu': cpu, 'ram': ram.percent}
print "sending"
r = requests.post("http://172.16.3.73:8000/monitor/receive/", params=payload)
