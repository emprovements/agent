import os, sys, time
import requests
import psutil

while 1:
	cpu = psutil.cpu_percent(interval=1)
	ram = psutil.virtual_memory()
		
	payload = {'cpu': cpu, 'ram': ram.percent}
	print "sending"
	r = requests.post("http://192.168.1.57:8000/monitor/receive/", params=payload)
	time.sleep(2)

