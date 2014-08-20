### Run Python scripts as a service example (ryrobes.com)
### Usage : python aservice.py install (or / then start, stop, remove)

import win32service
import win32serviceutil
import win32api
import win32con
import win32event
import win32evtlogutil
import os, sys, string, time

import subprocess
import psutil
import datetime

import requests
import psutil
import hashlib

from requests.auth import HTTPBasicAuth
import subprocess


class aservice(win32serviceutil.ServiceFramework):
	_svc_name_ = "MyServiceShortName"
	_svc_display_name_ = "My Serivce Long Fancy Name!"
	_svc_description_ = "THis is what my crazy little service does - aka a DESCRIPTION! WHoa!"

	def __init__(self, args):
		win32serviceutil.ServiceFramework.__init__(self, args)
		self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)           
 
	def SvcStop(self):
		self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
		win32event.SetEvent(self.hWaitStop) 
	
	def SvcDoRun(self):
		import servicemanager      
		servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,servicemanager.PYS_SERVICE_STARTED,(self._svc_name_, '')) 
		def hashfile(afile, hasher, blocksize=65536):
			buf = afile.read(blocksize)
			while len(buf) > 0:
				hasher.update(buf)
				buf = afile.read(blocksize)
			return hasher.hexdigest()

		

		#self.timeout = 640000    #640 seconds / 10 minutes (value is in milliseconds)
		#self.timeout = 120000     #120 seconds / 2 minutes
		self.timeout = 5000     
		# This is how long the service will wait to run / refresh itself (see script below)
		
		t_min_old = datetime.datetime.now().minute
		error = '0'
		updated = False
		computer = os.environ['COMPUTERNAME']
		python_path = sys.executable
		old_SHA = '0'
		dow_SHA = '05ab4ec3db03f6f20910806a2d46dda094601c65f03169c32be41e7ca59c072d'
		dow_SHA_success = False
		postAddr = 'http://172.16.3.52:8000/monitor/receive/'+computer
		getAddr = 'http://mikkel.pythonanywhere.com/static/AgentScript.py'

		script_path = "C:/agent/a.py"

		#fil = open("C:/Users/Michal/Desktop/agent.txt",'a')

		while 1:
			# Wait for service stop signal, if I timeout, loop again
			rc = win32event.WaitForSingleObject(self.hWaitStop, self.timeout)
			# Check to see if self.hWaitStop happened
			if rc == win32event.WAIT_OBJECT_0:
			# Stop signal encountered
				servicemanager.LogInfoMsg("SomeShortNameVersion - STOPPED!")  #For Event Log
				break
			else:
				#[actual service code between rests]
				print "alive"
				t_min_new = datetime.datetime.now().minute
				status = {'pc': computer, 'error': error}

				try:
					requests.post("http://192.168.1.57:8000/monitor/receive/", params=status)

					if (error == 5):
						error = 0

				except:
					print "unreachable"
					print "save to statistics about connection"
					reachable = False
					pass
				else:
					print "reachable - save to statistics about connection"
					reachable = True

				if (t_min_new != t_min_old):

					t_min_old = t_min_new
					print "Starting subprocess"
					p = subprocess.Popen(["C:/Users/Michal/.virtualenvs/agent/Scripts/Python.exe", script_path],stdout=subprocess.PIPE,)
	
					#stdout_value = p.communicate()
					


def ctrlHandler(ctrlType):
	return True

if __name__ == '__main__':   
	win32api.SetConsoleCtrlHandler(ctrlHandler, True)   
	win32serviceutil.HandleCommandLine(aservice)
