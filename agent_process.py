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
import datetime

from requests.auth import HTTPBasicAuth
import requests
import hashlib

import psutil



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
        computer = os.environ['COMPUTERNAME']
        old_SHA = '0'
        dow_SHA = '05ab4ec3db03f6f20910806a2d46dda094601c65f03169c32be41e7ca59c072d'
        dow_SHA_success = False

        postAddr = 'http://172.16.3.52:8000/monitor/receive/'+computer
        getAddr = 'http://mikkel.pythonanywhere.com/static/AgentScript.py'

        script_path = "C:/agent/AgentScript.py"

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
                    requests.post("http://172.16.3.73:8000/monitor/receive/", params=status)
                except:
                    print "unreachable"
                    print "save to statistics about connection"
                    reachable = False
                    pass
                else:
                    print "reachable - save to statistics about connection"
                    reachable = True
                if (t_min_new != t_min_old) and (reachable):
                    try:
                        run_status = p.poll()
                    except:
                        run_status = NOT_RUNNING
                    finally:
                        if downgrade_err == True:
                            run_status = 0
                    if run_status == None:
                    #subprocess is still running
                        if run_min > run_min_span:
                            TERMINATE
                            run_min = 0
                            if run_exec > run_exec_span:
                                run_exec = 0
                                run_min_span = 1
                                if new_code:
                                # Roll back to current path
                                    new_code = False
                                    exec_path = current_path
                                elif current_code:
                                #roll back to old code
                                    current_code = False
                                    exec_path = old_path
                                else:
                                #old code doesnt work
                                    downgrade_err = True
                            else:
                                run_exec += 1
                                run_min_span += 1
                        else:
                            run_min += 1
                    else:
                    #subprocess finished
                        run_min = 0
                        if run_status == 0:
                            #code finished successfully
                            if new_code:
                                #rewrite path to current
                                current_path = new_path
                                current_code = True
                                new_path = None
                                new_code = False
                                exec_path = current_path
                            else:
                            #current working code / check for updates
                            #ROUTINE TO FIND RIGHT SERVER
                            #Download:
                                print "Download dow_SHA"
	                            print "If download SHA successful !!! set dow_SHA_success flag !!!"
                                if (dow_SHA != SHA) and (dow_SHA_success):
	                                tryouts = 0
                                    cal_SHA = ''
                                    new_path = "C:/agent/AgentScript"+time.strftime('-%Y-%m-%d-%H-%M')+".py"
                                    while (dow_SHA != cal_SHA) and (tryouts < 5):
                                        print "Try to download new code"
                                        r = requests.get(getAddr, auth=HTTPBasicAuth("majkl","majkl"))
                                        with open(new_path, "wb") as code:
                                            code.write(r.content)
                                        if r.status_code == 200:
                                            print "calculation SHA"
                                            cal_SHA = hashfile(open(new_path, 'rb'), hashlib.sha256())
                                            print "Calculate SHA of new downloaded file: %s" % cal_SHA
                                        tryouts += 1
                                    if tryouts < 5:
                                        SHA = cal_SHA
                                        new_code = True
                                        downgrade_err = False
                                        run_exec = 0
                                        run_min_span = 1
                            #Download END
                        else:
                            #code finished with exception
                            run_exec = 0
                            run_min_span
                            #BLOCK OF CODE FROM PREVIOUS PART
			    
                #[actual service code between rests]

def ctrlHandler(ctrlType):
    return True

if __name__ == '__main__':   
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)   
    win32serviceutil.HandleCommandLine(aservice)
