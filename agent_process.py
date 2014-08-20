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
    _svc_name_ = "Windar_Agent"
    _svc_display_name_ = "Windar_Agent"
    _svc_description_ = "Windar Agent to monitor system and send statistics"

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
        SHA = '0'
        dow_SHA = '05ab4ec3db03f6f20910806a2d46dda094601c65f03169c32be41e7ca59c072d'
        dow_SHA_success = True

        reachable = False

        old_path = "C:/agent/AgentScript.py"
        exec_path = "C:/agent/AgentScript.py"
        current_path = "C:/agent/AgentScript.py"
        new_path = ''

        postAddr = 'http://172.16.3.62:8000/monitor/receive/'+computer
        getAddr = 'http://mikkel.pythonanywhere.com/static/AgentScript.py'

        script_path = "C:/agent/AgentScript.py"

        error = 0

        downgrade_err = False
        run_status = 0
        run_min = 0
        run_min_span = 1
        run_exec = 0
        run_exec_span = 3
        new_code = False
        current_code = False



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
                    requests.post("http://172.16.3.62:8000/monitor/receive/", params=status)
                except:
                    #ROUTINE TO FIND RIGHT SERVER=============================================<+
                    reachable = False
                    print "unreachable"
                    print "save to statistics about connection"
                    pass
                else:
                    reachable = True
                    print "reachable - save to statistics about connection"
                if (t_min_new != t_min_old) and (reachable):
                    t_min_old = t_min_new
                    try:
                        run_status = p.poll()
                    except:
                        run_status = 0
                        print "NOT RUNNING / 1st round"
                    finally:
                        print "Subprocess STATUS: %s" % run_status
                        if downgrade_err == True:
                            print "Downgrade_err TRUE"
                            run_status = 0
                    if run_status == None:
                        #subprocess is still running
                        print "Subprocess still running"
                        if run_min > run_min_span:
                            print "Kill process"
                            try:
                                p.kill()
                            except:
                                pass
                                print "No Process to kill"
                            finally:
                                run_min = 0
                            if run_exec > run_exec_span:
                                run_exec = 0
                                run_min_span = 1
                                if new_code:
                                    print "Roll back to current_path"
                                    # Roll back to current path
                                    new_code = False
                                    exec_path = current_path
                                elif current_code:
                                    print "Roll back to old_path"
                                    #roll back to old code
                                    current_code = False
                                    exec_path = old_path
                                else:
                                    print "downgrade_err = True"
                                    #old code doesnt work
                                    downgrade_err = True
                            else:
                                print "Extending running time"
                                run_exec += 1
                                run_min_span += 1
                        else:
                            run_min += 1
                    else:
                        #subprocess finished
                        print "Subprocess finished"
                        run_min = 0
                        if run_status == 0:
                            #code finished successfully
                            print "Code finished successfully or waiting to new code"
                            if new_code:
                                #rewrite path to current
                                print "New code so rewrite to current path"
                                current_path = new_path
                                current_code = True
                                new_path = None
                                new_code = False
                                exec_path = current_path
                            else:
                                #current working code / check for updates
                                print "Current working code so try:"
                                #Download:
                                print "Download dow_SHA"    	 
                                print "If download SHA successful !!! set dow_SHA_success flag !!!"
                                print "SHA %s" % SHA
                                if (dow_SHA != SHA) and (dow_SHA_success):
                                    print "Download SHA Successful and SHAs different"
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
                                        print "SHA MATCH!!!"
                                        SHA = cal_SHA
                                        new_code = True
                                        exec_path = new_path
                                        downgrade_err = False
                                        run_exec = 0
                                        run_min_span = 1
                                    else:
                                        print "Download & SHA Unsuccessful"
                                #Download END
                                else:
                                    print "Keep current code / no new code available"
                        else:
                            #code finished with exception
                            run_exec = 0
                            run_min_span = 1
                            #BLOCK OF CODE FROM PREVIOUS PART
                            if new_code:
                                print "Roll back to current_path"
                                # Roll back to current path
                                new_code = False
                                exec_path = current_path
                            elif current_code:
                                print "Roll back to old_path"
                                #roll back to old code
                                current_code = False
                                exec_path = old_path
                            else:
                                print "downgrade_err = True"
                                #old code doesnt work
                                downgrade_err = True

                    if (downgrade_err == False) and (run_min == 0):
                        print "Starting subprocess"
                        p = subprocess.Popen(["C:/Users/Michal/.virtualenvs/agent/Scripts/Python.exe", exec_path],stdout=subprocess.PIPE,)
                        run_exec += 1
                        run_min += 1
                #[actual service code between rests]

def ctrlHandler(ctrlType):
    return True

if __name__ == '__main__':   
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)   
    win32serviceutil.HandleCommandLine(aservice)
