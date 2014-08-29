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
        minute_counter = 0
        update_time = 1     # NOTE Change for Release version

        t_min_old = datetime.datetime.now().minute
        computer = os.environ['COMPUTERNAME']
        SHA = '0'
        dow_SHA = '0'

        reachable = False

        urls_path = "C:/Agent/urls.dat"
        online_path = "C:/Agent/online.dat"

        old_path = "C:/Agent/code/AgentCode.py"
        exec_path = "C:/Agent/code/AgentCode.py"
        current_path = "C:/Agent/code/AgentCode.py"
        new_path = ''

        getAddr = 'http://mikkel.pythonanywhere.com/static/AgentScript.py'
        shaAddr = 'http://mikkel.pythonanywhere.com/static/sha.txt'

        code_error = 0
        code_err_size = 0

        error = 0

        downgrade_err = False
        run_status = 0
        run_min = 0
        run_min_span = 1
        run_exec = 0
        run_exec_span = 3
        new_code = False
        current_code = False

        if os.path.exists(online_path):
            with open(online_path) as o:
                for line in o:
                    online_url = line[:-1]
        else:
            online_url = ''

        #fil = open("C:/Users/Michal/Desktop/agent.txt",'a')

        while 1:
            # Wait for service stop signal, if I timeout, loop again
            rc = win32event.WaitForSingleObject(self.hWaitStop, self.timeout)
            # Check to see if self.hWaitStop happened
            if rc == win32event.WAIT_OBJECT_0:
            # Stop signal encountered
                servicemanager.LogInfoMsg("Windar_Agent - STOPPED!")  #For Event Log
                break
            else:
                #[actual service code between rests]

                print "alive"
                t_min_new = datetime.datetime.now().minute
                if (t_min_new != t_min_old):
                    t_min_old = t_min_new
                    minute_counter += 1

                status = {'pc': computer[7:], 'upd_time': update_time, 'error': error, 'code_err': code_error}

                if not reachable:
                    try:
                        requests.get(online_url+'/monitor/online/', timeout=3)
                    except:
                        print "get back reaching is not reachable/tries from file"
                        with open(urls_path) as f:
                            for line in f:
                                if not reachable:
                                    try:
                                        print line[:-1]+'/monitor/online/'
                                        requests.get(line[:-1]+'/monitor/online/', timeout=3)
                                    except:
                                        print line[:-1]+" NOT REACHABLE"
                                        reachable = False
                                    else:
                                        print line[:-1]+" REACHABLE"
                                        online_url = line[:-1]
                                        try:
                                            os.remove(online_path)
                                        except:
                                            pass
                                        finally:
                                            with open (online_path, "w+") as online:
                                                online.write(line)
                                        reachable = True
                    else:
                        print "get_back reachable/try to post POST params"
                        reachable = True
                        try:
                            r = requests.post(online_url+"/monitor/receive/status/"+computer, params=status)
                        except:
                            reachable = False
                            print "unreachable"
                            print "save to statistics about connection"
                            pass
                        else:
                            #print r.content
                            reachable = True
                            print "reachable - save to statistics about connection"

                else:
                    try:
                        #requests.post("http://172.16.3.62:8000/monitor/receive/", params=status)
                        r = requests.post(online_url+"/monitor/receive/status/"+computer, params=status)
                    except:
                        reachable = False
                        print "unreachable again"
                        print "save to statistics about connection"
                        pass
                    else:
                        #print r.content
                        reachable = True
                        print "reachable - save to statistics about connection"

                if (minute_counter > (update_time-1)) and (reachable):
                    minute_counter = 0

                    try:
                        e = os.path.getsize("C:/Agent/error.log")
                    except:
                        pass
                    else:
                        if e > err_size:
                            code_error = 1
                            err_size = e
                        else:
                            code_error = 0

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
                                    os.remove(new_path)
                                    error = 1
                                elif current_code:
                                    print "Roll back to old_path/FIRST UPDATE OR UNUSUAL BEHAVIOUR"
                                    #roll back to old code
                                    current_code = False
                                    exec_path = old_path
                                    error = 2
                                else:
                                    print "downgrade_err = True"
                                    #old code doesnt work
                                    downgrade_err = True
                                    error = 3

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
                            print "Code finished successfully or waiting to new code after after downgrade error"

                            if new_code:
                                #rewrite path to current
                                print "New code so rewrite to current path"
                                if current_path != old_path:
                                    os.remove(current_path)
                                current_path = new_path
                                current_code = True
                                new_path = None
                                new_code = False
                                exec_path = current_path
                                error = 0
                            else:
                                #current working code / check for updates
                                print "Current working code so try to dow SHA if new is available:"
                                #Download:
                                try:
                                    r = requests.get(shaAddr, auth=HTTPBasicAuth("majkl","majkl"), timeout=5)  #REMOVE THIS LINE
                                    #r = requests.get(online_url+"/monitor/sha/"+computer[7:]", auth=HTTPBasicAuth("majkl","majkl")) NOTE!!! UPDATE with this
                                except:
                                    error = 11
                                    pass
                                else:
                                    if r.status_code == 200:
                                        dow_SHA = r.content
                                        print "SHA fetched: "+dow_SHA

                                if (dow_SHA != SHA):
                                    print "Download SHA Successful and SHAs different"
                                    tryouts = 0
                                    cal_SHA = ''

                                    while (dow_SHA != cal_SHA) and (tryouts < 5):
                                        print "Try to download new code"
                                        try:
                                            r = requests.get(getAddr, auth=HTTPBasicAuth("majkl","majkl"), timeout=5)  #REMOVE THIS LINE
                                            #r = requests.get(online_url+"/static/"+computer[7:]+"AgentCode.py", auth=HTTPBasicAuth("majkl","majkl")) NOTE!!! UPDATE with this
                                        except:
                                            error = 12
                                            pass
                                        else:
                                            if r.status_code == 200:
                                                new_path = "C:/Agent/code/AgentCode"+time.strftime('-%Y-%m-%d-%H-%M')+".py"
                                                with open(new_path, "wb") as code:
                                                    code.write(r.content)
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
                                        error = 20
                                    else:
                                        print "Download & SHA calculation Unsuccessful"
                                        os.remove(new_path)
                                        SHA = dow_SHA
                                        error = 13

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
                                os.remove(new_path)
                                new_code = False
                                exec_path = current_path
                                error = 4
                            elif current_code:
                                print "Roll back to old_path/FIRST UPDATE OR UNUSUAL BEHAVIOUR"
                                #roll back to old code
                                current_code = False
                                exec_path = old_path
                                error = 5
                            else:
                                print "downgrade_err = True"
                                #old code doesnt work
                                downgrade_err = True
                                error = 6

                    if (downgrade_err == False) and (run_min == 0):
                        print "Starting subprocess"
                        p = subprocess.Popen(["C:/Agent/Scripts/Python.exe", exec_path],stdout=subprocess.PIPE,)
                        run_exec += 1
                        run_min += 1
                #[actual service code between rests]

def ctrlHandler(ctrlType):
    return True

if __name__ == '__main__':   
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)   
    win32serviceutil.HandleCommandLine(aservice)
