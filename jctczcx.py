import psutil
import time 
import subprocess,sys,os

def get_pid(pname):
    for proc in psutil.process_iter():
        if proc.name() == pname:
            return proc.pid
    return None

while True :
    pid = get_pid('POWERPNT.EXE')
    if pid == None:
        subprocess.Popen('.\\main\\jczcx.exe')
        os.system('taskkill /f /im jcfy.exe')
        sys.exit()
    else:
        print('微软PPT已运行')
        time.sleep(1)
    
    