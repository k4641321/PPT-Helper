import psutil
import time 
import subprocess,sys

def get_pid(pname):
    for proc in psutil.process_iter():
        if proc.name() == pname:
            return proc.pid
    return None

while True :
    pid = get_pid('POWERPNT.EXE')
    if pid != None:
        subprocess.Popen('.\\main\\jcfy.exe')
        subprocess.Popen('.\\main\\jctczcx.exe')
        sys.exit()
    else:
        print('微软PPT未运行')
        time.sleep(1)
    
    