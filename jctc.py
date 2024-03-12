import win32com.client
import subprocess,time,sys,os

Application = win32com.client.Dispatch('PowerPoint.Application')
office='screenClass'

while True:
    type = Application.SlideShowWindows.Count
    if type == 0:
        subprocess.Popen('.\\main\\jcfy.exe')
        os.system('taskkill /f /im tool.exe')
        sys.exit()
    else:
        print("已放映")
        time.sleep(1)
