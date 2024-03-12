import win32com.client
import subprocess,time,sys

Application = win32com.client.Dispatch('PowerPoint.Application')
office='screenClass'

while True:
    type = Application.SlideShowWindows.Count
    if type == 1:
        subprocess.Popen('.\\main\\jctc.exe')
        subprocess.Popen('.\\main\\tool.exe')
        sys.exit()
    else:
        print("未放映")
        time.sleep(1)
