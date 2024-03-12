import win32com.client,win32con,win32gui
import subprocess

Application = win32com.client.Dispatch('PowerPoint.Application')
office='screenClass'
type = Application.SlideShowWindows.Count
print (type)