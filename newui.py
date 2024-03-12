import win32com.client,win32con,win32api,win32gui
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPainter, QColor
import pyautogui

office0310='paneClassDC'
Application = win32com.client.Dispatch('PowerPoint.Application')
pages = Application.SlideShowWindows(0).View.Slide.SlideIndex

def xfy(self):
    Application.SlideShowWindows.Item(1).View.Next()
    global pages 
    pages = Application.SlideShowWindows(1).View.Slide.SlideIndex

def sfy(self):
    a = Application.SlideShowWindows(1).View.Slide.SlideIndex
    Application.SlideShowWindows.Item(1).View.GotoSlide(a-1)
    global pages 
    pages= Application.SlideShowWindows(1).View.Slide.SlideIndex
    
def exit(self):
    Application.SlideShowWindows.Item(1).View.Exit()

def xp(self):
    office0310hwnd = win32gui.FindWindow(office0310,None)
    win32gui.ShowWindow(office0310hwnd,win32con.SW_SHOW)
    win32gui.SetForegroundWindow(office0310hwnd)
    pyautogui.hotkey('ctrl', 'e')

def sb(self):
    office0310hwnd = win32gui.FindWindow(office0310,None)
    win32gui.ShowWindow(office0310hwnd,win32con.SW_SHOW)
    win32gui.SetForegroundWindow(office0310hwnd)
    pyautogui.hotkey('ctrl', 'a')

def pen(self):
    office0310hwnd = win32gui.FindWindow(office0310,None)
    win32gui.ShowWindow(office0310hwnd,win32con.SW_SHOW)
    win32gui.SetForegroundWindow(office0310hwnd)
    pyautogui.hotkey('ctrl', 'p')







