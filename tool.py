# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'e:\chengxu\python\ppthelper\newui.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

#Made By k4641321
from PyQt5 import QtCore, QtGui, QtWidgets
import psutil
import win32com.client,win32con,win32gui
import pyautogui,sys
#导入必要的库

office='screenClass'
Application = win32com.client.Dispatch('PowerPoint.Application')
page = Application.SlideShowWindows(1).View.Slide.SlideIndex

sw, sh = pyautogui.size()
x = sw/2-350
y = sh-100

#窗口主体

def get_pid(exeifrun):
    for proc in psutil.process_iter():
        if proc.name() == exeifrun:
            return proc.pid
    print("该进程不存在！")
    return None

class Ui_newui(object):
    def setupUi(self, newui):
        newui.setObjectName("newui")
        newui.resize(700, 70)
        newui.move(int(x),int(y))
        icon = QtGui.QIcon.fromTheme("\\main\\icon.ico")
        newui.setWindowIcon(icon)
        newui.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        newui.setFixedSize(newui.width(), newui.height())
        self.mouse = QtWidgets.QPushButton(newui)
        self.mouse.setGeometry(QtCore.QRect(0, 0, 70, 70))
        self.mouse.setObjectName("mouse")
        self.mouse.clicked.connect(self.sb)
        self.xp = QtWidgets.QPushButton(newui)
        self.xp.setGeometry(QtCore.QRect(70, 0, 70, 70))
        self.xp.setObjectName("xp")
        self.xp.clicked.connect(self.xpc)
        self.pen = QtWidgets.QPushButton(newui)
        self.pen.setGeometry(QtCore.QRect(140, 0, 70, 70))
        self.pen.setObjectName("pen")
        self.pen.clicked.connect(self.pens)
        self.esc = QtWidgets.QPushButton(newui)
        self.esc.setGeometry(QtCore.QRect(210, 0, 70, 70))
        self.esc.setObjectName("esc")
        self.esc.clicked.connect(self.exit)
        self.pagedown = QtWidgets.QPushButton(newui)
        self.pagedown.setGeometry(QtCore.QRect(420, 0, 70, 70))
        self.pagedown.setObjectName("pagedown")
        self.pagedown.clicked.connect(self.xfy)
        self.pageup = QtWidgets.QPushButton(newui)
        self.pageup.setGeometry(QtCore.QRect(350, 0, 70, 70))
        self.pageup.setObjectName("pageup")
        self.pageup.clicked.connect(self.sfy)
        self.pages = QtWidgets.QLabel(newui)
        self.pages.setGeometry(QtCore.QRect(280, 0, 70, 70))
        self.pages.setObjectName("pages")
        self.blackscreen = QtWidgets.QPushButton(newui)
        self.blackscreen.setGeometry(QtCore.QRect(490, 0, 70, 70))
        self.blackscreen.setObjectName("blackscreen")
        self.blackscreen.clicked.connect(self.hp)
        self.whitescreen = QtWidgets.QPushButton(newui)
        self.whitescreen.setGeometry(QtCore.QRect(560, 0, 70, 70))
        self.whitescreen.setObjectName("whitescreen")
        self.whitescreen.clicked.connect(self.bp)
        self.runing = QtWidgets.QPushButton(newui)
        self.runing.setGeometry(QtCore.QRect(630, 0, 70, 70))
        self.runing.setObjectName("runing")
        self.runing.clicked.connect(self.yx)
        self.retranslateUi(newui)
        QtCore.QMetaObject.connectSlotsByName(newui)

    def retranslateUi(self, newui):
        _translate = QtCore.QCoreApplication.translate
        newui.setWindowTitle(_translate("newui", "PPThelper-by k4641321"))
        self.mouse.setText(_translate("newui", "鼠标"))
        self.xp.setText(_translate("newui", "橡皮"))
        self.pen.setText(_translate("newui", "笔"))
        self.esc.setText(_translate("newui", "结束放映"))
        self.pagedown.setText(_translate("newui", "下翻页"))
        self.pageup.setText(_translate("newui", "上翻页"))
        self.pages.setText(_translate("newui", '页码：'+str(page)))
        self.blackscreen.setText(_translate("newui","黑屏"))
        self.whitescreen.setText(_translate("newui","白屏"))
        self.runing.setText(_translate("newui","恢复PPT"))

    #重新运行PPT
    def yx(self):
        Application.SlideShowWindows(1).View.State = 1

    #黑屏函数
    def hp(self):
        Application.SlideShowWindows(1).View.State = 3

    #白屏函数
    def bp(self):
        Application.SlideShowWindows(1).View.State = 4

    #下翻页函数
    def xfy(self):
            _translate = QtCore.QCoreApplication.translate
            Application.SlideShowWindows.Item(1).View.Next()
            global page 
            page = Application.SlideShowWindows(1).View.Slide.SlideIndex
            self.pages.setText(_translate("newui", '页码：'+str(page)))
            app.processEvents()


    #上翻页函数
    def sfy(self):
         _translate = QtCore.QCoreApplication.translate
         a = Application.SlideShowWindows(1).View.Slide.SlideIndex
         if a==1:
            print('已是第一页')
         else:
            Application.SlideShowWindows.Item(1).View.GotoSlide(a-1)
            global page
            page= Application.SlideShowWindows(1).View.Slide.SlideIndex
            self.pages.setText(_translate("newui", '页码：'+str(page)))
            app.processEvents()

    #切换笔函数
    def pens(self):
            officeppt = win32gui.FindWindow(office,None)
            win32gui.ShowWindow(officeppt,win32con.SW_SHOW)
            win32gui.SetForegroundWindow(officeppt)
            pyautogui.hotkey('ctrl', 'p')

    #结束放映函数
    def exit(self):
        Application.SlideShowWindows.Item(1).View.Exit()

    #切换橡皮函数
    def xpc(self):
            officeppt = win32gui.FindWindow(office,None)
            win32gui.ShowWindow(officeppt,win32con.SW_SHOW)
            win32gui.SetForegroundWindow(officeppt)
            pyautogui.hotkey('ctrl', 'e')

    #切换鼠标函数
    def sb(self):
            officeppt = win32gui.FindWindow(office,None)
            win32gui.ShowWindow(officeppt,win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(officeppt)
            pyautogui.hotkey('ctrl', 'a')

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    newui = QtWidgets.QWidget()
    ui = Ui_newui()
    ui.setupUi(newui)
    newui.show()
    sys.exit(app.exec_())
