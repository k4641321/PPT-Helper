# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'e:\chengxu\python\ppthelperwithmsoffice\root.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
import os
import subprocess

subprocess.Popen('.\\main\\jczcx.exe')

class Ui_root(object):
    def setupUi(self, root):
        root.setObjectName("root")
        root.resize(350, 200)
        icon = QtGui.QIcon.fromTheme(".\\main\\icon.ico")
        root.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(root)
        self.label.setGeometry(QtCore.QRect(0, 0, 351, 151))
        self.label.setObjectName("label")
        self.exit = QtWidgets.QPushButton(root)
        self.exit.setGeometry(QtCore.QRect(210, 150, 141, 51))
        self.exit.setObjectName("exit")
        self.exit.clicked.connect(self.tc)
        self.label_2 = QtWidgets.QLabel(root)
        self.label_2.setGeometry(QtCore.QRect(0, 150, 211, 51))
        self.label_2.setObjectName("label_2")
        self.retranslateUi(root)
        QtCore.QMetaObject.connectSlotsByName(root)

    def retranslateUi(self, root):
        _translate = QtCore.QCoreApplication.translate
        root.setWindowTitle(_translate("root", "PPT Helper-by k4641321"))
        self.label.setText(_translate("root", "<html><head/><body><p align=\"center\"><span style=\" font-size:12pt;\">PPT Helper PPT帮手</span></p><p align=\"center\"><span style=\" font-size:11pt;\">此程序为创新开发参赛作品</span></p><p align=\"center\"><span style=\" font-size:11pt;\">所有内容均独创</span></p><p align=\"center\"><span style=\" font-size:11pt;\">设计者：龙岩四中 高一（2）班 林凯忻 </span></p><p align=\"center\"><span style=\" font-size:11pt;\">指导老师：杨俊</span></p></body></html>"))
        self.exit.setText(_translate("root", "退出程序"))
        self.label_2.setText(_translate("root", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"> 退出程序最好用右边的按钮    →→→</p>\n"
"<p style=\" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"> 以便彻底退出程序            →→→</p></body></html>"))
        
    def tc(self):
        os.system('taskkill /f /im jcfy.exe')
        os.system('taskkill /f /im jczcx.exe')
        os.system('taskkill /f /im jctc.exe')
        os.system('taskkill /f /im tool.exe')
        sys.exit()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    root = QtWidgets.QWidget()
    ui = Ui_root()
    ui.setupUi(root)
    root.show()
    sys.exit(app.exec_())