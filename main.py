import sys
import os.path
import win32gui
import pyautogui

from PyQt5.QtWidgets import QApplication
from PyQt5.uic import loadUi
from PyQt5 import QtGui
from atexit import register
from script import Script

def winEnumHandler(hwnd, widget):
    if win32gui.IsWindowVisible(hwnd):
        if '新 梦 想 世 界 - ' in win32gui.GetWindowText(hwnd):
            widget.comboBox_window.addItem(win32gui.GetWindowText(hwnd))

def pushButton_start_clicked(widget):
    s = Script(widget)
    s.run()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    widget = loadUi(os.path.abspath(os.path.dirname(__file__))+'/ui/main.ui')
    widget.setWindowTitle('新夢想世界腳本 v4.1')
    widget.setWindowIcon(QtGui.QIcon(os.path.abspath(os.path.dirname(__file__)) + '/image/dream.ico'))

    win32gui.EnumWindows(winEnumHandler, widget)

    widget.pushButton_start.clicked.connect(lambda:pushButton_start_clicked(widget))

    widget.show()
    app.exec_()
    pyautogui.press('space')
