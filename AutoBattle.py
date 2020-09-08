import win32gui
import pyautogui
import time

from PyQt5.QtCore import QThread, pyqtSignal

class AutoBattleThread(QThread):
    trigger = pyqtSignal(str)

    def __init__(self, hwnd):
        super(AutoBattleThread, self).__init__()
        self.hwnd = hwnd
        self._isRunning = True

    def stop(self):
        self._isRunning = False

    def run(self):
        while self._isRunning is True:
            win32gui.SetForegroundWindow(self.hwnd)
            pyautogui.keyDown('ctrl')
            pyautogui.keyDown('a')
            pyautogui.keyUp('a')
            pyautogui.keyUp('ctrl')
            self.trigger.emit('事件觸發：發送 CTRL+A 以執行自動戰鬥')
            time.sleep(300)