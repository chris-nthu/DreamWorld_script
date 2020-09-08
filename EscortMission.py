import os.path
import win32gui
import time
import pyautogui
import numpy as np
import cv2
import pytesseract

from PyQt5.QtCore import QThread, pyqtSignal
from PIL import ImageGrab, Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Tesseract-OCR\tesseract.exe'
tessdata_dir_config = '--tessdata-dir "C://Tesseract-OCR//tessdata"'

class EscortMissionThread(QThread):
    trigger = pyqtSignal(str)

    def __init__(self, hwnd, action_thread_child, HeightDiff):
        super(EscortMissionThread, self).__init__()
        self.hwnd = hwnd
        self.action_thread_child = action_thread_child
        self.HeightDiff = HeightDiff
        self._isRunning = True
    
    def stop(self):
        self._isRunning = False
        self.action_thread_child.terminate()
    
    def run(self):
        time.sleep(0.1)

        while self._isRunning is True:
            left, top, right, bot = win32gui.GetWindowRect(self.hwnd)
            img = ImageGrab.grab(bbox=(left+330, top+350-self.HeightDiff, left+810, top+470-self.HeightDiff))
            img = np.array(img)
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            img = cv2.resize(img, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_LANCZOS4)
            ret, img = cv2.threshold(img, 170, 255, cv2.THRESH_BINARY)
            img = cv2.GaussianBlur(img, (5, 5), 0)
            ImageText = pytesseract.image_to_string(img, lang='chi_sim', config=tessdata_dir_config)
            win32gui.SetForegroundWindow(self.hwnd)

            if '我 这 就 去' in ImageText:
                self.trigger.emit('觸發羅盤事件')
                pyautogui.click(left+415, top+432-self.HeightDiff, button='left')
                time.sleep(1)
                pyautogui.click(left+415, top+432-self.HeightDiff, button='right')
                time.sleep(2)
                compass_position = pyautogui.locateCenterOnScreen('screenshot\EscortMission\\compass.png', confidence=.7)
                if compass_position is not None:
                    pyautogui.click(compass_position)
                    time.sleep(1)
                    pyautogui.click(compass_position)
                    time.sleep(1)
                    pyautogui.click(compass_position)
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue

            elif '返 程' in ImageText:
                self.trigger.emit('觸發返程事件')
                pyautogui.click(left+306, top+425-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue

            elif '继 续 运' in ImageText:
                self.trigger.emit('觸發繼續運鏢事件')
                pyautogui.click(left+416, top+434-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue
            
            elif '完 成 了 30 次' in ImageText:
                self.trigger.emit('您已經完成 30 環運鏢')
                self.action_thread_child.terminate()
                pyautogui.press('space')
                break

            time.sleep(3)