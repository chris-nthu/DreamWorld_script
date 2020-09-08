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

class TreasureMapMissonThread(QThread):
    trigger = pyqtSignal(str)

    def __init__(self, hwnd, action_thread_child, HeightDiff):
        super(TreasureMapMissonThread, self).__init__()
        self.hwnd = hwnd
        self.action_thread_child = action_thread_child
        self.HeightDiff = HeightDiff
        self._isRunning = True
    
    def stop(self):
        self._isRunning = False
        self.action_thread_child.terminate()
    
    def analyzeDialogBox(self):
        win32gui.SetForegroundWindow(self.hwnd)
        left, top, right, bot = win32gui.GetWindowRect(self.hwnd)
        img = ImageGrab.grab(bbox=(left+330, top+350-self.HeightDiff, left+810, top+470-self.HeightDiff))
        img = np.array(img)
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        img = cv2.resize(img, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_LANCZOS4)
        ret, img = cv2.threshold(img, 170, 255, cv2.THRESH_BINARY)
        img = cv2.GaussianBlur(img, (5, 5), 0)
        ImageText = pytesseract.image_to_string(img, lang='chi_sim', config=tessdata_dir_config)
        return ImageText

    def run(self):
        while self._isRunning is True:
            ImageText = self.analyzeDialogBox()
            left, top, right, bot = win32gui.GetWindowRect(self.hwnd)

            if '打 听 专 有 藏 宝 图 信 息' in ImageText:
                self.trigger.emit('事件觸發：接取專有寶圖任務')
                pyautogui.click(left+531, top+417-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue

            elif '前 往 捉 拿' in ImageText:
                self.trigger.emit('事件觸發：捉拿事件')
                pyautogui.click(left+420, top+416-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue
            
            elif '前 往' in ImageText and '药' not in ImageText:
                self.trigger.emit('事件觸發：探查事件')
                time.sleep(2)
                explore_position = pyautogui.locateCenterOnScreen('screenshot\TreasureMapMission\\explore.png', confidence=.8)
                if explore_position is not None:
                    pyautogui.click(explore_position)
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue

            elif '进 入 战 斗' in ImageText:
                self.trigger.emit('事件觸發：戰鬥事件')
                pyautogui.click(left+474, top+408-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue
            
            elif '继 续 宝 图 任 务' in ImageText:
                self.trigger.emit('事件觸發：繼續任務事件')
                #pyautogui.click(left+430, top+418, button='left')
                time.sleep(2)
                continue_position = pyautogui.locateCenterOnScreen('screenshot\TreasureMapMission\\continue.png', confidence=.7)
                if continue_position is not None:
                    pyautogui.click(continue_position)
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=2)
                continue
            
            elif '药' in ImageText and '购 买 草 药' not in ImageText:
                self.trigger.emit('事件觸發：買藥事件')
                pyautogui.click(left+427, top+418-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)

                while True:
                    ImageText = self.analyzeDialogBox()

                    if '购 买 草 药' in ImageText:
                        time.sleep(1)
                        pyautogui.click(left+418, top+400-self.HeightDiff, button='left')
                        time.sleep(2)
                        pyautogui.click(left+692, top+599-self.HeightDiff, button='left')
                        time.sleep(1)
                        pyautogui.click(left+692, top+549-self.HeightDiff, button='right')
                        time.sleep(2)
                        hunter_position = pyautogui.locateCenterOnScreen('screenshot\TreasureMapMission\\hunter.png', confidence=.5)
                        if hunter_position is not None:
                            pyautogui.click(hunter_position)
                            time.sleep(1)
                            pyautogui.click(hunter_position)
                        time.sleep(0.1)
                        pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                        break

                continue
            
            elif '夺 取 宝 图' in ImageText or '定 日' in ImageText or '夺 回 宝 图' in ImageText:
                self.trigger.emit('事件觸發：奪取寶圖事件')
                #pyautogui.click(left+420, top+402, button='left')
                time.sleep(2)
                seize_position = pyautogui.locateCenterOnScreen('screenshot\TreasureMapMission\\seize.png', confidence=.5)
                if seize_position is not None:
                    pyautogui.click(seize_position)
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue
            
            elif '给 予 物 品' in ImageText or '予 物' in ImageText:
                self.trigger.emit('事件觸發：給予物品事件')
                pyautogui.click(left+421, top+399-self.HeightDiff, button='left')
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue
            
            elif '装 备' in ImageText and '前' in ImageText:
                self.trigger.emit('事件觸發：買裝事件')
                pyautogui.click(left+416, top+370-self.HeightDiff, button='right')
                time.sleep(3)
                pyautogui.keyDown('alt')
                pyautogui.keyDown('r')
                pyautogui.keyUp('r')
                pyautogui.keyUp('alt')
                time.sleep(3)
                pyautogui.moveTo(left+285, top+196-self.HeightDiff, 1)
                time.sleep(0.1)
                pyautogui.click(left+285, top+196-self.HeightDiff, button='left')
                time.sleep(3)
                pyautogui.click(left+591, top+640-self.HeightDiff, button='left')
                time.sleep(3)
                pyautogui.click(left+591, top+590-self.HeightDiff, button='right')
                time.sleep(1)
                pyautogui.click(left+591, top+590-self.HeightDiff, button='right')
                time.sleep(3)
                hunter_position = pyautogui.locateCenterOnScreen('screenshot\TreasureMapMission\\hunter.png', confidence=.5)
                if hunter_position is not None:
                    pyautogui.click(hunter_position)
                    time.sleep(1)
                    pyautogui.click(hunter_position)
                time.sleep(0.1)
                pyautogui.moveTo(left+330, top+300-self.HeightDiff, duration=0.2)
                continue

            elif '20 张' in ImageText:
                self.trigger.emit('<b style="color:blue;">INFO：您已完成專有寶圖任務</b>') 
                self.action_thread_child.terminate()
                pyautogui.press('space')
                break
            
            elif '> 20' in ImageText:
                self.trigger.emit('<b style="color:red;">WARNING：背包中的寶圖超過 20 張</b>')
                self.action_thread_child.terminate()
                pyautogui.press('space')
                break

            time.sleep(3)