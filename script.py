import win32gui
import threading
import pyautogui

from pynput import keyboard
from AutoBattle import AutoBattleThread
from TreasureMapMission import TreasureMapMissonThread
from EscortMission import EscortMissionThread

class Script():
    def __init__(self, widget):
        self.widget = widget
        self.windowTitle = widget.comboBox_window.currentText()
        self.action = widget.comboBox_action.currentText()
        self.hwnd = 0
        self.HeightDiff = 0
    
    def calHeightDiff(self):
        windowRec = win32gui.GetWindowRect(self.hwnd)
        clientRec = win32gui.GetClientRect(self.hwnd)
        MyWindowTitleHeight = (windowRec[3]-windowRec[1])-(clientRec[3]-clientRec[1])

        return 40 - MyWindowTitleHeight

    def MessageDisplay(self, str):
        self.widget.textBrowser_info.append(str)

    def on_press(self, key):
        if key == keyboard.Key.space:
            self.action_thread.stop()
            self.widget.pushButton_start.setEnabled(True)
            self.widget.pushButton_start.setText('啟 用 腳 本')
            self.MessageDisplay('<b style="color:blue;">INFO：暫停腳本</b>')
            self.MessageDisplay('')
            return False

    def monitorKey(self):
        with keyboard.Listener(on_press=self.on_press) as listener:
            listener.join()

    def run(self):
        if self.windowTitle == '':
            self.MessageDisplay('<b style="color:red;">ERROR：捕捉窗口失敗，檢查您是否已開啟遊戲</b>')
            return False
        
        # Get the game window
        self.hwnd = win32gui.FindWindow(0, self.windowTitle)
        self.widget.pushButton_start.setEnabled(False)
        self.widget.pushButton_start.setText('腳本執行中(欲停止請按空白鍵)')
        self.MessageDisplay('<b style="color:blue;">INFO：啟動腳本</b>')
        self.HeightDiff = self.calHeightDiff()

        # Monitor the event of stoping script
        threading.Thread(target=self.monitorKey).start()
        
        self.MessageDisplay('視窗名稱：' + self.windowTitle)
        self.MessageDisplay('啟用功能：' + self.action)

        if self.action == '靜修之心':
            self.MessageDisplay('功能說明：每 5 分鐘觸發 CTRL+A')
            self.action_thread = AutoBattleThread(self.hwnd)
            self.action_thread.start()
            self.action_thread.trigger.connect(self.MessageDisplay)

        elif self.action == '專有寶圖任務':
            self.MessageDisplay('功能說明：自動跑專有寶圖任務')
            action_thread_child = AutoBattleThread(self.hwnd)
            action_thread_child.start()
            self.action_thread = TreasureMapMissonThread(self.hwnd, action_thread_child, self.HeightDiff)
            self.action_thread.start()
            action_thread_child.trigger.connect(self.MessageDisplay)
            self.action_thread.trigger.connect(self.MessageDisplay)

        elif self.action == '運鏢任務':
            self.MessageDisplay('功能說明：自動跑標')
            action_thread_child = AutoBattleThread(self.hwnd)
            action_thread_child.start()
            self.action_thread = EscortMissionThread(self.hwnd, action_thread_child, self.HeightDiff)
            self.action_thread.start()
            action_thread_child.trigger.connect(self.MessageDisplay)
            self.action_thread.trigger.connect(self.MessageDisplay)
        
        elif self.action == '廚神任務':
            self.MessageDisplay('<b style="color:red;">ERROR：尚未完成</b>')
            self.action_thread = AutoBattleThread(self.hwnd)
            self.action_thread.start()
            pyautogui.press('space')