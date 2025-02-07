import pyautogui
import time
import os

class AutoClicker:
    def __init__(self):
        #滑鼠
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1  # 操作後暫停
        
    def click(self, x, y):
        """
        再截圖指定座標點擊
        
        :param x
        :param y
        """
        try:

            pyautogui.moveTo(x, y, duration=0.2)

            pyautogui.click()
        except Exception as e:
            print(f"點擊失敗ㄌ: {str(e)}")