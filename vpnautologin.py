import pyautogui
import time
from dotenv import load_dotenv


# winボタンプッシュ
pyautogui.press('win')

time.sleep(1)

# FortiClientを検索
pyautogui.write('forti')
pyautogui.press('enter')

time.sleep(4)

# FotiClientにパスワード入力
# pyautogui.click(x=976, y=694)
for i in range(3):
  pyautogui.press('tab')

load_dotenv()
PASSWORD = os.getenv('PASSWORD')
pyautogui.write(PASSWORD)
pyautogui.press('enter')

# outlookを表示
webbrowser.open("https://outlook.office.com/mail/")