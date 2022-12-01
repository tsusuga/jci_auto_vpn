import pyautogui
import time
from dotenv import load_dotenv
import os
import webbrowser
import win32com.client


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

# time.sleep(20)

# # outlookからPINを抽出する処理
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# inbox = outlook.GetDefaultFolder(6)

# # フォルダを取得
# accounts = outlook.Folders('My account')
# folders = accounts.Folders('受信トレイ').Folders('AuthCode')
# mails = folders.Items

# idx = len(mails)

# mail = mails[idx - 1]
# subject = mail.subject
# pincode = subject[10:16]
# print(pincode)

# pyautogui.write(pincode)
# pyautogui.press('enter')