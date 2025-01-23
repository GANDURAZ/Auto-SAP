import subprocess
import pyautogui
import pyperclip
import time
import datetime
import pandas as pd
from pandas import ExcelWriter
from conf import *
import py_win_keyboard_layout

py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)

today = datetime.datetime.now().strftime('%d.%m.%Y')
rez = pd.read_excel('Rezult.xlsx')

# open SAP GUI
sap_logopn = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
subprocess.Popen(sap_logopn)
time.sleep(1)
pyautogui.press('enter')
time.sleep(10)
pyautogui.hotkey('win', 'up')
time.sleep(1)
pyautogui.click(550, 215)
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.click(250, 240)
time.sleep(1)
pyperclip.copy(Login)  # переменная, которая хранит ваш блок вставки 
time.sleep(1)  # задержка
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyperclip.copy(Password)  # переменная, которая хранит ваш блок вставки 
time.sleep(1)  # задержка
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
#Залогинелся

for i in range(len(SC)):
    pyautogui.click(200, 205)
    pyautogui.click(200, 205)
    time.sleep(1)
    pyautogui.click(30, 145)
    time.sleep(1)
    pyautogui.click(520, 470)
    time.sleep(1)
    if i > 0: 
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
    pyautogui.typewrite(SC[i])
    pyautogui.press('f8')
    time.sleep(1)
    pyautogui.click(20, 175)
    time.sleep(1)
    pyautogui.click(40, 200)
    time.sleep(1)
    pyautogui.click(100, 245)
    pyautogui.click(100, 245)
    time.sleep(1)
    pyautogui.click(450, 650)
    pyautogui.typewrite(today)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(today)
    time.sleep(1)
    pyautogui.press('f8')
    time.sleep(1)
    if i in [3, 14, 18, 20, 21, 25, 28, 29, 33]: pyautogui.click(710, 235)
    else:
        pyautogui.click(710, 275)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    pyautogui.press('esc')
    time.sleep(1)
    prov = pyperclip.paste()
    rez.loc[i, 'Время выгрузки данных'] = time.strftime("%H:%M:%S", time.localtime())
    if prov == '':
        continue
    else:
        time.sleep(1)
        pyautogui.click(140, 225)
        pyautogui.click(140, 225)
        time.sleep(1)
        pyautogui.click(40, 210)
        time.sleep(1)
        pyautogui.click(350, 320)
        time.sleep(1)
        pyautogui.typewrite(SC[i])
        pyautogui.click(810, 350)
        time.sleep(1)
        pyautogui.click(90, 330)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('f8')
        time.sleep(1)
        pyautogui.press('f8')
        time.sleep(3) #Скопирывать значения
        pyautogui.click(1200, 185)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        rez.loc[i, 'Не упакованные шт.'] = float(pyperclip.paste().replace('.', "").replace(' ', "").replace(',', "."))
        pyautogui.click(1250, 185)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        rez.loc[i, 'Не собранные заказы'] = float(pyperclip.paste().replace('.', "").replace(' ', "").replace(',', "."))
        pyautogui.click(1400, 185)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        rez.loc[i, 'Не собранных штук'] = float(pyperclip.paste().replace('.', "").replace(' ', "").replace(',', "."))
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(1)

name = today + '.xlsx'

with pd.ExcelWriter(name) as writer:
    rez.to_excel(writer, index=False)
