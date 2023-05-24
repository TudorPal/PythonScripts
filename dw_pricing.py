import time
import pandas as pd
import openpyxl
import keyboard
import pyautogui

xls = pd.ExcelFile('Mesa.xlsx')

sheet = "Size-Glass Rect,Continuous Base"
style = "MSCF01GLMCB"

full = pd.read_excel('Mesa.xlsx', sheet_name=sheet)
df = full[full['Style'] == style]

price = df['Price']
description = df['Description']

size = len(df)
start = description.index.values[0]
stop = start + size
x_plus = 1810

x_desc = 960
y_desc = 225

x_price = 1240
y_price = 340

x_save = 730
y_save = 420
c = int(input("1 to start the program: "))
while c == 1:
    try:
        y_plus = 545
        for i in range(start, start + 6):
            pyautogui.moveTo(x_plus, y_plus, duration=1.5)
            pyautogui.click()
            pyautogui.moveTo(x_desc, y_desc, duration=0.7)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(description[i])
            pyautogui.moveTo(x_price, y_price, duration=0.05)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(str(price[i]))
            pyautogui.moveTo(x_save, y_save, duration=0.05)
            pyautogui.click()
            y_plus += 61
        y_plus -= 61
        for i in range(start + 6, stop):
            pyautogui.moveTo(x_plus, y_plus, duration=1.6)
            if i % 3 == 0:
                pyautogui.scroll(-72)
            else:
                pyautogui.scroll(-74)
            time.sleep(0.2)
            pyautogui.click()
            pyautogui.moveTo(x_desc, y_desc, duration=0.7)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(description[i])
            pyautogui.moveTo(x_price, y_price, duration=0.05)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(str(price[i]))
            pyautogui.moveTo(x_save, y_save, duration=0.05)
            pyautogui.click()
        c = 0
    except:
        print("program stopped")
        c = 0
        i = 0
        c = int(input("1 to restart the program: "))

# pyautogui.moveTo(x_plus, y_plus)
# for i in range(6, size):
#     if i % 3 == 0:
#         pyautogui.scroll(-72)
#     else:
#         pyautogui.scroll(-74)
