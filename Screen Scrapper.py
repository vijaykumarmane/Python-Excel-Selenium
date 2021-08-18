#!/usr/bin/python
# -*- coding: utf-8 -*-
import xlwings as xw
import pyautogui as pa
import time as tm
import win32clipboard as wc
import PIL
import re

path = excelFilePath

# Open workbook

wb = xw.Book(path)
sht = wb.sheets[0]
sht2 = wb.sheets[1]

page = 1
row = 2
for i in range(1, 101):
    sht.range((1, i)).value = page
    tm.sleep(4)
    pa.moveTo(39, 399, duration=.5)
    pa.click()
    pa.hotkey('ctrl', 'a')
    pa.hotkey('ctrl', 'c')
    tm.sleep(.5)
    pa.click()

    # To paste copied content: I think this works for excel range copy paste
#     sht.range((2, i)).select()
#     tm.sleep(0.3)
#     sht.api.PasteSpecial(-4163) # Not working so used simple as usual data.split("\n")
    # To read copy clipboard

    wc.OpenClipboard()
    data = wc.GetClipboardData()
    wc.CloseClipboard()
    tm.sleep(0.3)
    sht.range((2, i)).options(transpose=True).value = data.split('\n')
    pa.press('end')
    tm.sleep(0.3)
    pa.moveTo(843, 482, duration=.5)

    # Locate snipping image on the screen

    start = pa.locateCenterOnScreen(r'C:\Users\Admin\Desktop\Next.png')  # If the file is not a png file it will not work
    if str(start) == 'None':
        start = \
            pa.locateCenterOnScreen(r'C:\Users\Admin\Desktop\Next1.png')
    if str(start) == 'None':
        print 'Not detected!'
        break
    pa.moveTo(start)
    pa.click()
    tm.sleep(1)
    pa.moveTo(465, 54, duration=.5)
    pa.click()
    pa.hotkey('ctrl', 'c')

    # To read copy clipboard

    wc.OpenClipboard()
    data = wc.GetClipboardData()
    wc.CloseClipboard()
    page = re.findall('page=([0-9]+)', data)[0]
    page = int(page)

    # To print details

    for r in range(1, 250):
        if 'degree connection' in str(sht.range((r, i)).value):
            name = sht.range((r - 1, i)).value.replace('\r', '')
            designation = sht.range((r + 1, i)).value.replace('\r', '')
            location = sht.range((r + 2, i)).value.replace('\r', '')
            sht2.range((row, 1)).value = (re.findall('((.+))View ',
                    name)[0][0], designation, location)
            print (re.findall('((.+))View ', name)[0][0], designation,
                   location)
            row += 1
