import pyautogui
import time 
from datetime import timedelta
from math import floor
from pathlib import Path
import glob
import numpy as np
import os
import pandas as pd
import talib as ta
from pandas_datareader import data as pdr
import datetime
import math
from math import isinf
import yfinance as yf
import warnings
import pyperclip


current_directory = os.getcwd()
relative_path = (r'C:\Users\yokok\OneDrive\デスクトップ\data_j.xlsx')
absolute_path = os.path.join(current_directory, relative_path)
List_Data = pd.read_excel(absolute_path)

start_index = 0


for index, CodeNo in enumerate(List_Data['コード'][start_index:]):
    actual_row_number = start_index + index
    CodeNo = str(CodeNo)
    print(str(actual_row_number) + " : " + str(CodeNo))

    # SBIアプリをクリック
    pyautogui.moveTo(1235, 1064)
    pyautogui.click()
    time.sleep(1)

    # チャートボタンをクリック
    pyautogui.moveTo(616, 64)
    pyautogui.click()
    time.sleep(2)

    # 時系列をクリック
    pyautogui.moveTo(880, 206)
    pyautogui.click()
    time.sleep(1)

    # 銘柄コードを入力
    pyautogui.moveTo(322, 166)
    #for _ in range(2):
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(CodeNo)
    pyautogui.press('enter')
    time.sleep(1)

    # 照会をクリック
    pyautogui.moveTo(747, 195)
    time.sleep(1)
    pyautogui.click()
    pyautogui.click()
    time.sleep(3)

     # CSVエクスポートをクリックしてダウンロード
    pyautogui.moveTo(1172, 199)
    pyautogui.click()
    time.sleep(1)

    pyautogui.moveTo(550, 201)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyperclip.copy(r"C:\Users\yokok\Downloads\stock-data")
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')

    pyautogui.moveTo(417, 583)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    pyautogui.write(CodeNo)

    pyautogui.moveTo(983, 651)
    pyautogui.click()
    time.sleep(2)

    path = os.getcwd()
    os.chdir(r"C:\Users\yokok\Downloads\stock-data")

    file = f"{CodeNo}.csv"
    data = pd.read_csv(file)
    xlsxfile = f"{CodeNo}.xlsx"
    data.to_excel(xlsxfile,engine='openpyxl',index=False)
    os.remove(file)
    os.chdir(path)
