import pyautogui
import time 
from datetime import timedelta
from math import floor
from pathlib import Path
import glob
import numpy as np
from scipy.signal import argrelmax
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
import os
import matplotlib.pyplot as plt
import sympy as sym

# 対象ディレクトリのパス
target_directory = (r'C:\Users\yokok\Downloads\stock-data')

# ディレクトリ内のExcelファイル一覧を取得
xlsx_files = [f for f in os.listdir(target_directory) if f.endswith('.xlsx')]
for xlsx_file in xlsx_files:
    # Excelファイルを読み込む
    df = pd.read_excel(os.path.join(target_directory, xlsx_file))
    row_count = len(df)
    df['data_num']= range(1,row_count+1)
    df['5日平均']= df['5日平均'].replace(',','')
    df['5日平均']= pd.to_numeric(df['5日平均'])

    # xとyの列データを取得
    x =np.array(df['data_num'],dtype=np.float32)
    y = np.array(df['5日平均'],dtype=np.float32)

    # 20次の多項式近似
    coeffs = np.polyfit(x, y, 20)
    polynomial = np.poly1d(coeffs)
    print(polynomial)
    derivative = polynomial.deriv()

    # argrelmax関数を使って極大値のインデックスを見つけます
    indices = argrelmax(derivative)

    # 極大値が複数存在するかどうかを確認します
    if len(indices) > 1:
        # 極大値を取得します
        maxima = data[indices]

        # 極大値が増加しているかどうかを判定します
        is_increasing = np.all(np.diff(maxima) > 0)
        if is_increasing:
            # 極大値が増加している場合にのみ特定の処理を行います
            print("極大値が増加しています。特定の処理を行います。")
            # ここに特定の処理を書くことができます
        else:
    else:

    # 微分した関数を表示します
    print(derivative)


    # 推定値を計算
    estimated_y = np.polyval(coeffs, x)


    # グラフで表示
    plt.scatter(x, y, label='Data')
    plt.plot(x, estimated_y, color='red', label='Polynomial Fit')
    plt.xlabel('x')
    plt.ylabel('y')
    plt.legend()
    plt.title(f'Polynomial Fit for {xlsx_file}')
    plt.show()
    time.sleep(5)
    plt.close()