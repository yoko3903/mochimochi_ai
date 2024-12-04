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
import datetime
import warnings
import pyperclip
import os
import matplotlib.pyplot as plt

# 対象ディレクトリのパス
target_directory = (r'C:\Users\yokok\Downloads\stock-data')

# ディレクトリ内のExcelファイル一覧を取得
xlsx_files = [f for f in os.listdir(target_directory) if f.endswith('.xlsx')]
for xlsx_file in xlsx_files:
    # Excelファイルを読み込む
    df = pd.read_excel(os.path.join(target_directory, xlsx_file))
    row_count = len(df)
    df['data_num']= range(1,row_count+1)
    df['終値'] = df['終値'].astype(str).str.replace(',', '')
    df['終値']= pd.to_numeric(df['終値'])

    fast_period = 12
    slow_period = 26
    signal_period = 9

    df["MACD"], df["MACD_signal"], df["MACD_hist"] = ta.MACD(df['終値'], fast_period, slow_period, signal_period)
    #print(df["MACD"])
    #print(df["MACD_signal"])
    #print(df["MACD_hist"])


    # 売買シグナルの生成
    # MACDがSignalを下から上に突き抜けたとき
    df['Buy_Signal']  = (df["MACD"] > df["MACD_signal"]) & (df["MACD"].shift() < df["MACD_signal"].shift())  
    # MACDがSignalを上から下に突き抜けたとき
    df['Sell_Signal']  = (df["MACD"] < df["MACD_signal"]) & (df["MACD"].shift() > df["MACD_signal"].shift())  
    #print(df['Buy_Signal'])
    #print(df['Sell_Signal'])
    rsi_period = 14

    df['RSI'] = ta.RSI(df['終値'])
    df['Buy Signal'] = (df['RSI'] < 30).astype(int)
    df['Sell Signal'] = (df['RSI'] > 70).astype(int)
    #print(df['RSI'])
    #print(df['Buy Signal'])
    #print(df['Sell Signal'])

    # 新しいfigureを作成
    plt.figure(figsize=(24,16))

    # 終値をプロット
    plt.scatter(df['data_num'], df['終値'], label='Data')
    plt.plot(df['data_num'], df['終値'], color='red', label='Polynomial Fit')
    plt.plot(df['data_num'],df['終値'], label='終値')
    plt.show()
    plt.pause(2)
    plt.close()

    # MACDとMACD_signalをプロット
    plt.figure(figsize=(12,8))
    plt.plot(df['MACD'], label='MACD')
    plt.plot(df['MACD_signal'], label='MACD Signal')
    plt.show()
    plt.pause(2)
    plt.close()


    # RSIをプロット
    plt.figure(figsize=(12,8))
    plt.plot(df['RSI'], label='RSI')
    plt.show()
    plt.pause(2)
    plt.close()



    # Buy SignalとSell Signalをプロット
    plt.figure(figsize=(12,8))
    plt.plot(df[df['Buy Signal'] == 1].index, df['終値'][df['Buy Signal'] == 1], '^', markersize=10, color='g', label='buy signal')
    plt.plot(df[df['Sell Signal'] == 1].index, df['終値'][df['Sell Signal'] == 1], 'v', markersize=10, color='r', label='sell signal')

    # 凡例を表示
    plt.legend(loc='upper left')

    # グラフを表示
    plt.show()
    plt.pause(2)
    plt.close()







    # 最新のデータを取得
    latest_data = df.iloc[-1]
    filename = os.path.basename(xlsx_file)
    
    # MACDとRSIの両方が買いシグナルを示している場合
    if latest_data['Buy_Signal'] and latest_data['Buy Signal']:
        #print("Buy at price: ", latest_data['終値'])
        print(filename)

    # MACDとRSIの両方が売りシグナルを示している場合
    elif latest_data['Sell_Signal'] and latest_data['Sell Signal']:
        #print("Sell at price: ", latest_data['終値'])
        print(filename)
