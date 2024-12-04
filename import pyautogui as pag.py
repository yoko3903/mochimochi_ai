import pyautogui as pag

pag_alert = pag.alert(text='pyautogui', title='qiita', button='OK')
print(pag_alert) # 出力は「OK」