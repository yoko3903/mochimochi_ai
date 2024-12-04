import pyautogui

#まずは現在のマウスカーソルの位置を出力
print(pyautogui.position())
#メイン画面のサイズを取得 ※マルチモニタの場合はメイン設定のディスプレイが基準になる
print(pyautogui.size())