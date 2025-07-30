@echo off
rem @echo off 可以讓執行的指令本身不要顯示在畫面上，讓輸出更乾淨

echo [INFO] 準備啟動虛擬環境 (venv)...
call .\venv\Scripts\activate.bat

echo [INFO] 虛擬環境已啟動，準備執行 Python 腳本 (app.py)...
python app.py