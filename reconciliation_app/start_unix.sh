#!/bin/bash

# 進入腳本所在目錄
cd "$(dirname "$0")"

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null
then
    echo "Python 3 未安裝。請訪問 https://www.python.org/downloads/ 安裝 Python 3.8 或更高版本。"
    exit 1
fi

# 檢查並安裝依賴
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "安裝依賴失敗。請檢查您的網絡連接或 Python 環境。"
    exit 1
fi

# 啟動 Flask 應用 (後台運行)
nohup python3 app.py > app.log 2>&1 &

# 等待應用啟動
sleep 5

# 打開瀏覽器
if command -v xdg-open &> /dev/null; then
    xdg-open http://127.0.0.1:5000
elif command -v open &> /dev/null; then
    open http://127.0.0.1:5000
elif command -v gnome-open &> /dev/null; then
    gnome-open http://127.0.0.1:5000
else
    echo "請手動在瀏覽器中打開 http://127.0.0.1:5000"
fi

echo "對賬單應用已在瀏覽器中打開。請勿關閉此終端窗口，否則應用將停止。"
