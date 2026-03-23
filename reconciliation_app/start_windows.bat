@echo off
cd /d "%~dp0"

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python 未安裝。請訪問 https://www.python.org/downloads/ 安裝 Python 3.8 或更高版本。
    echo 安裝時請務必勾選 "Add Python to PATH" 選項。
    pause
    exit /b 1
)

REM 檢查並安裝依賴
pip install -r requirements.txt >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝依賴失敗。請檢查您的網絡連接或 Python 環境。
    pause
    exit /b 1
)

REM 啟動 Flask 應用
start /b python app.py

REM 打開瀏覽器
start "" http://127.0.0.1:5000

echo 對賬單應用已在瀏覽器中打開。請勿關閉此窗口。

REM 等待用戶關閉窗口
pause
