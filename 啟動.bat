@echo off
chcp 65001 >nul
echo 正在啟動 GCP 用量明細 PDF 產生器...

:: 確認 Python 已安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo 錯誤：找不到 Python，請先安裝 Python 3.10 以上版本。
    echo 下載網址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 安裝必要套件（已安裝則跳過）
echo 檢查必要套件...
pip install -q -r "%~dp0requirements.txt"

:: 啟動程式
python "%~dp0gcp_app.py"
