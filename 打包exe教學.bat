@echo off
chcp 65001 >nul
echo ================================================
echo  將程式打包成單一 .exe 檔案（同事免安裝 Python）
echo ================================================
echo.

:: 安裝 PyInstaller
pip install -q pyinstaller

:: 打包（--onefile = 單一 exe，--noconsole = 不顯示黑視窗，--clean = 清除暫存）
pyinstaller --onefile --noconsole --clean ^
  --name "GCP用量明細產生器" ^
  --add-data "gcp_pdf_generator.py;." ^
  "%~dp0gcp_app.py"

echo.
echo 打包完成！exe 檔案在 dist\ 資料夾內。
echo 將 dist\GCP用量明細產生器.exe 傳給同事即可，免安裝任何軟體。
pause
