@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] 同步样式...
python tools\inline-css.py
if errorlevel 1 (
  echo.
  echo [错误] 样式同步失败。请确认已安装 Python，并且在安装时勾选了 Add python.exe to PATH。
  echo.
  pause
  exit /b 1
)

echo [2/2] 启动预览服务器...
start "预览服务器（关闭本窗口即停止）" cmd /c "python -m http.server 8000"
timeout /t 2 >nul
start "" http://127.0.0.1:8000/

echo.
echo 预览地址：http://127.0.0.1:8000/
echo 关闭“预览服务器”窗口即可停止。
echo.
pause
