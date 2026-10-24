@echo off
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" (
  echo.
  echo 用法：把要编辑的 HTML 文件拖到 editor.cmd 上，再松手。
  echo 例如：把 weibo\index.html 拖过来。
  echo.
  pause
  exit /b 1
)

echo [1/2] 同步样式...
python tools\inline-css.py

echo [2/2] 启动编辑器...
echo 首次运行会下载编辑器组件，等一两分钟。
python "%~dp0tools\workbench\scripts\workbench.py" open "%~1" --editor-root "%~dp0." --port 4317

echo.
echo 编辑器地址：http://127.0.0.1:4317/
echo 改完记得点保存，然后刷新预览页。
echo.
pause
