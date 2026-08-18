@echo off
setlocal
cd /d "%~dp0"
set "BUNDLED_PY=C:\Users\lygmf\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" server.py
) else (
  python server.py
)
if errorlevel 1 (
  echo.
  echo 启动失败。请确认本机已安装 Python 和 python-docx。
  pause
)
endlocal
