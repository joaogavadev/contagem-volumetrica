@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >/dev/null 2>&1
if %errorlevel%==0 (
  py -3 main.py
  goto fim
)

python main.py
if errorlevel 1 (
  echo.
  echo Python nao encontrado. Rode primeiro:
  echo    powershell -ExecutionPolicy Bypass -File instalar.ps1
  echo.
  pause
)

:fim
