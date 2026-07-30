@echo off
setlocal
cd /d "%~dp0"
echo === 単価マスター メンテナンスツール（デモモード／DB未接続）===

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE ( where python3 >nul 2>nul && set "PYEXE=python3" )
if not defined PYEXE (
  echo [エラー] Python が見つかりません。https://www.python.org/downloads/ からインストールしてください。
  pause
  exit /b 1
)
%PYEXE% -m pip install --disable-pip-version-check flask
echo ブラウザを開きます: http://127.0.0.1:5000
start "" http://127.0.0.1:5000
%PYEXE% app.py --demo
echo.
pause
