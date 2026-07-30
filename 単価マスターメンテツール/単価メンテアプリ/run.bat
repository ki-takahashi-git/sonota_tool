@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo  単価マスター メンテナンスツール 起動
echo ============================================

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE ( where python3 >nul 2>nul && set "PYEXE=python3" )
if not defined PYEXE (
  echo.
  echo [エラー] Python が見つかりません。
  echo https://www.python.org/downloads/ からインストールし、
  echo インストール時に「Add python.exe to PATH」にチェックを入れてください。
  echo.
  pause
  exit /b 1
)
echo 使用するPython: %PYEXE%
%PYEXE% --version
echo.
echo 必要なライブラリを確認/インストールします（初回のみ時間がかかります）...
%PYEXE% -m pip install --disable-pip-version-check flask pymssql
if errorlevel 1 (
  echo.
  echo [警告] pymssql のインストールに失敗した可能性があります。pyodbc を試します...
  %PYEXE% -m pip install --disable-pip-version-check pyodbc
)
echo.
echo ブラウザを開きます: http://127.0.0.1:5000
start "" http://127.0.0.1:5000
%PYEXE% app.py
echo.
echo === サーバーが終了しました。このウィンドウは閉じてOKです。 ===
pause
