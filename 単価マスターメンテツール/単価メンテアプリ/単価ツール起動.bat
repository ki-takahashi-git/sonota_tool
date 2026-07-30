@echo off
setlocal
cd /d "%~dp0"
rem ---- 黒い画面を出さずに起動（pythonw）。ブラウザは自動で開きます ----
set "PYC="
where py >nul 2>nul && set "PYC=py"
if not defined PYC ( where python >nul 2>nul && set "PYC=python" )
if not defined PYC (
  echo Python が見つかりません。先に「初回セットアップ.bat」を実行してください。
  pause
  exit /b 1
)
%PYC% -c "import flask, openpyxl" 2>nul
if errorlevel 1 (
  echo 必要なライブラリが未導入です。先に「初回セットアップ.bat」を実行してください。
  pause
  exit /b 1
)
set "PYW="
where pythonw >nul 2>nul && set "PYW=pythonw"
if not defined PYW ( where pyw >nul 2>nul && set "PYW=pyw" )
if not defined PYW set "PYW=%PYC%"
start "" %PYW% "%~dp0app.py"
echo 起動しました。ブラウザが自動で開きます（数秒お待ちください）。
echo 開かない場合は http://127.0.0.1:5000 を開いてください。
