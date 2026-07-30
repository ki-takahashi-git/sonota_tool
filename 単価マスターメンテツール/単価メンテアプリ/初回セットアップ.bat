@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo  単価マスター メンテナンスツール 初回セットアップ
echo ============================================
set "PYC="
where py >nul 2>nul && set "PYC=py"
if not defined PYC ( where python >nul 2>nul && set "PYC=python" )
if not defined PYC (
  echo [エラー] Python が見つかりません。
  echo https://www.python.org/downloads/ からインストールし、
  echo インストール画面で「Add python.exe to PATH」にチェックを入れてください。
  echo その後もう一度このファイルを実行してください。
  echo.
  pause
  exit /b 1
)
echo 使用するPython: %PYC%
%PYC% --version
echo.
echo 必要なライブラリをインストールします（数分かかることがあります）...
%PYC% -m pip install --disable-pip-version-check flask pymssql waitress openpyxl
echo.
echo セットアップが完了しました。次回からは「単価ツール起動.bat」をダブルクリックしてください。
echo.
pause
