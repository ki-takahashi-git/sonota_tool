@echo off
setlocal
cd /d "%~dp0"
echo === Build EXE (Browser / Flask, PyInstaller) ===
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE goto NOPY
echo Installing required libraries...
%PYEXE% -m pip install pyinstaller flask pymssql openpyxl
echo Building exe (takes a few minutes)...
%PYEXE% -m PyInstaller --noconfirm --clean --onefile --name "GyakuKensaku_Web" --add-data "templates;templates" --add-data "config.ini;." --hidden-import pymssql --hidden-import openpyxl --collect-submodules flask app.py
echo.
echo Done. See dist\GyakuKensaku_Web.exe
echo Put config.ini next to the exe to change the DB settings.
echo Double-click the exe: it starts the server and opens the browser automatically.
goto END
:NOPY
echo [ERROR] Python not found. Install from https://www.python.org/downloads/
:END
echo.
pause
