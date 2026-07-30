@echo off
setlocal
cd /d "%~dp0"
echo === Gyaku-kensaku Tool (Browser / DEMO, no DB) ===
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE ( where python3 >nul 2>nul && set "PYEXE=python3" )
if not defined PYEXE goto NOPY
%PYEXE% -m pip install flask openpyxl
echo Opening browser: http://127.0.0.1:5001
start "" http://127.0.0.1:5001
%PYEXE% app.py --demo
goto END
:NOPY
echo [ERROR] Python not found.
echo Install from https://www.python.org/downloads/
:END
echo.
pause
