@echo off
setlocal
cd /d "%~dp0"
echo === Gyaku-kensaku Tool (Streamlit / DEMO, no DB) ===
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE ( where python3 >nul 2>nul && set "PYEXE=python3" )
if not defined PYEXE goto NOPY
%PYEXE% -m pip install streamlit openpyxl
echo Browser will open automatically (http://localhost:8501).
%PYEXE% -m streamlit run streamlit_app.py -- --demo
goto END
:NOPY
echo [ERROR] Python not found.
echo Install from https://www.python.org/downloads/
:END
echo.
pause
