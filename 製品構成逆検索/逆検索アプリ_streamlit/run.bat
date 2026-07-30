@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo  Gyaku-kensaku Tool (Streamlit)
echo ==========================================
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE ( where python3 >nul 2>nul && set "PYEXE=python3" )
if not defined PYEXE goto NOPY
%PYEXE% --version
echo Installing required libraries (first time only)...
%PYEXE% -m pip install streamlit pymssql openpyxl
echo Browser will open automatically (http://localhost:8501).
%PYEXE% -m streamlit run streamlit_app.py
goto END
:NOPY
echo [ERROR] Python not found.
echo Install from https://www.python.org/downloads/
echo and check "Add python.exe to PATH" during install.
:END
echo.
pause
