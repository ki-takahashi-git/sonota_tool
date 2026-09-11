@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set LOGFILE=%~dp0build_log.txt

echo Build log: %LOGFILE%
echo.

call :main > "%LOGFILE%" 2>&1
set RESULT=%ERRORLEVEL%

echo ================= LOG =================
type "%LOGFILE%"
echo ========================================
echo.
if "%RESULT%"=="0" (
  echo Done! See the dist folder for the finished .exe.
) else (
  echo Build FAILED. See the log above.
  echo ^(also saved to: %LOGFILE%^)
)
echo.
pause
exit /b %RESULT%

:main
echo Installing required packages...
pip install --timeout 15 -v -r requirements.txt
if errorlevel 1 exit /b 1

echo.
echo Installing PyInstaller (build tool)...
pip install --timeout 15 -v pyinstaller
if errorlevel 1 exit /b 1

echo.
echo Bumping version number...
python bump_version.py
if errorlevel 1 exit /b 1

echo.
echo Building EXE... this may take a few minutes.
rem exclude-module: heavy libraries that happen to be installed globally on
rem this dev machine but are never used by this tool. Excluding them keeps
rem the EXE smaller and faster to unpack on every launch.
pyinstaller --noconfirm --onefile --windowed --name "マミヤ支給明細照合ツール" --add-data "使い方.md;." --add-data "VERSION;." --exclude-module pandas --exclude-module scipy --exclude-module matplotlib --exclude-module cv2 --exclude-module pyarrow --exclude-module uvicorn --exclude-module websockets --exclude-module IPython --exclude-module notebook --exclude-module jupyter --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module torch --exclude-module tensorflow --exclude-module sklearn --exclude-module numpy.f2py --exclude-module tornado --exclude-module zmq --exclude-module PyQt5 --exclude-module PySide2 --exclude-module PySide6 --exclude-module PyQt6 --exclude-module docx --exclude-module pptx --exclude-module pydantic --exclude-module fastapi --exclude-module gradio --exclude-module streamlit shikyu_check_tool.py
if errorlevel 1 exit /b 1

exit /b 0
