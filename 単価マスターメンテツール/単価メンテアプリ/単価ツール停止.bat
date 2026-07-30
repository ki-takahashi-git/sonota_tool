@echo off
setlocal
echo ポート5000で動作中のツールを停止します...
set "FOUND="
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000" ^| findstr LISTENING') do (
  taskkill /PID %%a /F >nul 2>nul && set "FOUND=1"
)
if defined FOUND ( echo 停止しました。) else ( echo 起動中のツールは見つかりませんでした。)
echo.
pause
