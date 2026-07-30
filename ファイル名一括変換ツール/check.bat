@echo off
cd /d %~dp0
chcp 65001 > nul
echo === Python チェック ===
python --version 2>nul && echo [OK] python が使えます || echo [NG] python は使えません
py --version 2>nul && echo [OK] py ランチャーが使えます || echo [NG] py ランチャーもありません
python3 --version 2>nul && echo [OK] python3 が使えます || echo [NG] python3 もありません

echo.
echo === ファイル確認 ===
if exist rename_tool.py (echo [OK] rename_tool.py あり) else (echo [NG] rename_tool.py なし)
if exist build.bat (echo [OK] build.bat あり) else (echo [NG] build.bat なし)

echo.
echo === 現在のフォルダ ===
cd

echo.
pause
