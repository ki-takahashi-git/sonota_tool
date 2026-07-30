@echo off
chcp 65001 > nul
echo ========================================
echo  ファイル名一括変換ツール  ビルダー
echo ========================================
echo.

:: PyInstaller がなければインストール
pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo PyInstaller をインストールしています...
    pip install pyinstaller
    if errorlevel 1 (
        echo.
        echo [エラー] pip が見つかりません。Python がインストールされているか確認してください。
        pause
        exit /b 1
    )
)

echo ビルドを開始します...
echo.

pyinstaller ^
    --onefile ^
    --noconsole ^
    --name "ファイル名一括変換ツール" ^
    rename_tool.py

if errorlevel 1 (
    echo.
    echo [エラー] ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo  完了！
echo  dist\ファイル名一括変換ツール.exe
echo  を使ってください。
echo ========================================
echo.
pause
