Set-Location $PSScriptRoot

$pythonExe = "C:\Users\ki-takahashi\AppData\Local\Programs\Python\Python312\python.exe"

$distDir = $PSScriptRoot
$workDir = "$PSScriptRoot\build"

Write-Host "Python: $pythonExe" -ForegroundColor Green
& $pythonExe --version

Write-Host "Installing dependencies..." -ForegroundColor Cyan
& $pythonExe -m pip install pyinstaller pillow --quiet

Write-Host "Generating icon..." -ForegroundColor Cyan
& $pythonExe make_icon.py

$iconArg = if (Test-Path "app_icon.ico") { "--icon=app_icon.ico" } else { "" }

Write-Host "Building to $distDir..." -ForegroundColor Cyan
if ($iconArg) {
    & $pythonExe -m PyInstaller --onefile --noconsole --distpath $distDir --workpath $workDir --name "FileRenamer" $iconArg rename_tool.py
} else {
    & $pythonExe -m PyInstaller --onefile --noconsole --distpath $distDir --workpath $workDir --name "FileRenamer" rename_tool.py
}

if ($LASTEXITCODE -eq 0) {
    $finalName = "フォルダ&ファイル名一括変更ツール_TD規則準拠版.exe"
    if (Test-Path "$distDir\$finalName") { Remove-Item "$distDir\$finalName" -Force }
    Rename-Item "$distDir\FileRenamer.exe" $finalName
    Write-Host "Done! $distDir\$finalName" -ForegroundColor Green
    Start-Process explorer.exe $distDir
} else {
    Write-Host "Build failed." -ForegroundColor Red
}

Read-Host "Press Enter to exit"
