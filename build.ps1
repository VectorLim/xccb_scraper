$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
& .\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name xCCB_Report --collect-all tzdata dashboard.py
if ($LASTEXITCODE -ne 0) { throw 'Executable build failed' }
if (-not (Test-Path -LiteralPath 'dist\reviewer_names.txt')) {
    [System.IO.File]::WriteAllText((Join-Path $PSScriptRoot 'dist\reviewer_names.txt'), '')
}
Write-Output 'Built dist\xCCB_Report.exe'
