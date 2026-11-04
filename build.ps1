param(
    [string]$Python = (Join-Path $PSScriptRoot '.venv-build\Scripts\python.exe'),
    [string]$DistPath = 'dist'
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath $Python)) {
    throw 'Create .venv-build with Python 3.11-3.13 and install requirements-build.txt before building.'
}
& $Python -c "import sys, importlib.metadata as metadata; assert (3, 11) <= sys.version_info[:2] < (3, 14), 'DataSyncX requires Python 3.11-3.13'; assert metadata.version('datasyncx') == '1.1.6', 'Install requirements-build.txt'"
if ($LASTEXITCODE -ne 0) { throw 'Build prerequisites are not satisfied' }
$hrHelper = & $Python -c "import importlib.util, pathlib; print(pathlib.Path(importlib.util.find_spec('datasyncx').origin).parent / 'utils' / 'get_user_email.py')"
if ($LASTEXITCODE -ne 0) { throw 'Cannot locate the DataSyncX HR helper' }
& $Python -m PyInstaller --noconfirm --clean --log-level WARN --onefile --windowed --name xCCB_Report --distpath $DistPath --collect-all tzdata --collect-data datasyncx --copy-metadata datasyncx --collect-submodules oracledb --add-data "${hrHelper};datasyncx\utils" dashboard.py
if ($LASTEXITCODE -ne 0) { throw 'Executable build failed' }
$reviewerPath = Join-Path (Join-Path $PSScriptRoot $DistPath) 'reviewer_names.txt'
if (-not (Test-Path -LiteralPath $reviewerPath)) {
    [System.IO.File]::WriteAllText($reviewerPath, '')
}
Write-Output ("Built " + (Join-Path $DistPath 'xCCB_Report.exe'))
