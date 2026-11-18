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
& $Python -m PyInstaller --noconfirm --clean --log-level WARN --onefile --windowed --name xCCB_Report --distpath $DistPath --hidden-import win32timezone --collect-all tzdata --collect-data datasyncx --copy-metadata datasyncx --collect-submodules oracledb --add-data "${hrHelper};datasyncx\utils" dashboard.py
if ($LASTEXITCODE -ne 0) { throw 'Executable build failed' }
$executable = Join-Path $DistPath 'xCCB_Report.exe'
& $Python -c "import sys; from PyInstaller.archive.readers import CArchiveReader; archive = CArchiveReader(sys.argv[1]); pyz = archive.open_embedded_archive(next(name for name in archive.toc if name.endswith('.pyz'))); assert 'win32timezone' in pyz.toc, 'Missing frozen Windows authentication dependency: win32timezone'" $executable
if ($LASTEXITCODE -ne 0) { throw 'Packaged Windows authentication dependency check failed' }
Write-Output ("Built " + $executable)
