# Builds dist\Python ETL Multitasker\Python ETL Multitasker.exe (the GUI only; the engine stays the project's code).
# Usage (from the project folder):  powershell -ExecutionPolicy Bypass -File tools\build_exe.ps1 [-Python <python.exe>]
# Without -Python, the first "python" on PATH is used (or the "py" launcher).
param(
    [string]$Python = ""
)
$project = Split-Path -Parent $PSScriptRoot
if (-not $Python) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $Python = $cmd.Source }
    elseif (Get-Command py -ErrorAction SilentlyContinue) { $Python = (& py -3 -c "import sys; print(sys.executable)") }
}
if (-not $Python -or -not (Test-Path $Python)) { throw "Python not found: pass -Python <path to python.exe>" }
Set-Location $project
& $Python -m PyInstaller --noconfirm --clean --distpath (Join-Path $project "dist") `
    --workpath (Join-Path $project "build") (Join-Path $project "tools\python_etl_multitasker.spec")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed ($LASTEXITCODE)" }
Write-Output "Built: $(Join-Path $project 'dist\Python ETL Multitasker\Python ETL Multitasker.exe')"
