# Creates the desktop shortcut "Python ETL Multitasker" (runs the launcher with pythonw: no console window).
# Usage (from the project folder):  powershell -ExecutionPolicy Bypass -File tools\create_shortcut.ps1 [-Python <pythonw.exe>]
# Without -Python, pythonw.exe next to the first "python" on PATH is used (or the "py" launcher's interpreter).
param(
    [string]$Python = ""
)
$project = Split-Path -Parent $PSScriptRoot
if (-not $Python) {
    $exe = $null
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $exe = $cmd.Source }
    elseif (Get-Command py -ErrorAction SilentlyContinue) { $exe = (& py -3 -c "import sys; print(sys.executable)") }
    if ($exe) { $Python = Join-Path (Split-Path -Parent $exe) "pythonw.exe" }
}
$launcher = Join-Path $project "Python ETL Multitasker.pyw"
$icon = Join-Path $project "gui\resources\app.ico"
if (-not $Python -or -not (Test-Path $Python)) { throw "pythonw.exe not found: pass -Python <path to pythonw.exe>" }
if (-not (Test-Path $launcher)) { throw "launcher not found: $launcher" }
$desktop = [Environment]::GetFolderPath("Desktop")
$lnk = Join-Path $desktop "Python ETL Multitasker.lnk"
$shell = New-Object -ComObject WScript.Shell
$s = $shell.CreateShortcut($lnk)
$s.TargetPath = $Python
$s.Arguments = '"' + $launcher + '"'
$s.WorkingDirectory = $project
$s.IconLocation = $icon
$s.Description = "Python ETL Multitasker - desktop interface for sales_etl"
$s.Save()
Write-Output "Shortcut created: $lnk"
