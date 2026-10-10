param(
    [string]$Config = "config/default.json",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Python) {
    $venvPython = Join-Path $projectRoot ".venv/Scripts/python.exe"
    if (Test-Path -LiteralPath $venvPython) { $Python = $venvPython } else { $Python = "python" }
}
$env:PYTHONPATH = Join-Path $projectRoot "src"
$env:PYTHONUTF8 = "1"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $projectRoot $Config }
& $Python -m bijie_hydro.cli --config $configPath collect
exit $LASTEXITCODE

