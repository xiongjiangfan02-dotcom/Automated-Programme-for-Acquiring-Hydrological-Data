param([string]$Python = "python")

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$environmentPath = Join-Path $projectRoot ".venv"
$environmentPython = Join-Path $environmentPath "Scripts/python.exe"
if (-not (Test-Path -LiteralPath $environmentPython)) {
    & $Python -m venv $environmentPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
& $environmentPython -m pip install -r (Join-Path $projectRoot "requirements.lock.txt")
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $environmentPython -m pip install --no-deps -e $projectRoot
exit $LASTEXITCODE
