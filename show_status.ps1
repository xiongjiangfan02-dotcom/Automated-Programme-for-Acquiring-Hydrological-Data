param(
    [string]$Title = "贵州省大型水库实时水情采集",
    [string]$Message = "程序已执行。",
    [ValidateRange(1, 120)]
    [int]$TimeoutSeconds = 15,
    [ValidateSet(0, 16, 32, 48, 64)]
    [int]$Icon = 0
)

$ErrorActionPreference = "Stop"

try {
    # WScript.Shell.Popup is part of Windows and does not require an extra module.
    # The timeout makes the notification non-blocking for unattended daily runs.
    $shell = New-Object -ComObject WScript.Shell
    $shell.Popup($Message, $TimeoutSeconds, $Title, $Icon) | Out-Null
    exit 0
}
catch {
    # A scheduled task may run without an interactive desktop. The collector log
    # remains the source of truth; failure to show a popup must not fail collection.
    exit 0
}
