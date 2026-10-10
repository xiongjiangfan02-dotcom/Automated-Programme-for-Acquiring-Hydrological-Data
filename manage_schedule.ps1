[CmdletBinding()]
param(
    [ValidateSet("Install", "Show", "Remove")]
    [string]$Action = "Install",
    [string]$DailyAt = "",
    [ValidateRange(1, 24)][int]$DailyCount = 1,
    [string]$TaskName = "",
    [string]$Config = "config/default.json",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $projectRoot 'schedule_helpers.ps1')
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $projectRoot $Config }
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
    throw "配置文件不存在: $configPath"
}

$settings = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $DailyAt) { $DailyAt = [string]$settings.schedule.daily_time }
if (-not $PSBoundParameters.ContainsKey('DailyCount') -and $null -ne $settings.schedule.daily_count) {
    $configuredCount = $settings.schedule.daily_count
    if ($configuredCount -isnot [int] -and $configuredCount -isnot [long]) { throw '每天采集次数必须为整数。' }
    $DailyCount = $configuredCount
}
if (-not $TaskName) { $TaskName = [string]$settings.schedule.task_name }
if (-not $DailyAt) { $DailyAt = "09:00" }
if (-not $TaskName) { $TaskName = "贵州省大型水库实时水情采集" }

$dailyTimes = @(Get-CollectorDailyTimes -DailyAt $DailyAt -DailyCount $DailyCount)
if ([string]::IsNullOrWhiteSpace($TaskName) -or $TaskName.IndexOfAny([char[]]'\/*?[]') -ge 0) {
    throw "任务名称不能为空，也不能包含路径分隔符或通配符。"
}

$scheduledRunner = Join-Path $projectRoot "run_scheduled.ps1"
$taskDescription = "GuizhouHydroCollector: 每天采集贵州省大型水库实时水情并显示自动关闭的状态说明弹窗；当前用户登录时运行。"
$existing = Get-ScheduledTask -TaskName $TaskName -TaskPath '\' -ErrorAction SilentlyContinue
if ($existing -and $Action -ne "Show" -and
    (-not ([string]$existing.Description).StartsWith("GuizhouHydroCollector:", [StringComparison]::OrdinalIgnoreCase) -or
     $existing.Actions[0].Arguments.IndexOf($scheduledRunner, [StringComparison]::OrdinalIgnoreCase) -lt 0)) {
    throw "同名任务不是当前项目创建的任务，拒绝覆盖或删除: $TaskName"
}
switch ($Action) {
    "Show" {
        if (-not $existing) {
            Write-Output "定时任务不存在: $TaskName"
            exit 1
        }
        $info = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath '\'
        [pscustomobject]@{
            TaskName = $existing.TaskName
            State = $existing.State
            DailyAt = (@($existing.Triggers | ForEach-Object { $_.StartBoundary.Substring(11, 5) }) -join ', ')
            DailyCount = @($existing.Triggers).Count
            LastRunTime = $info.LastRunTime
            LastTaskResult = $info.LastTaskResult
            NextRunTime = $info.NextRunTime
            Execute = $existing.Actions[0].Execute
            Arguments = $existing.Actions[0].Arguments
        } | Format-List
        exit 0
    }
    "Remove" {
        if (-not $existing) {
            Write-Output "定时任务不存在，无需删除: $TaskName"
            exit 0
        }
        Unregister-ScheduledTask -TaskName $TaskName -TaskPath '\' -Confirm:$false
        Write-Output "已删除定时任务: $TaskName"
        exit 0
    }
}

if ($existing -and -not $Force) {
    throw "定时任务已存在。如需更新时间，请添加 -Force。任务名: $TaskName"
}

$windowsPowerShell = Join-Path $env:SystemRoot "System32/WindowsPowerShell/v1.0/powershell.exe"
if (-not (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf)) {
    throw "找不到 Windows PowerShell: $windowsPowerShell"
}

if ((Get-TimeZone).Id -ne "China Standard Time") {
    throw "定时任务要求电脑时区为北京时间（China Standard Time），请先在 Windows 设置中确认时区。"
}
$venvPython = Join-Path $projectRoot ".venv/Scripts/python.exe"
if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
    throw "未找到项目 Python 环境，请先执行 setup.ps1。"
}
$previousPythonPath = $env:PYTHONPATH
try {
    $env:PYTHONPATH = Join-Path $projectRoot "src"
    & $venvPython -m bijie_hydro.cli --config $configPath validate-template | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "配置或模板验证失败，未创建定时任务。" }
}
finally {
    $env:PYTHONPATH = $previousPythonPath
}

$taskArguments = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Config "{1}"' -f $scheduledRunner, $configPath
$taskAction = New-ScheduledTaskAction -Execute $windowsPowerShell -Argument $taskArguments -WorkingDirectory $projectRoot
$triggers = @(foreach ($clockTime in $dailyTimes) {
    $parts = $clockTime.Split(':')
    $nextRun = (Get-Date).Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
    if ($nextRun -le (Get-Date)) { $nextRun = $nextRun.AddDays(1) }
    New-ScheduledTaskTrigger -Daily -At $nextRun
})
$taskSettings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
$definition = New-ScheduledTask `
    -Action $taskAction `
    -Trigger $triggers `
    -Settings $taskSettings `
    -Principal $principal `
    -Description $taskDescription
Register-ScheduledTask -TaskName $TaskName -TaskPath '\' -InputObject $definition -Force | Out-Null

$registered = Get-ScheduledTask -TaskName $TaskName -TaskPath '\'
$taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath '\'
Write-Output "定时任务已创建或更新。"
[pscustomobject]@{
    TaskName = $registered.TaskName
    DailyAt = ($dailyTimes -join ', ')
    DailyCount = $DailyCount
    State = $registered.State
    NextRunTime = $taskInfo.NextRunTime
    RunAs = $currentUser
    LogonRequirement = "当前用户必须处于登录状态"
    Config = $configPath
} | Format-List
