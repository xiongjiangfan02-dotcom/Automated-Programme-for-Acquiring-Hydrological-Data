$ErrorActionPreference = 'Stop'
# Load CIM proxy functions before test replacements so autoloading cannot
# replace our functions and reach the real task service.
Import-Module ScheduledTasks
$testProject = Split-Path -Parent $PSScriptRoot
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('hydro-schedule-test-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
$testConfig = Join-Path $testRoot 'settings.json'
$testSettings = Get-Content -LiteralPath (Join-Path $testProject 'config/default.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$testSettings.schedule.task_name = 'HydroIsolatedTest-' + [guid]::NewGuid().ToString('N')
$script:registeredDefinition = $null
$script:registerCount = 0
$script:foreignTask = $false

# Replace only the task-service reads/writes. Constructors and template validation
# use the installed Windows APIs. No real task is installed, updated or removed.
function Get-ScheduledTask {
    param($TaskName, $TaskPath, $ErrorAction)
    if ($script:foreignTask) {
        return [pscustomobject]@{ Description = 'Another application'; Actions = @([pscustomobject]@{ Arguments = 'other.ps1' }) }
    }
    if ($script:registeredDefinition) {
        return [pscustomobject]@{
            TaskName = $TaskName; State = 'Ready'
            Triggers = $script:registeredDefinition.Triggers
            Actions = $script:registeredDefinition.Actions
            Description = $script:registeredDefinition.Description
        }
    }
}
function Register-ScheduledTask {
    param($TaskName, $TaskPath, $InputObject, [switch]$Force)
    $script:registeredDefinition = $InputObject
    $script:registerCount++
}
function Get-ScheduledTaskInfo {
    param($TaskName, $TaskPath)
    return [pscustomobject]@{ NextRunTime = [datetime]$script:registeredDefinition.Triggers[0].StartBoundary }
}
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "测试失败：$Message" }
}
try {
    foreach ($count in @(1, 2, 24)) {
        Assert-True ((Get-Command Register-ScheduledTask).ModuleName -eq '') '任务服务写入必须被测试替身拦截'
        $testSettings.schedule.daily_time = '23:59'
        $testSettings.schedule.daily_count = $count
        [System.IO.File]::WriteAllText($testConfig, ($testSettings | ConvertTo-Json -Depth 100), [System.Text.UTF8Encoding]::new($false))
        . (Join-Path $testProject 'manage_schedule.ps1') -Action Install -Config $testConfig -Force | Out-Null
        Assert-True ($null -ne $script:registeredDefinition) '注册必须由测试替身处理'
        Assert-True (@($script:registeredDefinition.Triggers).Count -eq $count) "每天$count 次的触发器数量"
        $registeredTimes = @($script:registeredDefinition.Triggers | ForEach-Object {
            [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([datetimeoffset]::Parse($_.StartBoundary), 'China Standard Time').ToString('HH:mm')
        })
        $expectedTimes = @(Get-CollectorDailyTimes '23:59' $count)
        Assert-True (($registeredTimes -join ',') -eq ($expectedTimes -join ',')) "触发时刻精确一致：实际 $($registeredTimes -join ',')，预期 $($expectedTimes -join ',')"
        foreach ($triggerItem in $script:registeredDefinition.Triggers) {
            Assert-True ([datetimeoffset]::Parse($triggerItem.StartBoundary).UtcDateTime -gt [datetime]::UtcNow) '首次触发安排在未来'
            Assert-True ($triggerItem.DaysInterval -eq 1) '每日循环'
        }
        Assert-True ($script:registeredDefinition.Actions[0].Arguments.Contains($testConfig)) '任务使用指定配置'
    }
    $beforeReject = $script:registerCount
    $script:foreignTask = $true
    $rejected = $false
    try { . (Join-Path $testProject 'manage_schedule.ps1') -Action Install -Config $testConfig -Force | Out-Null } catch { $rejected = $true }
    Assert-True ($rejected -and $script:registerCount -eq $beforeReject) '不能覆盖其他应用任务'
    Write-Output 'PASS: daily triggers (1/2/24), midnight wrap, future first run, configuration action, foreign-task protection; no task-service writes.'
}
finally {
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    $allowedPrefix = Join-Path ([System.IO.Path]::GetTempPath()) 'hydro-schedule-test-'
    if (-not $resolvedTestRoot.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw '拒绝清理预期临时目录外的路径。' }
    Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
}
