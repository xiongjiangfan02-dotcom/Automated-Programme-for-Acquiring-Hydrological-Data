param(
    [string]$Config = "config/default.json"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $projectRoot "run_once.ps1"
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $projectRoot $Config }
$configData = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$popupEnabled = if ($null -eq $configData.schedule.popup_enabled) { $true } else { [bool]$configData.schedule.popup_enabled }
$popupTimeout = if ($null -eq $configData.schedule.popup_timeout_seconds) { 15 } else { [int]$configData.schedule.popup_timeout_seconds }
$popupScript = Join-Path $projectRoot "show_status.ps1"
$archiveRoot = [string]$configData.paths.archive_root
if (-not [System.IO.Path]::IsPathRooted($archiveRoot)) {
    $configBase = Split-Path -Parent (Split-Path -Parent $configPath)
    $archiveRoot = Join-Path $configBase $archiveRoot
}
$logRoot = Join-Path $archiveRoot "scheduler"
$startedAt = Get-Date
$logFile = Join-Path $logRoot ("scheduled_{0}_{1}.log" -f $startedAt.ToString("yyyyMMdd_HHmmss"), $PID)
$collectorExitCode = 3
$collectorOutput = @()

New-Item -ItemType Directory -Path $logRoot -Force | Out-Null

try {
    "开始时间: $($startedAt.ToString('yyyy-MM-dd HH:mm:ss zzz'))" | Set-Content -LiteralPath $logFile -Encoding UTF8
    "配置文件: $configPath" | Add-Content -LiteralPath $logFile -Encoding UTF8
    # Capture output first so the summary can be shown in a popup while retaining
    # the collector's own failure / access-restriction exit code.
    $ErrorActionPreference = "Continue"
    $collectorOutput = @(& $runner -Config $configPath *>&1)
    $collectorOutput | ForEach-Object { $_.ToString() } | Add-Content -LiteralPath $logFile -Encoding UTF8
    $collectorOutput | Out-Host
    $collectorExitCode = $LASTEXITCODE
}
catch {
    $_ | Out-String | Add-Content -LiteralPath $logFile -Encoding UTF8
    $collectorExitCode = 3
}
finally {
    $ErrorActionPreference = "Stop"
    "结束时间: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))" | Add-Content -LiteralPath $logFile -Encoding UTF8
    "退出码: $collectorExitCode" | Add-Content -LiteralPath $logFile -Encoding UTF8
}

if ($popupEnabled -and (Test-Path -LiteralPath $popupScript -PathType Leaf)) {
    $lines = @($collectorOutput | ForEach-Object { $_.ToString() })
    $joined = [string]::Join([Environment]::NewLine, $lines)
    $summary = $null
    $jsonStart = $joined.IndexOf('{')
    $jsonEnd = $joined.LastIndexOf('}')
    if ($jsonStart -ge 0 -and $jsonEnd -gt $jsonStart) {
        try { $summary = $joined.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json } catch { $summary = $null }
    }
    if ($summary) {
        $status = [string]$summary.status
        $title = "贵州省大型水库实时水情：$status"
        $messageLines = @(
            "执行状态：$status",
            "批次：$($summary.batch_id)",
            "全国记录：$($summary.national_count)，贵州记录：$($summary.target_count)",
            "输出记录：$($summary.output_count)",
            "下次查看：$logFile"
        )
        if ($summary.output_file) { $messageLines += "成果文件：$($summary.output_file)" }
        if ($summary.errors -and $summary.errors.Count -gt 0) { $messageLines += "错误：$($summary.errors[0])" }
    }
    else {
        $title = if ($collectorExitCode -eq 0) { "贵州省大型水库实时水情：已完成" } else { "贵州省大型水库实时水情：执行失败" }
        $tail = if ($joined.Length -gt 900) { $joined.Substring($joined.Length - 900) } else { $joined }
        $messageLines = @("退出码：$collectorExitCode", "日志：$logFile", $tail)
    }
    & $popupScript -Title $title -Message ($messageLines -join [Environment]::NewLine) -TimeoutSeconds $popupTimeout -Icon $(if ($collectorExitCode -eq 0) { 64 } else { 16 })
}

exit $collectorExitCode
