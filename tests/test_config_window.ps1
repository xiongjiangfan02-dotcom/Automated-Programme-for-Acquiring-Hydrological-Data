param([string]$PreviewPath = '')
$ErrorActionPreference = 'Stop'
$testProject = Split-Path -Parent $PSScriptRoot
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('hydro-settings-test-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path (Join-Path $testRoot 'config'), (Join-Path $testRoot 'outputs'), (Join-Path $testRoot 'archive')
$testConfig = Join-Path $testRoot 'config/settings.json'
$utf8 = [System.Text.UTF8Encoding]::new($false)
$originalConfig = Get-Content -LiteralPath (Join-Path $testProject 'config/default.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$originalTemplate = [string]$originalConfig.paths.template
$templateHash = (Get-FileHash -LiteralPath $originalTemplate).Hash

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "测试失败：$Message" }
}
function Assert-Rejected {
    param([scriptblock]$Action, [string]$Message)
    $rejected = $false
    try { & $Action | Out-Null } catch { $rejected = $true }
    Assert-True $rejected $Message
}

try {
    $originalConfig.paths.output_root = 'outputs'
    $originalConfig.paths.archive_root = 'archive'
    $originalConfig | Add-Member -NotePropertyName custom_setting -NotePropertyValue ([pscustomobject]@{ retain = '保留其他配置' })
    [System.IO.File]::WriteAllText($testConfig, ($originalConfig | ConvertTo-Json -Depth 100), $utf8)
    . (Join-Path $testProject 'config_window.ps1') -Config $testConfig -NoShow
    Assert-True ($outputText.Text -eq (Join-Path $testRoot 'outputs')) '相对路径基准'
    Assert-True ($dailyCountNumeric.Value -eq 1) '默认每天一次'
    Assert-True ((@(Get-CollectorDailyTimes '09:00' 2) -join ',') -eq '09:00,21:00') '每天两次'
    Assert-True ((@(Get-CollectorDailyTimes '09:00' 3) -join ',') -eq '09:00,17:00,01:00') '跨午夜'
    Assert-True (@(Get-CollectorDailyTimes '23:59' 24 | Select-Object -Unique).Count -eq 24) '24个唯一时刻'
    Assert-Rejected { Get-CollectorDailyTimes '9:00' 1 } '拒绝不规范时间'
    Assert-Rejected { Get-CollectorDailyTimes '09:00' 25 } '拒绝过大次数'
    $dailyCountNumeric.Value = 24
    Assert-True ($previewLabel.Text.Contains('08:00')) '窗口预览随次数更新'
    $dailyCountNumeric.Value = 2
    $timeText.Text = '10:30'
    Assert-True ($previewLabel.Text.Contains('22:30')) '窗口预览随时间更新'

    # Instantiate and lay out actual WinForms controls without an interactive dialog.
    $form.CreateControl()
    $form.ShowInTaskbar = $false
    $form.Opacity = 0
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $form.PerformLayout()
    foreach ($control in $form.Controls) {
        Assert-True ($control.Left -ge 0 -and $control.Top -ge 0 -and $control.Right -le $form.ClientSize.Width -and $control.Bottom -le $form.ClientSize.Height) "控件越界：$($control.Text)"
    }
    if ($PreviewPath) {
        $bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
        try {
            $form.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
            $bitmap.Save($PreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }
    }
    $form.Hide()
    $initialBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($testConfig))
    $values = Validate-Input
    $null = Save-CollectorConfig $values
    $savedSettings = Get-Content -LiteralPath $testConfig -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($savedSettings.schedule.daily_time -eq '10:30' -and $savedSettings.schedule.daily_count -eq 2) '保存时间及次数'
    Assert-True ($savedSettings.custom_setting.retain -eq '保留其他配置') '保留未知配置字段'
    Assert-True ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes("$testConfig.bak")) -eq $initialBytes) '精确保留旧配置备份'
    $null = Save-CollectorConfig $values
    $beforeFailure = Get-ConfigFingerprint
    $badTemplate = Join-Path $testRoot 'invalid.xlsx'
    [System.IO.File]::WriteAllText($badTemplate, 'This is not an Excel workbook.', $utf8)
    $values.TemplatePath = $badTemplate
    Assert-Rejected { Save-CollectorConfig $values } '无效模板必须拒绝'
    Assert-True ((Get-ConfigFingerprint) -eq $beforeFailure) '模板校验失败不更改配置'
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $testRoot 'config') -Filter '*.pending.*').Count -eq 0) '清理候选配置'
    $values.TemplatePath = $originalTemplate
    $values.DailyTime = '25:00'
    Assert-Rejected { Save-CollectorConfig $values } '无效时间必须拒绝'
    Assert-True ((Get-ConfigFingerprint) -eq $beforeFailure) '输入失败不更改配置'
    $values.DailyTime = '09:00'
    [System.IO.File]::AppendAllText($testConfig, "`n", $utf8)
    $externalHash = Get-ConfigFingerprint
    Assert-Rejected { Save-CollectorConfig $values } '阻止覆盖外部修改'
    Assert-True ((Get-ConfigFingerprint) -eq $externalHash) '保留外部修改'
    Assert-True ((Get-FileHash -LiteralPath $originalTemplate).Hash -eq $templateHash) '正式模板字节不变'
    Write-Output 'PASS: scheduling preview, layout, validated atomic saves, backup, invalid input, concurrency, template preservation.'
}
finally {
    if ($form) { $form.Dispose() }
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    $allowedPrefix = Join-Path ([System.IO.Path]::GetTempPath()) 'hydro-settings-test-'
    if (-not $resolvedTestRoot.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw '拒绝清理预期临时目录外的路径。' }
    Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
}
