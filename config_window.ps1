[CmdletBinding()]
param(
    [string]$Config = 'config/default.json',
    [switch]$NoShow
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { [System.IO.Path]::GetFullPath($Config) } else { [System.IO.Path]::GetFullPath((Join-Path $projectRoot $Config)) }
$configBase = Split-Path -Parent (Split-Path -Parent $configPath)
. (Join-Path $projectRoot 'schedule_helpers.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Resolve-ConfigPathValue {
    param([string]$Value)
    if ([System.IO.Path]::IsPathRooted($Value)) { return [System.IO.Path]::GetFullPath($Value) }
    return [System.IO.Path]::GetFullPath((Join-Path $configBase $Value))
}

function Read-CollectorConfig {
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw "配置文件不存在：$configPath" }
    return Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Get-ConfigFingerprint {
    return (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
}

function Invoke-ConfigProcess {
    param([string]$Executable, [string[]]$Arguments)
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Executable
    # Windows paths cannot contain quotes. Double trailing backslashes for the
    # Windows command-line parser; invoke directly without another shell.
    $startInfo.Arguments = (($Arguments | ForEach-Object { '"' + ($_ -replace '(\\+)$', '$1$1') + '"' }) -join ' ')
    $startInfo.WorkingDirectory = $projectRoot
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $startInfo.EnvironmentVariables['PYTHONPATH'] = Join-Path $projectRoot 'src'
    $startInfo.EnvironmentVariables['PYTHONIOENCODING'] = 'utf-8'
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            throw '校验或定时任务操作超过60秒，已停止。请查看任务状态后重试。'
        }
        $output = ($stdout.GetAwaiter().GetResult() + [Environment]::NewLine + $stderr.GetAwaiter().GetResult()).Trim()
        if ($process.ExitCode -ne 0) { throw "操作失败（退出码 $($process.ExitCode)）：`r`n$output" }
        return $output
    }
    finally { $process.Dispose() }
}

function Validate-CollectorValues {
    param([hashtable]$Values)
    $null = @(Get-CollectorDailyTimes -DailyAt $Values.DailyTime -DailyCount $Values.DailyCount)
    if ($Values.Attempts -lt 1 -or $Values.Attempts -gt 3) { throw '单个请求最多尝试次数必须在1到3之间。' }
    foreach ($key in @('OutputPath', 'ArchivePath', 'TemplatePath')) {
        if ([string]::IsNullOrWhiteSpace($Values[$key])) { throw '存储目录和模板位置不能为空。' }
        $Values[$key] = Resolve-ConfigPathValue $Values[$key]
    }
    foreach ($item in @(@('成果存储', $Values.OutputPath), @('归档存储', $Values.ArchivePath))) {
        if ([string]::IsNullOrWhiteSpace($item[1]) -or -not (Test-Path -LiteralPath $item[1] -PathType Container)) {
            throw "$($item[0])目录不存在，请先创建或通过浏览按钮选择：$($item[1])"
        }
    }
    if (-not (Test-Path -LiteralPath $Values.TemplatePath -PathType Leaf) -or [System.IO.Path]::GetExtension($Values.TemplatePath) -ine '.xlsx') {
        throw '请选择存在的 .xlsx 模板文件。'
    }
}

function Validate-Input {
    $values = @{
        DailyTime = $timeText.Text.Trim()
        DailyCount = [int]$dailyCountNumeric.Value
        Attempts = [int]$attemptsNumeric.Value
        OutputPath = $outputText.Text.Trim()
        ArchivePath = $archiveText.Text.Trim()
        TemplatePath = $templateText.Text.Trim()
    }
    Validate-CollectorValues $values
    return $values
}

function Save-CollectorConfig {
    param([hashtable]$Values)
    Validate-CollectorValues $Values
    if ((Get-ConfigFingerprint) -ne $script:loadedFingerprint) { throw '配置文件已被其他窗口或程序修改，请关闭并重新打开本窗口后再保存。' }
    $settings = Read-CollectorConfig
    $settings.source.max_attempts = $Values.Attempts
    $settings.paths.template = (Resolve-ConfigPathValue $Values.TemplatePath).Replace('\', '/')
    $settings.paths.output_root = (Resolve-ConfigPathValue $Values.OutputPath).Replace('\', '/')
    $settings.paths.archive_root = (Resolve-ConfigPathValue $Values.ArchivePath).Replace('\', '/')
    if ($null -eq $settings.schedule) { $settings | Add-Member -NotePropertyName schedule -NotePropertyValue ([pscustomobject]@{}) -Force }
    $settings.schedule | Add-Member -NotePropertyName daily_time -NotePropertyValue $Values.DailyTime -Force
    $settings.schedule | Add-Member -NotePropertyName daily_count -NotePropertyValue $Values.DailyCount -Force
    $backupPath = "$configPath.bak"
    # Keep the candidate in the config directory so relative paths retain their base.
    $candidatePath = "$configPath.pending.$([guid]::NewGuid().ToString('N')).json"
    try {
        [System.IO.File]::WriteAllText($candidatePath, ($settings | ConvertTo-Json -Depth 100), [System.Text.UTF8Encoding]::new($false))
        $python = Join-Path $projectRoot '.venv/Scripts/python.exe'
        if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw '项目Python环境不存在，请先运行 setup.ps1。' }
        $null = Invoke-ConfigProcess $python @('-m', 'bijie_hydro.cli', '--config', $candidatePath, 'validate-template')
        if ((Get-ConfigFingerprint) -ne $script:loadedFingerprint) { throw '校验期间配置被其他程序修改，本次未保存。请重新打开窗口。' }
        [System.IO.File]::Replace($candidatePath, $configPath, $backupPath)
        $script:loadedFingerprint = Get-ConfigFingerprint
    }
    finally {
        if (Test-Path -LiteralPath $candidatePath) { Remove-Item -LiteralPath $candidatePath -Force }
    }
    return $backupPath
}

function Update-ScheduledTask {
    $powershell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $manager = Join-Path $projectRoot 'manage_schedule.ps1'
    $null = Invoke-ConfigProcess $powershell @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $manager, '-Action', 'Install', '-Config', $configPath, '-Force')
}

try {
    $script:loadedFingerprint = Get-ConfigFingerprint
    $settings = Read-CollectorConfig
}
catch {
    if ($NoShow) { throw }
    [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '无法打开配置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '贵州省大型水库实时水情 - 采集配置'
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.ClientSize = New-Object System.Drawing.Size(800, 630)
$form.StartPosition = 'CenterScreen'
$form.MaximizeBox = $false
$form.FormBorderStyle = 'FixedDialog'

function Add-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    [void]$form.Controls.Add($label)
    return $label
}

$title = Add-Label '采集参数设置' 24 18 740 32
$title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 16, [System.Drawing.FontStyle]::Bold)
$hint = Add-Label '设置执行时刻、采集频次和文件位置；保存时只读校验模板。' 26 58 740 26
$hint.ForeColor = [System.Drawing.Color]::DimGray

function Add-Field {
    param([string]$Label, [int]$Top, [System.Windows.Forms.Control]$Control, [string]$ButtonText = '')
    $null = Add-Label $Label 28 $Top 218 28
    $Control.Location = New-Object System.Drawing.Point(250, ($Top - 2))
    $Control.Size = New-Object System.Drawing.Size(420, 28)
    [void]$form.Controls.Add($Control)
    if ($ButtonText) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $ButtonText
        $button.Location = New-Object System.Drawing.Point(684, ($Top - 3))
        $button.Size = New-Object System.Drawing.Size(88, 30)
        [void]$form.Controls.Add($button)
        return $button
    }
}

$timeText = New-Object System.Windows.Forms.TextBox
$timeText.Text = if ($settings.schedule.daily_time) { [string]$settings.schedule.daily_time } else { '09:00' }
Add-Field '起始执行时间（HH:mm）' 104 $timeText
$dailyCountNumeric = New-Object System.Windows.Forms.NumericUpDown
$dailyCountNumeric.Minimum = 1
$dailyCountNumeric.Maximum = 24
$dailyCountNumeric.Value = if ($null -ne $settings.schedule.daily_count) { [decimal]$settings.schedule.daily_count } else { 1 }
Add-Field '每天采集次数（1～24）' 146 $dailyCountNumeric
$previewLabel = Add-Label '' 250 184 520 74
$previewLabel.ForeColor = [System.Drawing.Color]::SteelBlue
function Update-TimePreview {
    try {
        $times = @(Get-CollectorDailyTimes $timeText.Text.Trim() ([int]$dailyCountNumeric.Value))
        $previewLabel.Text = '北京时间，每日执行：' + ($times -join '、')
    }
    catch { $previewLabel.Text = '请输入24小时制时间，例如 09:00。' }
}
$timeText.Add_TextChanged({ Update-TimePreview })
$dailyCountNumeric.Add_ValueChanged({ Update-TimePreview })
Update-TimePreview

$attemptsNumeric = New-Object System.Windows.Forms.NumericUpDown
$attemptsNumeric.Minimum = 1
$attemptsNumeric.Maximum = 3
$attemptsNumeric.Value = if ($null -ne $settings.source.max_attempts) { [decimal]$settings.source.max_attempts } else { 3 }
Add-Field '单个请求最多尝试（含首次）' 272 $attemptsNumeric

function Select-Folder {
    param([System.Windows.Forms.TextBox]$Target, [string]$Description)
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    try {
        $dialog.Description = $Description
        $dialog.SelectedPath = $Target.Text
        $dialog.ShowNewFolderButton = $true
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $Target.Text = $dialog.SelectedPath }
    }
    finally { $dialog.Dispose() }
}

$outputText = New-Object System.Windows.Forms.TextBox
$outputText.Text = Resolve-ConfigPathValue ([string]$settings.paths.output_root)
$outputBrowse = Add-Field '成果 Excel 存储位置' 316 $outputText '浏览...'
$outputBrowse.Add_Click({ Select-Folder $outputText '选择成果目录，可新建文件夹' })
$archiveText = New-Object System.Windows.Forms.TextBox
$archiveText.Text = Resolve-ConfigPathValue ([string]$settings.paths.archive_root)
$archiveBrowse = Add-Field '原始数据、日志归档位置' 358 $archiveText '浏览...'
$archiveBrowse.Add_Click({ Select-Folder $archiveText '选择归档目录，可新建文件夹' })
$templateText = New-Object System.Windows.Forms.TextBox
$templateText.Text = Resolve-ConfigPathValue ([string]$settings.paths.template)
$templateBrowse = Add-Field 'Excel 模板位置' 400 $templateText '浏览...'
$templateBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialog.Title = '选择 Excel 模板（只读校验）'
        $dialog.Filter = 'Excel 工作簿 (*.xlsx)|*.xlsx'
        $dialog.FileName = $templateText.Text
        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $templateText.Text = $dialog.FileName }
    }
    finally { $dialog.Dispose() }
})
$pathHint = Add-Label '目录须存在，可通过“浏览”新建。修改位置只影响后续文件；历史文件不迁移。字体缓存和去重状态继续使用原配置位置。' 28 444 742 48
$pathHint.ForeColor = [System.Drawing.Color]::DimGray
$statusLabel = Add-Label '就绪。更改时间或每天次数后，请使用“保存并更新任务”。' 28 504 742 48

function Set-Status {
    param([string]$Text, [bool]$IsError = $false)
    $statusLabel.Text = $Text
    $statusLabel.ForeColor = if ($IsError) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::DarkGreen }
    $statusLabel.Refresh()
}

function Add-ActionButton {
    param([string]$Text, [int]$X, [int]$Width)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, 572)
    $button.Size = New-Object System.Drawing.Size($Width, 36)
    [void]$form.Controls.Add($button)
    return $button
}
$saveButton = Add-ActionButton '保存配置' 354 108
$applyButton = Add-ActionButton '保存并更新任务' 476 172
$closeButton = Add-ActionButton '关闭' 662 110
$closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $closeButton

function Save-FromForm {
    param([bool]$ApplyTask)
    $saved = $false
    $saveButton.Enabled = $false
    $applyButton.Enabled = $false
    $closeButton.Enabled = $false
    $form.UseWaitCursor = $true
    try {
        $values = Validate-Input
        Set-Status '正在校验配置及模板，通过后保存...'
        $backup = Save-CollectorConfig $values
        $saved = $true
        if ($ApplyTask) {
            Set-Status '配置已保存，正在更新 Windows 定时任务...'
            Update-ScheduledTask
            Set-Status ('已保存并更新任务，每天{0}次。备份：{1}' -f $values.DailyCount, $backup)
            $message = '配置及定时任务已更新。执行时刻：' + (@(Get-CollectorDailyTimes $values.DailyTime $values.DailyCount) -join '、')
        }
        else {
            Set-Status "已保存配置。更改时间或次数需更新任务。备份：$backup"
            $message = '配置已保存。时间或每天次数的更改，需点击“保存并更新任务”才会应用到 Windows 任务计划。'
        }
        [void][System.Windows.Forms.MessageBox]::Show($form, $message, '完成', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        $message = if ($saved) { "配置已保存，但未确认定时任务更新成功。请检查任务状态后重试。`r`n$($_.Exception.Message)" } else { "配置未保存。`r`n$($_.Exception.Message)" }
        Set-Status $message $true
        [void][System.Windows.Forms.MessageBox]::Show($form, $message, '操作失败', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $form.UseWaitCursor = $false
        $saveButton.Enabled = $true
        $applyButton.Enabled = $true
        $closeButton.Enabled = $true
    }
}
$saveButton.Add_Click({ Save-FromForm $false })
$applyButton.Add_Click({ Save-FromForm $true })
$form.Add_Shown({ $form.Activate() })
if (-not $NoShow) {
    try { [void]$form.ShowDialog() }
    finally { $form.Dispose() }
}
