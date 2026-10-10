function Get-CollectorDailyTimes {
    param(
        [Parameter(Mandatory = $true)][string]$DailyAt,
        [ValidateRange(1, 24)][int]$DailyCount = 1
    )
    $parsedTime = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($DailyAt, 'HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedTime)) {
        throw '执行时间必须使用24小时制 HH:mm，例如 09:00。'
    }
    $startMinutes = $parsedTime.Hour * 60 + $parsedTime.Minute
    for ($index = 0; $index -lt $DailyCount; $index++) {
        $minuteOfDay = ($startMinutes + [int][math]::Floor($index * 1440.0 / $DailyCount)) % 1440
        '{0:00}:{1:00}' -f [int][math]::Floor($minuteOfDay / 60.0), ($minuteOfDay % 60)
    }
}
