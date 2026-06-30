param(
  [string]$Url = "http://192.168.68.104/cellmonitor",
  [string]$MainUrl = "",
  [int]$IntervalSeconds = 60,
  [string]$OutDir = ".\cellmonitor-logs",
  [int]$MaxCells = 0,
  [string]$InputFile = "",
  [string]$MainInputFile = "",
  [string]$CanLogUrl = "",
  [string]$CanLogInputFile = "",
  [switch]$UseCanLog,
  [switch]$Once
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$CompactCsv = Join-Path $OutDir "battery_log.csv"
$EventsCsv = Join-Path $OutDir "battery_events.csv"
$RjxzsCanCsv = Join-Path $OutDir "rjxzs_can_status.csv"
$ErrorLog = Join-Path $OutDir "cellmonitor_errors.log"
$script:LastEventKey = ""
$script:LastOpenMosEventAt = $null

function Get-MainPageUrl {
  if ($MainUrl -ne "") {
    return $MainUrl
  }

  $builder = [System.UriBuilder]::new($Url)
  $builder.Path = "/"
  $builder.Query = ""
  $builder.Fragment = ""
  return $builder.Uri.AbsoluteUri
}

function Get-CanLogPageUrl {
  if ($CanLogUrl -ne "") {
    return $CanLogUrl
  }

  $builder = [System.UriBuilder]::new($Url)
  $builder.Path = "/canlog"
  $builder.Query = ""
  $builder.Fragment = ""
  return $builder.Uri.AbsoluteUri
}

function Get-PageContent {
  param(
    [string]$Url,
    [string]$InputFile
  )

  if ($InputFile -ne "") {
    return Get-Content -Raw -Path $InputFile
  }

  $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 20
  return $response.Content
}

function Convert-HtmlToText {
  param([string]$Html)

  $text = $Html -replace "(?i)<br\s*/?>", "`n"
  $text = $text -replace "(?i)</h[1-6]>", "`n"
  $text = $text -replace "(?i)</div>", "`n"
  $text = $text -replace "<[^>]+>", " "
  $text = [System.Net.WebUtility]::HtmlDecode($text)
  $text = $text -replace "\u00a0", " "
  $text = $text -replace "[ \t]+", " "
  $text = $text -replace "\r", ""
  return $text.Trim()
}

function Get-NumberMatch {
  param(
    [string]$Text,
    [string]$Pattern
  )

  $match = [regex]::Match($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if (-not $match.Success) {
    return $null
  }

  return [double]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-TextMatch {
  param(
    [string]$Text,
    [string]$Pattern
  )

  $match = [regex]::Match($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if (-not $match.Success) {
    return ""
  }

  return $match.Groups[1].Value.Trim()
}

function Get-JsArrayItems {
  param(
    [string]$Html,
    [string]$ArrayName
  )

  $pattern = "const\s+" + [regex]::Escape($ArrayName) + "\s*=\s*\[(.*?)\];"
  $match = [regex]::Match(
    $Html,
    $pattern,
    [System.Text.RegularExpressions.RegexOptions]::Singleline
  )

  if (-not $match.Success) {
    throw "Could not find JavaScript array '$ArrayName' in cellmonitor HTML"
  }

  return $match.Groups[1].Value -split "," |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -ne "" }
}

function Convert-PowerToW {
  param(
    [Nullable[double]]$Value,
    [string]$Unit
  )

  if ($null -eq $Value) {
    return $null
  }

  if ($Unit -ieq "kW") {
    return [int]($Value * 1000)
  }

  return [int]$Value
}

function Get-MainSample {
  param([string]$Html)

  $text = Convert-HtmlToText -Html $Html
  $socPct = Get-NumberMatch -Text $text -Pattern "SOC:\s*([-+]?\d+(?:\.\d+)?)\s*%"
  $totalCapacityKwh = Get-NumberMatch -Text $text -Pattern "Total capacity:\s*([-+]?\d+(?:\.\d+)?)\s*kWh"
  $remainingCapacityKwh = Get-NumberMatch -Text $text -Pattern "Remaining capacity:\s*([-+]?\d+(?:\.\d+)?)\s*kWh"
  $powerValue = Get-NumberMatch -Text $text -Pattern "Power:\s*([-+]?\d+(?:\.\d+)?)\s*(kW|W)"
  $powerUnit = Get-TextMatch -Text $text -Pattern "Power:\s*[-+]?\d+(?:\.\d+)?\s*(kW|W)"
  $maxChargeValue = Get-NumberMatch -Text $text -Pattern "Max charge power:\s*([-+]?\d+(?:\.\d+)?)\s*(kW|W)"
  $maxChargeUnit = Get-TextMatch -Text $text -Pattern "Max charge power:\s*[-+]?\d+(?:\.\d+)?\s*(kW|W)"
  $maxDischargeValue = Get-NumberMatch -Text $text -Pattern "Max discharge power:\s*([-+]?\d+(?:\.\d+)?)\s*(kW|W)"
  $maxDischargeUnit = Get-TextMatch -Text $text -Pattern "Max discharge power:\s*[-+]?\d+(?:\.\d+)?\s*(kW|W)"

  if ($null -eq $socPct -and $null -ne $remainingCapacityKwh -and $null -ne $totalCapacityKwh -and $totalCapacityKwh -gt 0) {
    $socPct = [Math]::Round(($remainingCapacityKwh / $totalCapacityKwh) * 100, 2)
  }

  return [pscustomobject]@{
    Software = Get-TextMatch -Text $text -Pattern "Software:\s*([^\s]+)"
    BoardTemperatureC = Get-NumberMatch -Text $text -Pattern "Hardware:.*?@\s*([-+]?\d+(?:\.\d+)?)\s*[^0-9\r\n]*C"
    SocPct = $socPct
    SohPct = Get-NumberMatch -Text $text -Pattern "SOH:\s*([-+]?\d+(?:\.\d+)?)\s*%"
    VoltageV = Get-NumberMatch -Text $text -Pattern "Voltage:\s*([-+]?\d+(?:\.\d+)?)\s*V"
    CurrentA = Get-NumberMatch -Text $text -Pattern "Current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    PowerW = Convert-PowerToW -Value $powerValue -Unit $powerUnit
    TotalCapacityKwh = $totalCapacityKwh
    RemainingCapacityKwh = $remainingCapacityKwh
    MaxDischargePowerW = Convert-PowerToW -Value $maxDischargeValue -Unit $maxDischargeUnit
    MaxChargePowerW = Convert-PowerToW -Value $maxChargeValue -Unit $maxChargeUnit
    MaxDischargeCurrentA = Get-NumberMatch -Text $text -Pattern "Max discharge current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    MaxChargeCurrentA = Get-NumberMatch -Text $text -Pattern "Max charge current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    TemperatureMinC = Get-NumberMatch -Text $text -Pattern "Temperature min/max:\s*([-+]?\d+(?:\.\d+)?)\s*[^/\r\n]*C"
    TemperatureMaxC = Get-NumberMatch -Text $text -Pattern "Temperature min/max:\s*[-+]?\d+(?:\.\d+)?\s*[^/\r\n]*C\s*/\s*([-+]?\d+(?:\.\d+)?)\s*[^0-9\r\n]*C"
    BatteryState = Get-TextMatch -Text $text -Pattern "(Battery (?:idle|charging[^\r\n]*|discharging[^\r\n]*))"
    SystemStatus = Get-TextMatch -Text $text -Pattern "System status:\s*([^\r\n]+)"
    PowerStatus = Get-TextMatch -Text $text -Pattern "Power status:\s*([^\r\n]+)"
    RjxzsChargeMos = Get-TextMatch -Text $text -Pattern "RJXZS BMS:\s*Charge MOS:\s*([^|\r\n]+)"
    RjxzsDischargeMos = Get-TextMatch -Text $text -Pattern "Discharge MOS:\s*([^|\r\n]+)"
    RjxzsDefaultChannel = Get-TextMatch -Text $text -Pattern "Default channel:\s*([^|\r\n]+)"
    RjxzsLog = Get-TextMatch -Text $text -Pattern "Log:\s*([^\r\n]+)"
  }
}

function Get-CleanText {
  param([object]$Value)

  if ($null -eq $Value) {
    return ""
  }

  return ([string]$Value).Trim()
}

function Test-MosOpen {
  param([string]$MosStatus)

  return (Get-CleanText $MosStatus).ToUpperInvariant() -eq "OFF"
}

function Get-RecoveryObservation {
  param(
    [pscustomobject]$CellSample,
    [pscustomobject]$MainSample,
    [pscustomobject]$RjxzsCanStatus
  )

  if ($null -eq $MainSample) {
    return "main_page_unavailable"
  }

  $chargeMos = Get-CleanText $MainSample.RjxzsChargeMos
  $dischargeMos = Get-CleanText $MainSample.RjxzsDischargeMos
  if ($chargeMos -eq "" -and $null -ne $RjxzsCanStatus) {
    $chargeMos = Get-CleanText $RjxzsCanStatus.ChargeMos
  }
  if ($dischargeMos -eq "" -and $null -ne $RjxzsCanStatus) {
    $dischargeMos = Get-CleanText $RjxzsCanStatus.DischargeMos
  }

  $chargeOpen = Test-MosOpen $chargeMos
  $dischargeOpen = Test-MosOpen $dischargeMos
  if (-not $chargeOpen -and -not $dischargeOpen) {
    return "mos_closed"
  }

  $blocks = @()
  $rjxzsLog = Get-CleanText $MainSample.RjxzsLog
  if ($rjxzsLog -eq "" -and $null -ne $RjxzsCanStatus) {
    $rjxzsLog = Get-CleanText $RjxzsCanStatus.HistoricalLog
  }
  $systemStatus = Get-CleanText $MainSample.SystemStatus
  $activeProtections = if ($null -ne $RjxzsCanStatus) { Get-CleanText $RjxzsCanStatus.ActiveProtections } else { "" }

  if ($activeProtections -ne "" -and $activeProtections -ne "None") {
    $blocks += "active_protection=$activeProtections"
  }

  if ($rjxzsLog -ne "" -and $rjxzsLog -ne "None") {
    $blocks += "rjxzs_log=$rjxzsLog"
  }

  if ($systemStatus -ne "" -and $systemStatus -ne "OK") {
    $blocks += "system_status=$systemStatus"
  }

  if ($CellSample.DeltaMv -gt 100) {
    $blocks += "cell_delta_gt_100mV"
  }

  if ($blocks.Count -eq 0) {
    return "mos_open_no_visible_block"
  }

  return $blocks -join ";"
}

function Write-EventIfNeeded {
  param(
    [pscustomobject]$CellSample,
    [pscustomobject]$MainSample,
    [pscustomobject]$RjxzsCanStatus
  )

  if ($null -eq $MainSample) {
    return
  }

  $chargeMos = Get-CleanText $MainSample.RjxzsChargeMos
  $dischargeMos = Get-CleanText $MainSample.RjxzsDischargeMos
  $defaultChannel = Get-CleanText $MainSample.RjxzsDefaultChannel
  $rjxzsLog = Get-CleanText $MainSample.RjxzsLog
  if ($chargeMos -eq "" -and $null -ne $RjxzsCanStatus) {
    $chargeMos = Get-CleanText $RjxzsCanStatus.ChargeMos
  }
  if ($dischargeMos -eq "" -and $null -ne $RjxzsCanStatus) {
    $dischargeMos = Get-CleanText $RjxzsCanStatus.DischargeMos
  }
  if ($defaultChannel -eq "" -and $null -ne $RjxzsCanStatus) {
    $defaultChannel = Get-CleanText $RjxzsCanStatus.DefaultChannel
  }
  if ($rjxzsLog -eq "" -and $null -ne $RjxzsCanStatus) {
    $rjxzsLog = Get-CleanText $RjxzsCanStatus.HistoricalLog
  }
  $systemStatus = Get-CleanText $MainSample.SystemStatus
  $powerStatus = Get-CleanText $MainSample.PowerStatus
  $batteryState = Get-CleanText $MainSample.BatteryState
  $activeStatusHex = if ($null -ne $RjxzsCanStatus) { Get-CleanText $RjxzsCanStatus.StatusAccountingHex } else { "" }
  $activeProtections = if ($null -ne $RjxzsCanStatus) { Get-CleanText $RjxzsCanStatus.ActiveProtections } else { "" }
  $eventKey = "$chargeMos|$dischargeMos|$defaultChannel|$rjxzsLog|$systemStatus|$powerStatus|$batteryState|$activeStatusHex|$activeProtections"

  $chargeOpen = Test-MosOpen $chargeMos
  $dischargeOpen = Test-MosOpen $dischargeMos
  $mosOpen = $chargeOpen -or $dischargeOpen
  $now = Get-Date
  $reasons = @()

  if ($script:LastEventKey -eq "") {
    $reasons += "first_sample"
  } elseif ($eventKey -ne $script:LastEventKey) {
    $reasons += "state_change"
  }

  if ($mosOpen) {
    $reasons += "mos_open"
    if ($reasons.Count -eq 1 -and $null -ne $script:LastOpenMosEventAt -and
        ($now - $script:LastOpenMosEventAt).TotalMinutes -ge 15) {
      $reasons += "mos_open_15min"
    }
  }

  if ($rjxzsLog -ne "" -and $rjxzsLog -ne "None") {
    $reasons += "rjxzs_log"
  }

  if ($activeProtections -ne "" -and $activeProtections -ne "None") {
    $reasons += "active_protection"
  }

  if ($systemStatus -ne "" -and $systemStatus -ne "OK") {
    $reasons += "system_not_ok"
  }

  if ($reasons.Count -eq 0) {
    $script:LastEventKey = $eventKey
    return
  }

  [pscustomobject][ordered]@{
    timestamp = $CellSample.Timestamp
    reason = ($reasons | Select-Object -Unique) -join ";"
    recovery_observation = Get-RecoveryObservation -CellSample $CellSample -MainSample $MainSample -RjxzsCanStatus $RjxzsCanStatus
    rjxzs_charge_mos = $chargeMos
    rjxzs_discharge_mos = $dischargeMos
    rjxzs_default_channel = $defaultChannel
    rjxzs_log = $rjxzsLog
    active_status_hex = $activeStatusHex
    active_protections = $activeProtections
    active_channel_status_bit = if ($null -ne $RjxzsCanStatus) { $RjxzsCanStatus.ChannelStatusBit } else { "" }
    active_current_polarity = if ($null -ne $RjxzsCanStatus) { $RjxzsCanStatus.CurrentPolarity } else { "" }
    system_status = $systemStatus
    power_status = $powerStatus
    battery_state = $batteryState
    current_a = $MainSample.CurrentA
    voltage_v = $MainSample.VoltageV
    power_w = $MainSample.PowerW
    soc_pct = $MainSample.SocPct
    min_mv = $CellSample.MinMv
    max_mv = $CellSample.MaxMv
    delta_mv = $CellSample.DeltaMv
    min_cell = $CellSample.MinCell
    max_cell = $CellSample.MaxCell
    temperature_min_c = $MainSample.TemperatureMinC
    temperature_max_c = $MainSample.TemperatureMaxC
  } | Export-Csv -Path $EventsCsv -NoTypeInformation -Append -Encoding UTF8

  $script:LastEventKey = $eventKey
  if ($mosOpen) {
    $script:LastOpenMosEventAt = $now
  } else {
    $script:LastOpenMosEventAt = $null
  }
}

function Get-ActiveProtectionText {
  param([int]$StatusAccounting)

  $items = @()
  if (($StatusAccounting -band 0x001) -ne 0) { $items += "Over temperature" }
  if (($StatusAccounting -band 0x002) -ne 0) { $items += "Overcharge" }
  if (($StatusAccounting -band 0x004) -ne 0) { $items += "Battery string error" }
  if (($StatusAccounting -band 0x008) -ne 0) { $items += "Overcurrent" }
  if (($StatusAccounting -band 0x010) -ne 0) { $items += "Overdischarge" }

  if ($items.Count -eq 0) {
    return "None"
  }

  return $items -join ";"
}

function Get-RjxzsHistoricalLogText {
  param([int]$Code)

  switch ($Code) {
    0x00 { return "None" }
    0x01 { return "Overcurrent protection" }
    0x02 { return "Overdischarge protection" }
    0x03 { return "Overcharge protection" }
    0x04 { return "Over temperature protection" }
    0x05 { return "Battery string error protection" }
    0x06 { return "Damaged charging relay" }
    0x07 { return "Damaged discharge relay" }
    0x08 { return "Low voltage power outage protection" }
    0x09 { return "Voltage difference protection" }
    0x0A { return "Low temperature protection" }
    default { return "Unknown log code" }
  }
}

function Get-RjxzsDefaultChannelText {
  param([int]$Code)

  switch ($Code) {
    0x01 { return "ON after power-on" }
    0x02 { return "OFF after power-on" }
    default { return "Unknown" }
  }
}

function Test-RjxzsChargeAllowed {
  param([object]$MosStatus)

  if ($null -eq $MosStatus) {
    return $null
  }

  return (([int]$MosStatus -band 0x02) -ne 0)
}

function Test-RjxzsDischargeAllowed {
  param([object]$MosStatus)

  if ($null -eq $MosStatus) {
    return $null
  }

  return (([int]$MosStatus -band 0x01) -ne 0)
}

function Get-RjxzsMosText {
  param([object]$Allowed)

  if ($null -eq $Allowed) {
    return ""
  }

  if ($Allowed) {
    return "ON"
  }

  return "OFF"
}

function Get-RjxzsCanStatusSample {
  param(
    [string]$Html,
    [string]$Timestamp
  )

  $text = Convert-HtmlToText -Html $Html
  $matches = [regex]::Matches(
    $text,
    "(?im)\bRX\d+\s+F5\s+\[(\d+)\]\s+((?:[0-9A-F]{2}\s*){8})"
  )

  $statusBytes = $null
  $mosBytes = $null
  $logBytes = $null
  $defaultChannelBytes = $null
  foreach ($match in $matches) {
    $bytes = @($match.Groups[2].Value -split "\s+" | Where-Object { $_ -ne "" })
    if ($bytes.Count -ge 8 -and $bytes[0].ToUpperInvariant() -eq "06") {
      $statusBytes = $bytes
    } elseif ($bytes.Count -ge 8 -and $bytes[0].ToUpperInvariant() -eq "51") {
      $mosBytes = $bytes
    } elseif ($bytes.Count -ge 8 -and $bytes[0].ToUpperInvariant() -eq "53") {
      $logBytes = $bytes
    } elseif ($bytes.Count -ge 8 -and $bytes[0].ToUpperInvariant() -eq "54") {
      $defaultChannelBytes = $bytes
    }
  }

  if ($null -eq $statusBytes -and $null -eq $mosBytes -and $null -eq $logBytes -and $null -eq $defaultChannelBytes) {
    return $null
  }

  $statusValues = if ($null -ne $statusBytes) { @($statusBytes | ForEach-Object { [Convert]::ToInt32($_, 16) }) } else { $null }
  $mosValues = if ($null -ne $mosBytes) { @($mosBytes | ForEach-Object { [Convert]::ToInt32($_, 16) }) } else { $null }
  $logValues = if ($null -ne $logBytes) { @($logBytes | ForEach-Object { [Convert]::ToInt32($_, 16) }) } else { $null }
  $defaultChannelValues = if ($null -ne $defaultChannelBytes) { @($defaultChannelBytes | ForEach-Object { [Convert]::ToInt32($_, 16) }) } else { $null }
  $hostTemperatureRaw = if ($null -ne $statusValues) { ($statusValues[1] -shl 8) -bor $statusValues[2] } else { $null }
  $statusAccounting = if ($null -ne $statusValues) { ($statusValues[3] -shl 8) -bor $statusValues[4] } else { $null }
  $equalizationStartingVoltage = if ($null -ne $statusValues) { ($statusValues[5] -shl 8) -bor $statusValues[6] } else { $null }
  $mosStatus = if ($null -ne $mosValues) { $mosValues[7] } else { $null }
  $historicalLogCode = if ($null -ne $logValues) { $logValues[7] } else { $null }
  $defaultChannelCode = if ($null -ne $defaultChannelValues) { $defaultChannelValues[7] } else { $null }
  $chargeAllowed = Test-RjxzsChargeAllowed -MosStatus $mosStatus
  $dischargeAllowed = Test-RjxzsDischargeAllowed -MosStatus $mosStatus

  return [pscustomobject]@{
    Timestamp = $Timestamp
    SourceUrl = Get-CanLogPageUrl
    StatusFrameSeen = ($null -ne $statusBytes)
    MosFrameSeen = ($null -ne $mosBytes)
    HistoricalLogFrameSeen = ($null -ne $logBytes)
    DefaultChannelFrameSeen = ($null -ne $defaultChannelBytes)
    RawFrame = if ($null -ne $statusBytes) { ($statusBytes -join " ") } else { "" }
    MosRawFrame = if ($null -ne $mosBytes) { ($mosBytes -join " ") } else { "" }
    HistoricalLogRawFrame = if ($null -ne $logBytes) { ($logBytes -join " ") } else { "" }
    DefaultChannelRawFrame = if ($null -ne $defaultChannelBytes) { ($defaultChannelBytes -join " ") } else { "" }
    HostTemperatureRaw = $hostTemperatureRaw
    StatusAccounting = $statusAccounting
    StatusAccountingHex = if ($null -ne $statusAccounting) { "0x{0:X4}" -f $statusAccounting } else { "" }
    HostTemperatureBelowZero = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x100) -ne 0) } else { "" }
    ChannelStatusBit = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x080) -ne 0) } else { "" }
    CurrentPolarity = if ($null -ne $statusAccounting) { if (($statusAccounting -band 0x040) -ne 0) { "charge" } else { "discharge" } } else { "" }
    Balancing = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x020) -ne 0) } else { "" }
    Overdischarge = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x010) -ne 0) } else { "" }
    Overcurrent = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x008) -ne 0) } else { "" }
    BatteryStringError = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x004) -ne 0) } else { "" }
    Overcharge = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x002) -ne 0) } else { "" }
    Overtemperature = if ($null -ne $statusAccounting) { (($statusAccounting -band 0x001) -ne 0) } else { "" }
    ActiveProtections = if ($null -ne $statusAccounting) { Get-ActiveProtectionText -StatusAccounting $statusAccounting } else { "" }
    MosStatusHex = if ($null -ne $mosStatus) { "0x{0:X2}" -f $mosStatus } else { "" }
    ChargeAllowed = $chargeAllowed
    DischargeAllowed = $dischargeAllowed
    ChargeMos = Get-RjxzsMosText -Allowed $chargeAllowed
    DischargeMos = Get-RjxzsMosText -Allowed $dischargeAllowed
    HistoricalLogCodeHex = if ($null -ne $historicalLogCode) { "0x{0:X2}" -f $historicalLogCode } else { "" }
    HistoricalLog = if ($null -ne $historicalLogCode) { Get-RjxzsHistoricalLogText -Code $historicalLogCode } else { "" }
    DefaultChannelCodeHex = if ($null -ne $defaultChannelCode) { "0x{0:X2}" -f $defaultChannelCode } else { "" }
    DefaultChannel = if ($null -ne $defaultChannelCode) { Get-RjxzsDefaultChannelText -Code $defaultChannelCode } else { "" }
    EqualizationStartingVoltageMv = $equalizationStartingVoltage
  }
}

function Write-RjxzsCanStatus {
  param([pscustomobject]$RjxzsCanStatus)

  if ($null -eq $RjxzsCanStatus) {
    return
  }

  [pscustomobject][ordered]@{
    timestamp = $RjxzsCanStatus.Timestamp
    source_url = $RjxzsCanStatus.SourceUrl
    status_frame_seen = $RjxzsCanStatus.StatusFrameSeen
    mos_frame_seen = $RjxzsCanStatus.MosFrameSeen
    historical_log_frame_seen = $RjxzsCanStatus.HistoricalLogFrameSeen
    default_channel_frame_seen = $RjxzsCanStatus.DefaultChannelFrameSeen
    raw_frame = $RjxzsCanStatus.RawFrame
    mos_raw_frame = $RjxzsCanStatus.MosRawFrame
    historical_log_raw_frame = $RjxzsCanStatus.HistoricalLogRawFrame
    default_channel_raw_frame = $RjxzsCanStatus.DefaultChannelRawFrame
    host_temperature_raw = $RjxzsCanStatus.HostTemperatureRaw
    status_accounting_hex = $RjxzsCanStatus.StatusAccountingHex
    active_protections = $RjxzsCanStatus.ActiveProtections
    mos_status_hex = $RjxzsCanStatus.MosStatusHex
    charge_allowed = $RjxzsCanStatus.ChargeAllowed
    discharge_allowed = $RjxzsCanStatus.DischargeAllowed
    charge_mos = $RjxzsCanStatus.ChargeMos
    discharge_mos = $RjxzsCanStatus.DischargeMos
    historical_log_code_hex = $RjxzsCanStatus.HistoricalLogCodeHex
    historical_log = $RjxzsCanStatus.HistoricalLog
    default_channel_code_hex = $RjxzsCanStatus.DefaultChannelCodeHex
    default_channel = $RjxzsCanStatus.DefaultChannel
    host_temperature_below_zero = $RjxzsCanStatus.HostTemperatureBelowZero
    channel_status_bit = $RjxzsCanStatus.ChannelStatusBit
    current_polarity = $RjxzsCanStatus.CurrentPolarity
    balancing = $RjxzsCanStatus.Balancing
    overdischarge = $RjxzsCanStatus.Overdischarge
    overcurrent = $RjxzsCanStatus.Overcurrent
    battery_string_error = $RjxzsCanStatus.BatteryStringError
    overcharge = $RjxzsCanStatus.Overcharge
    overtemperature = $RjxzsCanStatus.Overtemperature
    equalization_starting_voltage_mv = $RjxzsCanStatus.EqualizationStartingVoltageMv
  } | Export-Csv -Path $RjxzsCanCsv -NoTypeInformation -Append -Encoding UTF8
}

function Get-CellmonitorSample {
  param([string]$Html)

  $dataItems = Get-JsArrayItems -Html $Html -ArrayName "data"
  $balancingItems = Get-JsArrayItems -Html $Html -ArrayName "balancing"

  if ($dataItems.Count -eq 0) {
    throw "Cell voltage array is empty"
  }

  $voltages = @($dataItems | ForEach-Object { [int]$_ })
  $balancing = @($balancingItems | ForEach-Object { $_.ToLowerInvariant() -eq "true" })

  $minMv = ($voltages | Measure-Object -Minimum).Minimum
  $maxMv = ($voltages | Measure-Object -Maximum).Maximum
  $minIndex = [array]::IndexOf($voltages, [int]$minMv)
  $maxIndex = [array]::IndexOf($voltages, [int]$maxMv)
  $balancingCells = @()

  for ($i = 0; $i -lt $balancing.Count; $i++) {
    if ($balancing[$i]) {
      $balancingCells += ($i + 1)
    }
  }

  return [pscustomobject]@{
    Timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    Voltages = $voltages
    CellCount = $voltages.Count
    MinMv = [int]$minMv
    MaxMv = [int]$maxMv
    DeltaMv = [int]($maxMv - $minMv)
    MinCell = $minIndex + 1
    MaxCell = $maxIndex + 1
    BalancingCells = ($balancingCells -join ";")
  }
}

function Write-Sample {
  param(
    [pscustomobject]$CellSample,
    [pscustomobject]$MainSample
  )

  $cellColumns = if ($MaxCells -gt 0) { $MaxCells } else { $CellSample.CellCount }

  $row = [ordered]@{
    timestamp = $CellSample.Timestamp
    current_a = if ($null -ne $MainSample) { $MainSample.CurrentA } else { "" }
    voltage_v = if ($null -ne $MainSample) { $MainSample.VoltageV } else { "" }
    power_w = if ($null -ne $MainSample) { $MainSample.PowerW } else { "" }
    soc_pct = if ($null -ne $MainSample) { $MainSample.SocPct } else { "" }
    soh_pct = if ($null -ne $MainSample) { $MainSample.SohPct } else { "" }
    remaining_capacity_kwh = if ($null -ne $MainSample) { $MainSample.RemainingCapacityKwh } else { "" }
    total_capacity_kwh = if ($null -ne $MainSample) { $MainSample.TotalCapacityKwh } else { "" }
    cell_count = $CellSample.CellCount
    min_mv = $CellSample.MinMv
    max_mv = $CellSample.MaxMv
    delta_mv = $CellSample.DeltaMv
    min_cell = $CellSample.MinCell
    max_cell = $CellSample.MaxCell
    balancing_cells = $CellSample.BalancingCells
    max_charge_current_a = if ($null -ne $MainSample) { $MainSample.MaxChargeCurrentA } else { "" }
    max_discharge_current_a = if ($null -ne $MainSample) { $MainSample.MaxDischargeCurrentA } else { "" }
    max_charge_power_w = if ($null -ne $MainSample) { $MainSample.MaxChargePowerW } else { "" }
    max_discharge_power_w = if ($null -ne $MainSample) { $MainSample.MaxDischargePowerW } else { "" }
    temperature_min_c = if ($null -ne $MainSample) { $MainSample.TemperatureMinC } else { "" }
    temperature_max_c = if ($null -ne $MainSample) { $MainSample.TemperatureMaxC } else { "" }
    board_temperature_c = if ($null -ne $MainSample) { $MainSample.BoardTemperatureC } else { "" }
    battery_state = if ($null -ne $MainSample) { $MainSample.BatteryState } else { "" }
    system_status = if ($null -ne $MainSample) { $MainSample.SystemStatus } else { "" }
    rjxzs_charge_mos = if ($null -ne $MainSample) { $MainSample.RjxzsChargeMos } else { "" }
    rjxzs_discharge_mos = if ($null -ne $MainSample) { $MainSample.RjxzsDischargeMos } else { "" }
    rjxzs_log = if ($null -ne $MainSample) { $MainSample.RjxzsLog } else { "" }
    software = if ($null -ne $MainSample) { $MainSample.Software } else { "" }
  }

  for ($cell = 1; $cell -le $cellColumns; $cell++) {
    $index = $cell - 1
    $row["cell_{0:D3}_mv" -f $cell] = if ($index -lt $CellSample.Voltages.Count) { $CellSample.Voltages[$index] } else { "" }
  }

  [pscustomobject]$row | Export-Csv -Path $CompactCsv -NoTypeInformation -Append -Encoding UTF8
}

Write-Host "Logging $Url every $IntervalSeconds seconds"
Write-Host "Main page: $(Get-MainPageUrl)"
Write-Host "CSV: $CompactCsv"
Write-Host "Events CSV: $EventsCsv"
if ($UseCanLog) {
  Write-Host "CAN log page: $(Get-CanLogPageUrl)"
  Write-Host "RJXZS CAN status CSV: $RjxzsCanCsv"
}
Write-Host "One row per sample. Press Ctrl+C to stop."

while ($true) {
  try {
    $cellHtml = Get-PageContent -Url $Url -InputFile $InputFile
    $cellSample = Get-CellmonitorSample -Html $cellHtml

    $mainSample = $null
    try {
      $mainHtml = Get-PageContent -Url (Get-MainPageUrl) -InputFile $MainInputFile
      $mainSample = Get-MainSample -Html $mainHtml
    } catch {
      $line = "{0} MAIN_PAGE_ERROR {1}" -f (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz"), $_.Exception.Message
      Add-Content -Path $ErrorLog -Value $line -Encoding UTF8
      Write-Warning $line
    }

    $rjxzsCanStatus = $null
    if ($UseCanLog) {
      try {
        $canLogHtml = Get-PageContent -Url (Get-CanLogPageUrl) -InputFile $CanLogInputFile
        $rjxzsCanStatus = Get-RjxzsCanStatusSample -Html $canLogHtml -Timestamp $cellSample.Timestamp
        Write-RjxzsCanStatus -RjxzsCanStatus $rjxzsCanStatus
      } catch {
        $line = "{0} CAN_LOG_ERROR {1}" -f (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz"), $_.Exception.Message
        Add-Content -Path $ErrorLog -Value $line -Encoding UTF8
        Write-Warning $line
      }
    }

    Write-Sample -CellSample $cellSample -MainSample $mainSample
    Write-EventIfNeeded -CellSample $cellSample -MainSample $mainSample -RjxzsCanStatus $rjxzsCanStatus

    $socText = if ($null -ne $mainSample -and $null -ne $mainSample.SocPct) {
      " soc={0}%" -f $mainSample.SocPct
    } else {
      " soc=?"
    }

    $currentText = if ($null -ne $mainSample -and $null -ne $mainSample.CurrentA) {
      " current={0}A" -f $mainSample.CurrentA
    } else {
      " current=?"
    }

    $consoleChargeMos = if ($null -ne $mainSample) { Get-CleanText $mainSample.RjxzsChargeMos } else { "" }
    $consoleDischargeMos = if ($null -ne $mainSample) { Get-CleanText $mainSample.RjxzsDischargeMos } else { "" }
    if ($consoleChargeMos -eq "" -and $null -ne $rjxzsCanStatus) {
      $consoleChargeMos = Get-CleanText $rjxzsCanStatus.ChargeMos
    }
    if ($consoleDischargeMos -eq "" -and $null -ne $rjxzsCanStatus) {
      $consoleDischargeMos = Get-CleanText $rjxzsCanStatus.DischargeMos
    }

    $mosText = if ($consoleChargeMos -ne "" -or $consoleDischargeMos -ne "") {
      " mos=C:{0}/D:{1}" -f $consoleChargeMos, $consoleDischargeMos
    } else {
      ""
    }

    $activeText = if ($null -ne $rjxzsCanStatus) {
      " active={0}" -f $rjxzsCanStatus.ActiveProtections
    } else {
      ""
    }

    Write-Host ("{0}{1}{2}{3}{4} cells={5} min={6}mV(c{7}) max={8}mV(c{9}) delta={10}mV" -f `
      $cellSample.Timestamp,
      $socText,
      $currentText,
      $mosText,
      $activeText,
      $cellSample.CellCount,
      $cellSample.MinMv,
      $cellSample.MinCell,
      $cellSample.MaxMv,
      $cellSample.MaxCell,
      $cellSample.DeltaMv)
  } catch {
    $line = "{0} ERROR {1}" -f (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz"), $_.Exception.Message
    Add-Content -Path $ErrorLog -Value $line -Encoding UTF8
    Write-Warning $line
  }

  if ($Once) {
    break
  }

  Start-Sleep -Seconds $IntervalSeconds
}
