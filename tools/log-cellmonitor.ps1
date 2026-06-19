param(
  [string]$Url = "http://192.168.68.104/cellmonitor",
  [string]$MainUrl = "",
  [int]$IntervalSeconds = 60,
  [string]$OutDir = ".\cellmonitor-logs",
  [int]$MaxCells = 0,
  [string]$InputFile = "",
  [string]$MainInputFile = "",
  [switch]$Once
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$CompactCsv = Join-Path $OutDir "battery_log.csv"
$ErrorLog = Join-Path $OutDir "cellmonitor_errors.log"

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
  $maxChargeValue = Get-NumberMatch -Text $text -Pattern "Max charge power:\s*([-+]?\d+(?:\.\d+)?)\s*(kW|W)"
  $maxChargeUnit = Get-TextMatch -Text $text -Pattern "Max charge power:\s*[-+]?\d+(?:\.\d+)?\s*(kW|W)"
  $maxDischargeValue = Get-NumberMatch -Text $text -Pattern "Max discharge power:\s*([-+]?\d+(?:\.\d+)?)\s*(kW|W)"
  $maxDischargeUnit = Get-TextMatch -Text $text -Pattern "Max discharge power:\s*[-+]?\d+(?:\.\d+)?\s*(kW|W)"

  return [pscustomobject]@{
    Software = Get-TextMatch -Text $text -Pattern "Software:\s*([^\s]+)"
    BoardTemperatureC = Get-NumberMatch -Text $text -Pattern "Hardware:.*?@\s*([-+]?\d+(?:\.\d+)?)\s*[^0-9\r\n]*C"
    SocPct = Get-NumberMatch -Text $text -Pattern "SOC:\s*([-+]?\d+(?:\.\d+)?)%"
    SohPct = Get-NumberMatch -Text $text -Pattern "SOH:\s*([-+]?\d+(?:\.\d+)?)%"
    VoltageV = Get-NumberMatch -Text $text -Pattern "Voltage:\s*([-+]?\d+(?:\.\d+)?)\s*V"
    CurrentA = Get-NumberMatch -Text $text -Pattern "Current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    PowerW = Get-NumberMatch -Text $text -Pattern "Power:\s*([-+]?\d+(?:\.\d+)?)\s*W"
    MaxDischargePowerW = Convert-PowerToW -Value $maxDischargeValue -Unit $maxDischargeUnit
    MaxChargePowerW = Convert-PowerToW -Value $maxChargeValue -Unit $maxChargeUnit
    MaxDischargeCurrentA = Get-NumberMatch -Text $text -Pattern "Max discharge current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    MaxChargeCurrentA = Get-NumberMatch -Text $text -Pattern "Max charge current:\s*([-+]?\d+(?:\.\d+)?)\s*A"
    TemperatureMinC = Get-NumberMatch -Text $text -Pattern "Temperature min/max:\s*([-+]?\d+(?:\.\d+)?)\s*[^/\r\n]*C"
    TemperatureMaxC = Get-NumberMatch -Text $text -Pattern "Temperature min/max:\s*[-+]?\d+(?:\.\d+)?\s*[^/\r\n]*C\s*/\s*([-+]?\d+(?:\.\d+)?)\s*[^0-9\r\n]*C"
    BatteryState = Get-TextMatch -Text $text -Pattern "(Battery (?:idle|charging[^\r\n]*|discharging[^\r\n]*))"
    SystemStatus = Get-TextMatch -Text $text -Pattern "System status:\s*([^\r\n]+)"
    RjxzsChargeMos = Get-TextMatch -Text $text -Pattern "RJXZS BMS:\s*Charge MOS:\s*([^|\r\n]+)"
    RjxzsDischargeMos = Get-TextMatch -Text $text -Pattern "Discharge MOS:\s*([^|\r\n]+)"
    RjxzsLog = Get-TextMatch -Text $text -Pattern "Log:\s*([^\r\n]+)"
  }
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

    Write-Sample -CellSample $cellSample -MainSample $mainSample

    $currentText = if ($null -ne $mainSample -and $null -ne $mainSample.CurrentA) {
      " current={0}A" -f $mainSample.CurrentA
    } else {
      ""
    }

    Write-Host ("{0} cells={1} min={2}mV(c{3}) max={4}mV(c{5}) delta={6}mV{7}" -f `
      $cellSample.Timestamp,
      $cellSample.CellCount,
      $cellSample.MinMv,
      $cellSample.MinCell,
      $cellSample.MaxMv,
      $cellSample.MaxCell,
      $cellSample.DeltaMv,
      $currentText)
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
