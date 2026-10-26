# Live Excel view of the Caller ID database.
# Opens Excel and refreshes the data every few seconds while Excel stays open. Nothing to download.
# The database workbook itself is NOT opened, so the two websites can keep writing into it.
param([int]$Seconds = 3)
$ErrorActionPreference = 'Stop'

$port = 4001
try {
  $cfg = Get-Content (Join-Path $PSScriptRoot 'config.json') -Raw | ConvertFrom-Json
  if ($cfg.operatorPort) { $port = [int]$cfg.operatorPort }
} catch { }
$base = "http://localhost:$port"

$feeds = @(
  @{ Sheet = 'Requests'; Url = "$base/api/requests.csv"; LinkCol = 'Q'; UrlCol = 'P' },
  @{ Sheet = 'Numbers'; Url = "$base/api/numbers.csv"; LinkCol = $null; UrlCol = $null },
  @{ Sheet = 'Packages'; Url = "$base/api/packages.csv"; LinkCol = $null; UrlCol = $null }
)
$tmp = Join-Path $env:TEMP 'callerid-live'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

function Get-Feed($feed) {
  $file = Join-Path $tmp ($feed.Sheet + '.csv')
  Invoke-WebRequest -Uri $feed.Url -OutFile $file -UseBasicParsing
  return $file
}

try { [void](Invoke-WebRequest -Uri "$base/api/info" -UseBasicParsing -TimeoutSec 5) }
catch {
  Write-Host "Cannot reach the Operator website at $base"
  Write-Host "Start the websites first:  npm start   (in the project folder)"
  exit 1
}

try { $xl = New-Object -ComObject Excel.Application }
catch { Write-Host 'Microsoft Excel was not found on this computer.'; exit 1 }
$xl.Visible = $true
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -lt $feeds.Count) {
  [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
}

function Update-Links($ws, $feed) {
  if (-not $feed.LinkCol) { return }
  $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row   # last used row (xlUp)
  $ws.Range("$($feed.LinkCol)1").Value2 = 'link'
  if ($last -ge 2) {
    $ws.Range("$($feed.LinkCol)2:$($feed.LinkCol)$last").Formula = "=IF($($feed.UrlCol)2="""","""",HYPERLINK($($feed.UrlCol)2,""Open file""))"
  }
}

$live = @()
for ($i = 0; $i -lt $feeds.Count; $i++) {
  $ws = $wb.Worksheets.Item($i + 1)
  $ws.Name = $feeds[$i].Sheet
  $file = Get-Feed $feeds[$i]
  $qt = $ws.QueryTables.Add("TEXT;$file", $ws.Range('A1'))
  $qt.TextFileParseType = 1            # delimited
  $qt.TextFileCommaDelimiter = $true
  $qt.TextFilePlatform = 65001         # UTF-8
  $qt.FieldNames = $true
  $qt.RefreshStyle = 0                 # overwrite cells
  $qt.AdjustColumnWidth = $true
  $qt.BackgroundQuery = $false
  [void]$qt.Refresh($false)
  Update-Links $ws $feeds[$i]
  $ws.Columns.AutoFit() | Out-Null
  $live += [pscustomobject]@{ Ws = $ws; Qt = $qt; Feed = $feeds[$i] }
}
$wb.Worksheets.Item(1).Activate()
$wb.Saved = $true

Write-Host ''
Write-Host "Live view is running (refresh every $Seconds s). Keep this window open."
Write-Host 'Close Excel, or press Ctrl+C here, to stop.'
while ($true) {
  Start-Sleep -Seconds $Seconds
  try { [void]$xl.Version } catch { break }          # Excel was closed
  try {
    if ($xl.Ready) {                                 # skip while you are typing in a cell
      foreach ($l in $live) {
        [void](Get-Feed $l.Feed)
        [void]$l.Qt.Refresh($false)
        Update-Links $l.Ws $l.Feed
      }
      $wb.Saved = $true
    }
  } catch { }                                        # Excel busy: try again on the next round
}
Write-Host 'Excel closed. Live view stopped.'
