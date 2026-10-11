<#
  Phase 0 automated checks (3, 6, 8 from docs/PHASE0.md) in a separate, hidden Excel instance.
  Your open Excel windows are not touched.
  Usage:  pwsh tools/RunPhase0Checks.ps1 [-Xll <path>]
#>
param(
    [switch]$NoAddIn,   # control run: same Excel lifecycle without loading the add-in
    [switch]$LoadOnly,  # load the add-in, run no checks (isolates exit problems)
    [string]$Xll = (Join-Path $PSScriptRoot "..\src\DecisionTree.Excel\bin\Debug\net48\DecisionTreeStudio64.xll")
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $env:LOCALAPPDATA ("DecisionTreeStudio\logs\dts-" + (Get-Date -Format yyyyMMdd) + ".log")
$work = Join-Path $env:TEMP "dts-phase0"
New-Item -ItemType Directory -Force $work | Out-Null
$book = Join-Path $work "phase0.xlsx"
if (Test-Path $book) { Remove-Item $book -Force }

function LogLength { if (Test-Path $logFile) { (Get-Content $logFile).Count } else { 0 } }

# Waits for a log line matching $pattern written after line $from. Returns the line or $null.
function WaitForLog([int]$from, [string]$pattern, [int]$seconds = 120) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $logFile) {
            $hit = Get-Content $logFile | Select-Object -Skip $from | Where-Object { $_ -match $pattern } | Select-Object -First 1
            if ($hit) { return $hit }
        }
        Start-Sleep -Milliseconds 250
    }
    return $null
}

# Out-of-process COM: every object this script holds keeps Excel alive until released.
function Release($o) { if ($null -ne $o) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } }

function Result([string]$name, [bool]$pass, [string]$detail) {
    $mark = if ($pass) { "PASS" } else { "FAIL" }
    Write-Output ("[{0}] {1}: {2}" -f $mark, $name, $detail)
}

# All COM objects live inside this function, so they are out of scope (collectable) afterwards.
function Invoke-Checks {
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $script:xlPid = (Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1).Id

    try {
        $wb = $xl.Workbooks.Add()
        if ($NoAddIn) { Result "Control run" $true "add-in not loaded"; return }
        if (-not $xl.RegisterXLL((Resolve-Path $Xll).Path)) { throw "RegisterXLL failed for $Xll" }
        if ($LoadOnly) { Start-Sleep -Seconds 2; Result "Load-only run" $true "add-in loaded, no checks"; return }

        # ---- Check 6: render speed + only owned shapes deleted ----
        foreach ($n in 500, 1000) {
            $start = LogLength
            $xl.Run("DTS_SpikeRender$n") | Out-Null
            $line = WaitForLog $start "Rendered $n nodes|failed"
            Result "Render $n nodes" ($line -match "Rendered") ($line -replace '^.*?(INFO|ERROR) ', '')
        }

        $sheet = $xl.ActiveSheet
        $mine = $sheet.Shapes.AddShape(1, 600, 20, 40, 40)   # a user's own shape
        $mine.Name = "UserShape"
        $start = LogLength
        $xl.Run("DTS_SpikeClear") | Out-Null
        $line = WaitForLog $start "Deleted \d+ owned shapes|failed"
        $kept = $sheet.Shapes.Count -eq 1 -and $sheet.Shapes.Item(1).Name -eq "UserShape"
        Result "Clear keeps other shapes" $kept (($line -replace '^.*?(INFO|ERROR) ', '') + "; remaining shapes: " + $sheet.Shapes.Count)
        Release $mine; $mine = $null
        Release $sheet; $sheet = $null

        # ---- Check 3: custom XML part survives save / close / reopen as .xlsx ----
        $start = LogLength
        $xl.Run("DTS_SpikeWriteXml") | Out-Null
        $written = WaitForLog $start "XML part written|failed"
        $stamp = if ($written -match "at (\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)") { $Matches[1] } else { "?" }
        $wb.SaveAs($book, 51)   # 51 = xlOpenXMLWorkbook (.xlsx, no macros)
        $wb.Close($false)
        Release $wb
        $wb = $xl.Workbooks.Open($book)
        $start = LogLength
        $xl.Run("DTS_SpikeReadXml") | Out-Null
        $read = WaitForLog $start "Found in|No spike XML part|failed"
        Result "XML part survives save/reopen" ($read -match [regex]::Escape($stamp)) ("written $stamp; read: " + ($read -replace '^.*?(INFO|WARN|ERROR) ', ''))
    }
    catch {
        Result "Run" $false $_.Exception.Message
    }
    finally {
        Release $mine; Release $sheet
        if ($null -ne $wb) { try { $wb.Close($false) } catch { }; Release $wb }
        $xl.Quit()
        Release $xl
    }
}

Invoke-Checks

# ---- Check 8: Excel exits cleanly ----
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
$gone = $false
for ($i = 0; $i -lt 80 -and -not $gone; $i++) {
    Start-Sleep -Milliseconds 250
    $gone = -not (Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue)
}
Result "Excel exits cleanly" $gone ("process $script:xlPid " + $(if ($gone) { "ended" } else { "still running" }))
