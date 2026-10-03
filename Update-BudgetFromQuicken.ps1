<#
.SYNOPSIS
    Updates column C (Spent/Received) of the current month's sheet in the budget
    workbook using amounts parsed from a Quicken report.

.DESCRIPTION
    Reads a Quicken report (clipboard or -QuickenPath file) via Get-QuickenReportData
    (Parse-QuickenReport.ps1), then for each Group/Category subtotal name in that data
    looks for a matching name in column A of the target month's sheet. When exactly one
    matching row has a numeric value in column C, that cell is overwritten with the Quicken
    amount unless its formula uses a SUM function (including SUMIF and SUMIFS). SUM formulas
    are left alone, while formulas that add individual items can be updated. Ambiguous
    (multiple candidate rows) or unmatched names are reported so they can be handled manually.

    Connects to Excel via COM. If the workbook is already open, it is updated in place;
    otherwise it is opened (and, if this script started Excel itself, closed again when done).

.PARAMETER QuickenPath
    Optional path to a text file with the Quicken report instead of reading the clipboard.

.PARAMETER WorkbookPath
    Path to the budget workbook. Defaults to the standard xBudget 2026.xlsm location.

.PARAMETER SheetName
    Optional sheet name override. By default, the month is determined from the Quicken report date range.

.PARAMETER DryRun
    Report what would change without writing or saving anything.

.PARAMETER DebugMode
    Keep the PowerShell console open after the script finishes or encounters an error.
#>
[CmdletBinding()]
param(
    [string]$QuickenPath,
    [string]$WorkbookPath = 'C:\Users\Glenn\My Drive\Finances\Budget\Budget 2026.xlsm',
    [string]$SheetName,
    [switch]$DryRun,
    [switch]$DebugMode
)

# Keep a transcript because the shortcut may launch this script without a visible console.
$logDirectory = Join-Path $PSScriptRoot 'logs'
New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logDirectory ((Get-Date).ToString('yyyyMMddHHmmss') + '.log')
$transcriptStarted = $false
Start-Transcript -Path $logPath -Append | Out-Null
$transcriptStarted = $true

function Stop-RunTranscript {
    if ($script:transcriptStarted) {
        Stop-Transcript | Out-Null
        $script:transcriptStarted = $false
    }
}

function Wait-DebugConsole {
    if ($DebugMode) {
        Read-Host 'Debug mode: press Enter to close this console'
    }
}

trap {
    Stop-RunTranscript
    Wait-DebugConsole
    break
}

# Define the small set of Windows APIs needed to find, activate, and send keys to windows.
if (-not ('QuickenReport.NativeMethods' -as [type])) {
    Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace QuickenReport {
    public static class NativeMethods {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }
}
'@
}

function Find-VisibleWindow {
    param(
        [string]$TitlePattern,
        [string]$ExactTitle
    )

    # Enumerate top-level windows instead of relying on process MainWindowTitle, which can be blank.
    $matchingWindows = [System.Collections.Generic.List[object]]::new()
    $callback = [QuickenReport.NativeMethods+EnumWindowsProc] {
        param($hWnd, $lParam)
        if ([QuickenReport.NativeMethods]::IsWindowVisible($hWnd)) {
            $titleBuffer = [Text.StringBuilder]::new(512)
            [void][QuickenReport.NativeMethods]::GetWindowText($hWnd, $titleBuffer, $titleBuffer.Capacity)
            $title = $titleBuffer.ToString()
            if ($title -like $TitlePattern) {
                $matchingWindows.Add([PSCustomObject]@{ Handle = $hWnd; Title = $title })
            }
        }
        return $true
    }
    [void][QuickenReport.NativeMethods]::EnumWindows($callback, [IntPtr]::Zero)

    $window = $matchingWindows |
        Sort-Object @{ Expression = { if ($_.Title -ieq $ExactTitle) { 0 } else { 1 } } } |
        Select-Object -First 1
    return $window
}

function Copy-MonthlyExpensesReport {
    # The report window is separate from Quicken's main window, so locate it after the shortcut runs.
    $window = Find-VisibleWindow -TitlePattern '*Monthly Expenses*' -ExactTitle 'Monthly Expenses'
    if (-not $window) {
        throw "Could not find a visible window with 'Monthly Expenses' in its title."
    }

    $shell = New-Object -ComObject WScript.Shell
    $activated = [QuickenReport.NativeMethods]::SetForegroundWindow($window.Handle)
    if (-not $activated) {
        $activated = $shell.AppActivate($window.Title)
    }
    if (-not $activated) {
        throw "Could not activate the '$($window.Title)' window."
    }

    Start-Sleep -Milliseconds 150
    $shell.SendKeys('^c')
    Start-Sleep -Milliseconds 300
}

function Wait-ForMonthlyExpensesWindow {
    param([int]$TimeoutMs = 10000)

    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    while ((Get-Date) -lt $deadline) {
        $window = Find-VisibleWindow -TitlePattern '*Monthly Expenses*' -ExactTitle 'Monthly Expenses'
        if ($window) { return $window }
        Start-Sleep -Milliseconds 250
    }
    return $null
}

function Wait-ForForegroundWindow {
    param(
        [IntPtr]$Handle,
        [int]$TimeoutMs = 3000
    )

    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    while ((Get-Date) -lt $deadline) {
        if ([QuickenReport.NativeMethods]::GetForegroundWindow() -eq $Handle) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return $false
}

function Open-AndCopy-MonthlyExpensesReport {
    # Open the report from the real Quicken window, then copy its grid to the clipboard.
    $quickenWindow = Find-VisibleWindow -TitlePattern '*Quicken Classic Deluxe*' -ExactTitle 'Quicken Classic Deluxe'
    if (-not $quickenWindow) {
        throw "Could not find a visible Quicken window."
    }
    if (-not [QuickenReport.NativeMethods]::SetForegroundWindow($quickenWindow.Handle)) {
        throw "Could not activate the '$($quickenWindow.Title)' window."
    }

    Write-Host "Activating Quicken window '$($quickenWindow.Title)' and opening Monthly Expenses..." -ForegroundColor Cyan
    $shell = New-Object -ComObject WScript.Shell
    if (-not $shell.AppActivate($quickenWindow.Title)) {
        throw "Could not bring the '$($quickenWindow.Title)' window to the foreground."
    }
    if (-not (Wait-ForForegroundWindow -Handle $quickenWindow.Handle)) {
        throw "The '$($quickenWindow.Title)' window never became the foreground window."
    }

    $maxAttempts = 3
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $shell.SendKeys('%+E')
        if (Wait-ForMonthlyExpensesWindow) { break }

        if ($attempt -lt $maxAttempts) {
            Write-Host "The report window did not appear; resending the shortcut (attempt $($attempt + 1) of $maxAttempts)." -ForegroundColor Yellow
            [void][QuickenReport.NativeMethods]::SetForegroundWindow($quickenWindow.Handle)
            [void](Wait-ForForegroundWindow -Handle $quickenWindow.Handle -TimeoutMs 1500)
        } else {
            throw "The Monthly Expenses report window did not appear after $maxAttempts attempts."
        }
    }
    Copy-MonthlyExpensesReport
}

if (-not $QuickenPath) {
    # File-based runs are deterministic and do not need to drive the Quicken UI.
    Open-AndCopy-MonthlyExpensesReport
}

# Load the shared parser after clipboard acquisition so the UI automation remains isolated here.
. (Join-Path $PSScriptRoot 'Parse-QuickenReport.ps1')

# Quicken category name -> spreadsheet category name.
$categoryMap = @{
    'charitable gifts'       = 'Charitable Gifts'
    'classroom'              = 'Classroom'
    'clothing'               = 'Clothing'
    'food-groceries'         = 'Grocery'
    'food-restaurants'       = 'Restaurants'
    'housing-maintenance'    = 'Housing'
    'insurance'              = 'Insurance'
    'medical'                = 'Medical'
    'personal-gifts'         = 'Gifts'
    'personal-hair care'     = 'Hair Care'
    'personal-miscellaneous' = 'Miscellaneous'
    'personal-pets'          = 'Pets'
    'recreation-vacation'    = 'Vacation'
    'transportation-gas'     = 'Gas'
    'transportation-insurance' = 'Car Insurance'
    'transportation-repairs' = 'Car Repairs'
    'utilities'              = 'Utilities'
    'utilities-entertainment' = 'Entertainment'
    'hsabank'                = 'HSA/Reimburse'
    'savings'                = 'Savings'
    "kohl's"                = 'Debt Misc'
    "american express"         = 'Debt Misc'
    "discover"                = 'Debt Misc'
    "united airlines"          = 'Debt Misc'
    'carecredit'             = 'Debt CareCredit'
    "jessica's student loan" = 'Student Loan'
    'rocket mortgage escrow' = 'Home Mortgage'
    'rocket mortgage interest' = 'Home Mortgage'
    'rocket mortgage'        = 'Home Mortgage'
    'schwab joint tenant'    = 'Retirement'
    'schwab roth contributory ira' = 'Retirement'
}
$utilitiesChildMap = @{
    'cable'       = 'Internet'
    'cellular'    = 'Cellular'
    'electricity' = 'Electric'
    'hoa'         = 'HOA'
    'storm water' = 'Storm Water'
    'trash'       = 'Trash'
    'water'       = 'Water/Sewer'
}
$entertainmentChildMap = @{
    'netflix'    = 'Netflix'
    'google pay' = 'Paramount+'
    'paypal'     = 'Hulu'
}
$vendorMap = @{
    'state farm'           = 'State Farm'
    'state farm insurance' = 'State Farm'
}
$protectedSpreadsheetNames = @('Savings')

# Acquire structured report data from the copied report or the optional input file.
$data = Get-QuickenReportData -Path $QuickenPath

$dateRangePattern = '^\s*(\d{1,2}/\d{1,2}/\d{4})\s+through\s+(\d{1,2}/\d{1,2}/\d{4})\s*$'
if ($data.DateRange -notmatch $dateRangePattern) {
    Write-Warning "Could not determine a single report month from the Quicken date range '$($data.DateRange)'. No spreadsheet updates were made."
    Stop-RunTranscript
    Wait-DebugConsole
    return
}

try {
    $startDate = [DateTime]::Parse($Matches[1], [Globalization.CultureInfo]::InvariantCulture)
    $endDate = [DateTime]::Parse($Matches[2], [Globalization.CultureInfo]::InvariantCulture)
} catch {
    Write-Warning "Could not parse the Quicken date range '$($data.DateRange)'. No spreadsheet updates were made."
    Stop-RunTranscript
    Wait-DebugConsole
    return
}

if ($startDate.Year -ne $endDate.Year -or $startDate.Month -ne $endDate.Month) {
    Write-Warning "The Quicken report spans more than one month ('$($data.DateRange)'). No spreadsheet updates were made."
    Stop-RunTranscript
    Wait-DebugConsole
    return
}

# The report month determines the target worksheet unless an explicit override is supplied.
$SheetName = $startDate.ToString('MMMM')

# Map Quicken groups/categories/vendors to the names used by the budget workbook.
$currentSection = $null
$mappedEntries = foreach ($entry in $data.Entries | Where-Object {
        $_.Type -in 'Group', 'Category' -or
        ($_.Type -eq 'Transaction' -and $_.Description)
    }) {
    $map = $categoryMap
    if ($entry.Type -eq 'Transaction') {
        $sourceName = $entry.Description.Trim()
        $vendorKey = $sourceName.ToLowerInvariant()
        $canonicalVendor = if ($vendorMap.ContainsKey($vendorKey)) { $vendorMap[$vendorKey] } else { $sourceName }
        $quickenKey = $canonicalVendor.ToLowerInvariant()
        if ($entry.Section -ieq 'INCOME') {
            $map = @{ $quickenKey = $canonicalVendor }
            $scope = 'Income'
        } elseif ($currentSection -ieq 'Medical') {
            $map = @{ $quickenKey = $canonicalVendor }
            $scope = 'Medical'
        } elseif ($currentSection -ieq 'Utilities-Entertainment') {
            $quickenKey = $sourceName.ToLowerInvariant()
            $map = $entertainmentChildMap
            $scope = 'Entertainment'
        } else {
            continue
        }
    } else {
        $quickenKey = $entry.Name.Trim().ToLowerInvariant()
        if ($categoryMap.ContainsKey($quickenKey)) {
            $map = $categoryMap
            $currentSection = $entry.Name
        } else {
            $map = $utilitiesChildMap
        }
    }
    if ($map.ContainsKey($quickenKey)) {
        $spreadsheetName = $map[$quickenKey]
        if ($entry.Type -ne 'Transaction') { $sourceName = $entry.Name }
    }
    if ($spreadsheetName) {
        [PSCustomObject]@{
            Name   = $spreadsheetName
            Source = $sourceName
            Amount = $entry.Amount
            Scope  = $scope
        }
    }
    $spreadsheetName = $null
    $scope = $null
}
# Combine duplicate mapped entries into one spreadsheet candidate, preserving a formula when useful.
$candidates = foreach ($group in ($mappedEntries | Group-Object -Property Name)) {
    $amount = ($group.Group | Measure-Object -Property Amount -Sum).Sum
    $formula = if ($group.Count -gt 1) {
        $terms = @($group.Group | ForEach-Object {
                '(' + $_.Amount.ToString('0.################', [Globalization.CultureInfo]::InvariantCulture) + ')'
            })
        '=' + ($terms -join '+')
    }
    [PSCustomObject]@{
        Name    = $group.Name
        Source  = ($group.Group.Source -join ' + ')
        Amount  = $amount
        Formula = $formula
        Scope   = ($group.Group.Scope | Where-Object { $_ } | Select-Object -First 1)
    }
}

if (-not $candidates) {
    Write-Warning "No category/group entries found in the Quicken report."
    Stop-RunTranscript
    Wait-DebugConsole
    return
}

# Connect to an existing workbook when possible; otherwise open the configured workbook through COM.
$excel = $null
$workbook = $null
$startedExcel = $false
$openedWorkbook = $false

try {
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $excel = $null
    }

    if ($excel) {
        try {
            $workbooks = $excel.Workbooks
            $workbook = $workbooks | Where-Object { $_.FullName -eq $WorkbookPath }
            if ($null -eq $workbooks) { $excel = $null }
        } catch {
            $excel = $null
            $workbooks = $null
        }
    }

    if (-not $excel) {
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $workbooks = $excel.Workbooks
        } catch {
            throw "Could not start Microsoft Excel: $($_.Exception.Message)"
        }
        $excel.Visible = $true
        $startedExcel = $true
    }

    if (-not $workbook) {
        if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
        if ($null -eq $workbooks) { throw "Microsoft Excel is unavailable or its workbooks collection could not be opened." }
        $workbook = $workbooks.Open($WorkbookPath)
        $openedWorkbook = $true
    }

    $sheet = $workbook.Sheets | Where-Object { $_.Name -eq $SheetName }
    if (-not $sheet) { throw "Sheet '$SheetName' not found in $WorkbookPath" }

    # Read the relevant columns in bulk to avoid a COM round-trip for every spreadsheet cell.
    $usedRange = $sheet.UsedRange
    $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1

    # Bulk-read columns A, C, and H in single COM round-trips instead of per-cell calls (much faster).
    $namesArr = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($lastRow, 1)).Value2
    $formulaArr = $sheet.Range($sheet.Cells.Item(1, 3), $sheet.Cells.Item($lastRow, 3)).Formula
    $valueArr = $sheet.Range($sheet.Cells.Item(1, 3), $sheet.Cells.Item($lastRow, 3)).Value2
    $flagsArr = $sheet.Range($sheet.Cells.Item(1, 8), $sheet.Cells.Item($lastRow, 8)).Value2

    # Index eligible worksheet rows, including reserved blank slots for special child categories.
    $rowsByName = @{}
    $medicalBlankRows = [System.Collections.Generic.List[int]]::new()
    $incomeMiscBlankRows = [System.Collections.Generic.List[int]]::new()
    $medicalParentRow = $null
    $incomeMiscChildBlock = $false
    for ($row = 1; $row -le $lastRow; $row++) {
        $name = $namesArr[$row, 1]
        $flag = $flagsArr[$row, 1]
        $isIndividualEntry = $flag -is [string] -and $flag.Trim() -ieq 'X'
        $nameKey = if ($name -is [string]) { $name.Trim().ToLowerInvariant() } else { $null }
        if ($nameKey -eq 'income misc') { $incomeMiscChildBlock = $true }
        $isIncomeChild = $incomeMiscChildBlock -and $nameKey -ne 'income'
        if ($incomeMiscChildBlock -and $nameKey -eq 'income') { $incomeMiscChildBlock = $false }
        if ($nameKey -eq 'medical') { $medicalParentRow = $row }
        if ($medicalParentRow -and $row -gt $medicalParentRow -and $nameKey -and $categoryMap.ContainsKey($nameKey)) { $medicalParentRow = $null }
        $mappedChildNames = @($utilitiesChildMap.Values) + @($entertainmentChildMap.Values)
        $isMedicalChild = $medicalParentRow -and $row -gt $medicalParentRow -and $isIndividualEntry
        $isMappedChild = $nameKey -and ($mappedChildNames -contains $name.Trim() -or $isMedicalChild -or $isIncomeChild)
        if ((-not $isIndividualEntry -or $isMappedChild) -and $name -is [string] -and $name.Trim()) {
            $key = $name.Trim().ToLowerInvariant()
            if (-not $rowsByName.ContainsKey($key)) { $rowsByName[$key] = [System.Collections.Generic.List[int]]::new() }
            $rowsByName[$key].Add($row)
        } elseif ($isMedicalChild -and -not $nameKey) {
            $medicalBlankRows.Add($row)
        } elseif ($isIncomeChild -and -not $nameKey) {
            $incomeMiscBlankRows.Add($row)
        }
    }

    # Resolve each Quicken candidate to exactly one writable budget row before changing Excel.
    $updated = 0
    foreach ($candidate in $candidates) {
        if ($protectedSpreadsheetNames -contains $candidate.Name) {
            Write-Host "Skipping protected spreadsheet category '$($candidate.Name)' from '$($candidate.Source)'." -ForegroundColor DarkGray
            continue
        }

        $matches = $rowsByName[$candidate.Name.Trim().ToLowerInvariant()]

        if (-not $matches -and $candidate.Scope -eq 'Medical' -and $medicalBlankRows.Count -gt 0) {
            $newRow = [int]$medicalBlankRows[0]
            $medicalBlankRows.RemoveAt(0)
            if (-not $DryRun) {
                $newName = [string]$candidate.Name
                $sheet.Range("A$newRow").Value = $newName
            }
            $rowsByName[$candidate.Name.Trim().ToLowerInvariant()] = [System.Collections.Generic.List[int]]::new()
            $rowsByName[$candidate.Name.Trim().ToLowerInvariant()].Add($newRow)
            $matches = $rowsByName[$candidate.Name.Trim().ToLowerInvariant()]
        }

        if (-not $matches -and $candidate.Scope -eq 'Income' -and $incomeMiscBlankRows.Count -gt 0) {
            $newRow = [int]$incomeMiscBlankRows[0]
            $incomeMiscBlankRows.RemoveAt(0)
            if (-not $DryRun) {
                $newName = [string]$candidate.Name
                $sheet.Range("A$newRow").Value = $newName
            }
            $rowsByName[$candidate.Name.Trim().ToLowerInvariant()] = [System.Collections.Generic.List[int]]::new()
            $rowsByName[$candidate.Name.Trim().ToLowerInvariant()].Add($newRow)
            $matches = $rowsByName[$candidate.Name.Trim().ToLowerInvariant()]
        }

        if (-not $matches) {
            Write-Warning "No spreadsheet match for Quicken category '$($candidate.Source)' (mapped to '$($candidate.Name)')."
            continue
        }

        # SUM-family formulas remain authoritative; additive formulas can be replaced by Quicken values.
        $writableRows = $matches | Where-Object {
            $f = $formulaArr[$_, 1]
            $value = $valueArr[$_, 1]
            -not ($f -is [string] -and $f -match '(?i)(^|[^A-Z0-9_])SUM[A-Z0-9_]*\s*\(') -and
                $value -is [ValueType] -and
                $value -isnot [bool]
        }

        if ($writableRows.Count -eq 0) {
            Write-Host "Skipping '$($candidate.Name)' from '$($candidate.Source)': matching row(s) $($matches -join ', ') are not writable numeric cells in column C." -ForegroundColor DarkGray
            continue
        }

        if ($writableRows.Count -gt 1) {
            Write-Warning "Ambiguous match for '$($candidate.Name)' from '$($candidate.Source)': rows $($writableRows -join ', ') all have writable column C. Skipping."
            continue
        }

        $targetRow = $writableRows[0]
        $oldValue = $valueArr[$targetRow, 1]
        $newValue = if ($candidate.Formula) { $candidate.Formula } else { $candidate.Amount }

        Write-Host ("{0,-24} <- {1,-28} row {2,-4} C: {3,10} -> {4,10}" -f $candidate.Name, $candidate.Source, $targetRow, $oldValue, $newValue) -ForegroundColor Green

        # DryRun reports the same decisions without changing or saving the workbook.
        if (-not $DryRun) {
            if ($candidate.Formula) {
                $sheet.Cells.Item($targetRow, 3).Formula = $candidate.Formula
            } else {
                $sheet.Cells.Item($targetRow, 3).Value2 = $candidate.Amount
            }
        }
        $updated++
    }

    if ($DryRun) {
        Write-Host "`nDry run: $updated cell(s) would be updated. No changes saved." -ForegroundColor Yellow
    } else {
        $workbook.Save()
        Write-Host "`nSaved $WorkbookPath ($updated cell(s) updated)." -ForegroundColor Cyan
    }
}
finally {
    # Release COM ownership and close Excel only when this script started it.
    if ($excel) {
        if ($openedWorkbook -and $workbook -and -not $DryRun) {
            # workbook was opened by this script and already saved above; leave it open for review
        }
        if ($startedExcel -and -not $openedWorkbook) {
            $excel.Quit()
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

Stop-RunTranscript
Wait-DebugConsole
