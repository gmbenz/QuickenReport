<#
.SYNOPSIS
    Updates column C (Spent/Received) of the current month's sheet in the budget
    workbook using amounts parsed from a Quicken report.

.DESCRIPTION
    Reads a Quicken report (clipboard or -QuickenPath file) via Get-QuickenReportData
    (Parse-QuickenReport.ps1), then for each Group/Category subtotal name in that data
    looks for a matching name in column A of the target month's sheet. When exactly one
    matching row has a plain (non-formula) numeric value in column C, that cell is
    overwritten with the Quicken amount. Rows whose column C is a formula are left alone,
    and ambiguous (multiple candidate rows) or unmatched names are reported so they can be
    handled manually.

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
#>
[CmdletBinding()]
param(
    [string]$QuickenPath,
    [string]$WorkbookPath = 'C:\Users\Glenn\My Drive\Finances\Budget\Budget 2026.xlsm',
    [string]$SheetName,
    [switch]$DryRun
)

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

$data = Get-QuickenReportData -Path $QuickenPath

$dateRangePattern = '^\s*(\d{1,2}/\d{1,2}/\d{4})\s+through\s+(\d{1,2}/\d{1,2}/\d{4})\s*$'
if ($data.DateRange -notmatch $dateRangePattern) {
    Write-Warning "Could not determine a single report month from the Quicken date range '$($data.DateRange)'. No spreadsheet updates were made."
    return
}

try {
    $startDate = [DateTime]::Parse($Matches[1], [Globalization.CultureInfo]::InvariantCulture)
    $endDate = [DateTime]::Parse($Matches[2], [Globalization.CultureInfo]::InvariantCulture)
} catch {
    Write-Warning "Could not parse the Quicken date range '$($data.DateRange)'. No spreadsheet updates were made."
    return
}

if ($startDate.Year -ne $endDate.Year -or $startDate.Month -ne $endDate.Month) {
    Write-Warning "The Quicken report spans more than one month ('$($data.DateRange)'). No spreadsheet updates were made."
    return
}

$SheetName = $startDate.ToString('MMMM')

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
$candidates = foreach ($group in ($mappedEntries | Group-Object -Property Name)) {
    [PSCustomObject]@{
        Name   = $group.Name
        Source = ($group.Group.Source -join ' + ')
        Amount = ($group.Group | Measure-Object -Property Amount -Sum).Sum
        Scope  = ($group.Group.Scope | Where-Object { $_ } | Select-Object -First 1)
    }
}

if (-not $candidates) {
    Write-Warning "No category/group entries found in the Quicken report."
    return
}

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

    $usedRange = $sheet.UsedRange
    $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1

    # Bulk-read columns A, C, and H in single COM round-trips instead of per-cell calls (much faster).
    $namesArr = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($lastRow, 1)).Value2
    $formulaArr = $sheet.Range($sheet.Cells.Item(1, 3), $sheet.Cells.Item($lastRow, 3)).Formula
    $valueArr = $sheet.Range($sheet.Cells.Item(1, 3), $sheet.Cells.Item($lastRow, 3)).Value2
    $flagsArr = $sheet.Range($sheet.Cells.Item(1, 8), $sheet.Cells.Item($lastRow, 8)).Value2

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

    $updated = 0
    foreach ($candidate in $candidates) {
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

        $writableRows = $matches | Where-Object {
            $f = $formulaArr[$_, 1]
            $value = $valueArr[$_, 1]
            -not ($f -is [string] -and $f.StartsWith('=')) -and
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

        Write-Host ("{0,-24} <- {1,-28} row {2,-4} C: {3,10} -> {4,10}" -f $candidate.Name, $candidate.Source, $targetRow, $oldValue, $candidate.Amount) -ForegroundColor Green

        if (-not $DryRun) {
            $sheet.Cells.Item($targetRow, 3).Value2 = $candidate.Amount
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
