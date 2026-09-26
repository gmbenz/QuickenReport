<#
.SYNOPSIS
    Parses a Quicken report (copied from the Quicken report grid) into structured data.

.DESCRIPTION
    Defines Get-QuickenReportData, which turns the tab-delimited clipboard/text export
    from Quicken into an ordered list of entries (groups, category subtotals, and
    transactions) plus the report title, date range, and overall total. Shared by
    Show-QuickenReport.ps1 (console display) and Update-BudgetFromQuicken.ps1
    (spreadsheet update).

.PARAMETER Path
    Optional path to a text file to read instead of the clipboard (useful for testing).

.PARAMETER ShowVoids
    Include Quicken's zero-amount "**VOID**" placeholder rows (hidden by default).
#>

function Get-QuickenReportData {
    [CmdletBinding()]
    param(
        [string]$Path,
        [switch]$ShowVoids
    )

    if ($Path) {
        if (-not (Test-Path $Path)) { throw "File not found: $Path" }
        $raw = Get-Content -Path $Path -Raw
    } else {
        $raw = Get-Clipboard -Raw
    }

    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "Clipboard is empty or does not contain text. Copy a Quicken report first."
    }

    $lines = $raw -split "`r?`n"

    $dateRegex = '^\d{1,2}/\d{1,2}/\d{4}$'

    $title = $null
    $dateRange = $null
    $overallTotal = $null
    $sawHeaderRow = $false
    $entries = [System.Collections.Generic.List[object]]::new()
    $currentCategory = $null
    $currentGroup = $null
    $negatedGroups = 'EXPENSES', 'TRANSFERS'
    $reportGroups = 'INCOME', 'EXPENSES', 'TRANSFERS'

    $dateIndex = $null
    $accountIndex = $null
    $descriptionIndex = $null
    $memoIndex = $null
    $amountIndex = $null

    # Pre-scan the leading lines (before the transaction header) for the title / date range
    foreach ($line in $lines) {
        if ($line -match "`t") { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if (-not $title) { $title = $line.Trim() }
        elseif (-not $dateRange) { $dateRange = $line.Trim() }
    }

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $cols = $line -split "`t"

        # Skip the title/date-range lines and capture the column indexes from the header row.
        if (-not $sawHeaderRow) {
            if ($cols.Count -le 1) { continue }
            $headerNames = @($cols | ForEach-Object { $_.Trim() })
            if ($headerNames -contains 'Date' -and $headerNames -contains 'Amount') {
                $dateIndex = [Array]::IndexOf($headerNames, 'Date')
                $accountIndex = [Array]::IndexOf($headerNames, 'Account')
                $descriptionIndex = [Array]::IndexOf($headerNames, 'Description')
                $memoIndex = [Array]::IndexOf($headerNames, 'Memo')
                $amountIndex = [Array]::IndexOf($headerNames, 'Amount')
                $sawHeaderRow = $true
            }
            continue
        }

        if ($cols.Count -le $amountIndex) { continue }

        $col1 = $cols[$dateIndex].Trim()
        $col2 = $cols[$accountIndex].Trim()
        $col3 = $cols[$descriptionIndex].Trim()
        $memo = if ($memoIndex -ge 0 -and $cols.Count -gt $memoIndex) { $cols[$memoIndex].Trim() } else { $null }
        $amountText = $cols[$amountIndex].Trim()
        if ([string]::IsNullOrWhiteSpace($col1) -or [string]::IsNullOrWhiteSpace($amountText)) { continue }

        $amount = [double]::Parse(($amountText -replace ',', ''), [System.Globalization.CultureInfo]::InvariantCulture)

        if ($col1 -match $dateRegex) {
            # Transaction row
            if (-not $ShowVoids -and $col3 -eq '**VOID**' -and $amount -eq 0) { continue }

            if ($currentGroup -in $negatedGroups) { $amount = -$amount }
            $entries.Add([PSCustomObject]@{
                Type        = 'Transaction'
                Category    = $currentCategory
                Section     = $currentGroup
                Date        = $col1
                Account     = $col2
                Description = $col3
                Memo        = $memo
                Amount      = $amount
            })
        }
        elseif ($col1 -eq 'OVERALL TOTAL') {
            $overallTotal = $amount
        }
        else {
            # Group / category subtotal row
            $isGroup = ($col1 -ceq $col1.ToUpperInvariant())
            if ($isGroup -and $col1 -in $reportGroups) { $currentGroup = $col1 }
            $currentCategory = $col1

            # Quicken reports expenses/transfers as money leaving accounts; negate so they match the budget sheet's positive "spent" convention.
            if ($currentGroup -in $negatedGroups) { $amount = -$amount }
            $entries.Add([PSCustomObject]@{
                Type           = if ($isGroup) { 'Group' } else { 'Category' }
                Name           = $col1
                Amount         = $amount
            })
        }
    }

    [PSCustomObject]@{
        Title        = $title
        DateRange    = $dateRange
        OverallTotal = $overallTotal
        Entries      = $entries
    }
}
