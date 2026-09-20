<#
.SYNOPSIS
    Reads a Quicken report (copied from the Quicken report grid) out of the clipboard
    and displays it as a readable, colorized console report.

.DESCRIPTION
    Quicken reports copied to the clipboard come out as tab-delimited text with rows like:
        <title>
        <blank>
        <date range>
        <blank>
        \tDate\tAccount\tDescription\tAmount\t          (header row)
        \tINCOME\t\t\t338.16\t                          (group/category subtotal row)
        \t9/3/2026\tChecking\tAlly\t0.16\t               (transaction row)
        ...
        \tOVERALL TOTAL\t\t\t-7,369.25\t

    This script parses that structure and prints:
      - the report title / date range
      - each group (INCOME, EXPENSES, TRANSFERS, ...) and category with its subtotal
      - each transaction indented underneath its category
      - the overall total

.PARAMETER Path
    Optional path to a text file to read instead of the clipboard (useful for testing).

.PARAMETER ShowVoids
    Include Quicken's zero-amount "**VOID**" placeholder rows (hidden by default).

.PARAMETER NoColor
    Disable colored output.

.EXAMPLE
    .\Show-QuickenReport.ps1
    Reads the report currently on the clipboard and prints it.
#>
[CmdletBinding()]
param(
    [string]$Path,
    [switch]$ShowVoids,
    [switch]$NoColor
)

function Format-Amount([double]$Value) {
    $Value.ToString('N2', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Write-Colored([string]$Text, [string]$Color) {
    if ($NoColor) { Write-Host $Text } else { Write-Host $Text -ForegroundColor $Color }
}

. (Join-Path $PSScriptRoot 'Parse-QuickenReport.ps1')

$amountWidth = 12
$nameWidth = 24

$data = Get-QuickenReportData -Path $Path -ShowVoids:$ShowVoids

Write-Host ""
if ($data.Title) { Write-Colored $data.Title 'White' }
if ($data.DateRange) { Write-Colored $data.DateRange 'Gray' }

foreach ($entry in $data.Entries) {
    switch ($entry.Type) {
        'Transaction' {
            $line1 = "    {0,-10} {1,-20} {2,-24} {3,$amountWidth}" -f $entry.Date, $entry.Account, $entry.Description, (Format-Amount $entry.Amount)
            $color = if ($entry.Amount -lt 0) { 'Red' } else { 'DarkGreen' }
            Write-Colored $line1 $color
        }
        'Group' {
            $line1 = "{0,-$nameWidth} {1,$amountWidth}" -f $entry.Name, (Format-Amount $entry.Amount)
            Write-Host ""
            Write-Colored $line1 'Cyan'
        }
        'Category' {
            $line1 = "  {0,-$($nameWidth - 2)} {1,$amountWidth}" -f $entry.Name, (Format-Amount $entry.Amount)
            $color = if ($entry.Amount -lt 0) { 'Yellow' } else { 'Green' }
            Write-Colored $line1 $color
        }
    }
}

Write-Host ""
if ($null -ne $data.OverallTotal) {
    $color = if ($data.OverallTotal -lt 0) { 'Red' } else { 'Green' }
    Write-Colored ("{0,-$nameWidth} {1,$amountWidth}" -f 'OVERALL TOTAL', (Format-Amount $data.OverallTotal)) $color
}
Write-Host ""
