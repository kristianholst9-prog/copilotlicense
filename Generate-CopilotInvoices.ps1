<#
.SYNOPSIS
    Generates Copilot license invoices per department from an AD group.

.DESCRIPTION
    Reads members of a specified AD group, groups them by Department,
    and generates one Excel invoice per department. Price is 3500 kr per member.

.PARAMETER GroupName
    The name of the AD group to query.

.PARAMETER OutputFolder
    Folder where invoice Excel files are saved. Defaults to .\Invoices.

.PARAMETER InvoiceNumberStart
    Starting invoice number. Each department gets the next sequential number.

.PARAMETER InvoiceDate
    Invoice date. Defaults to today.

.EXAMPLE
    .\Generate-CopilotInvoices.ps1 -GroupName "Copilot-Licenses"
    .\Generate-CopilotInvoices.ps1 -GroupName "Copilot-Licenses" -OutputFolder "C:\Invoices" -InvoiceNumberStart 1050
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$GroupName,

    [string]$OutputFolder = ".\Invoices",

    [int]$InvoiceNumberStart = 1001,

    [datetime]$InvoiceDate = (Get-Date),

    [string]$SellerName    = "Din Organisasjon AS",
    [string]$SellerOrgNr   = "000 000 000",
    [string]$Product        = "Microsoft 365 Copilot lisens",
    [decimal]$PricePerUnit  = 3500
)

#region --- Dependency check ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
}
Import-Module ImportExcel -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "ActiveDirectory module is not available. Install RSAT or run this on a domain-joined machine with AD tools."
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop
#endregion

#region --- Fetch AD group members ---
Write-Host "Querying AD group: '$GroupName'..." -ForegroundColor Cyan

try {
    $members = Get-ADGroupMember -Identity $GroupName -Recursive |
        Where-Object { $_.objectClass -eq 'user' } |
        ForEach-Object {
            Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, Department, departmentNumber, Title, EmailAddress
        }
} catch {
    Write-Error "Failed to query AD group '$GroupName': $_"
    exit 1
}

if (-not $members) {
    Write-Warning "No user members found in group '$GroupName'."
    exit 0
}

Write-Host "Found $($members.Count) member(s)." -ForegroundColor Green
#endregion

#region --- Group by department ---
$byDepartment = $members | Group-Object {
    if ($_.Department) { $_.Department.Trim() } else { "Unknown" }
} | Sort-Object Name

Write-Host "Departments found: $($byDepartment.Count)" -ForegroundColor Cyan
$byDepartment | ForEach-Object { Write-Host "  - $($_.Name): $($_.Count) user(s)" }
#endregion

#region --- Create output folder ---
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$OutputFolder = (Resolve-Path $OutputFolder).Path
#endregion

#region --- Generate one Excel invoice per department ---
$invoiceNumber = $InvoiceNumberStart

foreach ($dept in $byDepartment) {
    $deptName    = $dept.Name
    $users       = $dept.Group | Sort-Object DisplayName
    $quantity    = $users.Count
    $totalAmount = $quantity * $PricePerUnit

    # Try to get department number from first user's departmentNumber attribute
    $deptNumber = ($users | Where-Object { $_.departmentNumber } | Select-Object -First 1).departmentNumber
    if (-not $deptNumber) { $deptNumber = "" }

    $invNumStr   = "FAKTURA-{0:D4}" -f $invoiceNumber
    $safeDept    = $deptName -replace '[\\/:*?"<>|]', '_'
    $outFile     = Join-Path $OutputFolder "$invNumStr`_$safeDept.xlsx"

    # ---- Build Excel workbook ----
    $xl = Open-ExcelPackage -Path $outFile -KillExcel

    $ws = Add-Worksheet -ExcelPackage $xl -WorksheetName "Faktura"

    # Helper: write a cell value and optional bold/style
    function Set-Cell {
        param($sheet, $row, $col, $value, [switch]$Bold, [switch]$Right)
        $sheet.Cells[$row, $col].Value = $value
        if ($Bold) { $sheet.Cells[$row, $col].Style.Font.Bold = $true }
        if ($Right) { $sheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right }
    }

    # --- Header block ---
    Set-Cell $ws 1 1 $SellerName -Bold
    Set-Cell $ws 2 1 "Org.nr: $SellerOrgNr"
    Set-Cell $ws 4 1 "FAKTURA" -Bold
    $ws.Cells[4,1].Style.Font.Size = 18

    Set-Cell $ws 6 1 "Fakturanummer:"  -Bold
    Set-Cell $ws 6 2 $invNumStr
    Set-Cell $ws 7 1 "Fakturadato:"    -Bold
    Set-Cell $ws 7 2 ($InvoiceDate.ToString("yyyy-MM-dd"))
    Set-Cell $ws 8 1 "Forfall:"        -Bold
    Set-Cell $ws 8 2 ($InvoiceDate.AddDays(30).ToString("yyyy-MM-dd"))

    # --- Bill-to block ---
    Set-Cell $ws 6 5 "Faktureres til:"  -Bold
    Set-Cell $ws 7 5 $deptName
    if ($deptNumber) {
        Set-Cell $ws 8 5 "Avd.nr: $deptNumber"
    }

    # --- Line-items header (row 11) ---
    $tableStart = 11
    $headers = @("Beskrivelse", "Bruker", "E-post", "Tittel")
    for ($c = 1; $c -le $headers.Count; $c++) {
        Set-Cell $ws $tableStart $c $headers[$c - 1] -Bold
        $ws.Cells[$tableStart, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$tableStart, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
        $ws.Cells[$tableStart, $c].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    }

    # --- User rows ---
    $row = $tableStart + 1
    foreach ($user in $users) {
        $ws.Cells[$row, 1].Value = $Product
        $ws.Cells[$row, 2].Value = if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName }
        $ws.Cells[$row, 3].Value = $user.EmailAddress
        $ws.Cells[$row, 4].Value = $user.Title

        # Alternate row shading
        if ($row % 2 -eq 0) {
            for ($c = 1; $c -le 4; $c++) {
                $ws.Cells[$row, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[$row, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217, 225, 242))
            }
        }
        $row++
    }

    # --- Summary block ---
    $summaryRow = $row + 1
    Set-Cell $ws $summaryRow 1 "Antall lisenser:"   -Bold
    Set-Cell $ws $summaryRow 2 $quantity

    $summaryRow++
    Set-Cell $ws $summaryRow 1 "Pris per lisens:"   -Bold
    $ws.Cells[$summaryRow, 2].Value = $PricePerUnit
    $ws.Cells[$summaryRow, 2].Style.Numberformat.Format = "#,##0 kr"

    $summaryRow++
    Set-Cell $ws $summaryRow 1 "Total eks. mva.:"   -Bold
    $ws.Cells[$summaryRow, 2].Value = $totalAmount
    $ws.Cells[$summaryRow, 2].Style.Numberformat.Format = "#,##0 kr"
    $ws.Cells[$summaryRow, 2].Style.Font.Bold = $true
    $ws.Cells[$summaryRow, 2].Style.Font.Size = 13

    # Thin border around summary
    $sumRange = $ws.Cells[$summaryRow - 2, 1, $summaryRow, 2]
    $sumRange.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)

    # --- Column widths ---
    $ws.Column(1).Width = 42
    $ws.Column(2).Width = 30
    $ws.Column(3).Width = 34
    $ws.Column(4).Width = 28
    $ws.Column(5).Width = 22

    # --- Freeze top rows and add a thin border under the header row ---
    $ws.View.FreezePanes($tableStart + 1, 1)

    Close-ExcelPackage $xl -Show:$false

    Write-Host "Created: $outFile  ($quantity users, total $('{0:N0}' -f $totalAmount) kr)" -ForegroundColor Green
    $invoiceNumber++
}
#endregion

Write-Host ""
Write-Host "Done. $($byDepartment.Count) invoice(s) written to: $OutputFolder" -ForegroundColor Cyan
