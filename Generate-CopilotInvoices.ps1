<#
.SYNOPSIS
    Genererer Copilot-lisens fakturaer per avdeling fra en AD-gruppe.

.DESCRIPTION
    Leser medlemmer av en angitt AD-gruppe, grupperer dem etter avdeling,
    og genererer én Excel-faktura per avdeling. Pris er 3500 kr per medlem.

.PARAMETER GroupName
    Navnet på AD-gruppen som skal hentes.

.PARAMETER OutputFolder
    Mappe der faktura-filer lagres.

.PARAMETER InvoiceNumberStart
    Startnummer for fakturanummerering. Hver avdeling får neste ledige nummer.

.PARAMETER InvoiceDate
    Fakturadato. Standardverdi er dagens dato.

.EXAMPLE
    .\Generate-CopilotInvoices.ps1
    .\Generate-CopilotInvoices.ps1 -InvoiceNumberStart 1050
#>

[CmdletBinding()]
param(
    [string]$GroupName           = "M365_LIC_Copilot_for_M365",

    [string]$OutputFolder        = "C:\temp",

    [int]$InvoiceNumberStart     = 1001,

    [datetime]$InvoiceDate       = (Get-Date),

    [string]$SellerName          = "Universitetet i Bergen",
    [string]$SellerOrgNr         = "N/A",
    [string]$Product             = "Microsoft 365 Copilot lisens",
    [decimal]$PricePerUnit       = 3500
)

#region --- Sjekk avhengigheter ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel-modulen ble ikke funnet. Installerer..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
}
Import-Module ImportExcel -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "ActiveDirectory-modulen er ikke tilgjengelig. Installer RSAT eller kjør skriptet på en domenetilkoblet maskin."
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop
#endregion

#region --- Hent medlemmer fra AD-gruppe ---
Write-Host "Henter medlemmer fra AD-gruppe: '$GroupName'..." -ForegroundColor Cyan

try {
    $members = Get-ADGroupMember -Identity $GroupName -Recursive |
        Where-Object { $_.objectClass -eq 'user' } |
        ForEach-Object {
            Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, Department, departmentNumber, Title, EmailAddress
        }
} catch {
    Write-Error "Kunne ikke hente AD-gruppe '$GroupName': $_"
    exit 1
}

if (-not $members) {
    Write-Warning "Ingen brukere funnet i gruppen '$GroupName'."
    exit 0
}

Write-Host "Fant $($members.Count) bruker(e)." -ForegroundColor Green
#endregion

#region --- Grupper etter avdeling ---
$byDepartment = $members | Group-Object {
    if ($_.Department) { $_.Department.Trim() } else { "Ukjent" }
} | Sort-Object Name

Write-Host "Antall avdelinger: $($byDepartment.Count)" -ForegroundColor Cyan
$byDepartment | ForEach-Object { Write-Host "  - $($_.Name): $($_.Count) bruker(e)" }
#endregion

#region --- Opprett mappe for fakturaer ---
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$OutputFolder = (Resolve-Path $OutputFolder).Path
#endregion

#region --- Generer én Excel-faktura per avdeling ---
$invoiceNumber = $InvoiceNumberStart

foreach ($dept in $byDepartment) {
    $deptName    = $dept.Name
    $users       = $dept.Group | Sort-Object DisplayName
    $quantity    = $users.Count
    $totalAmount = $quantity * $PricePerUnit

    # Hent avdelingsnummer fra første bruker med utfylt departmentNumber
    $deptNumber = ($users | Where-Object { $_.departmentNumber } | Select-Object -First 1).departmentNumber
    if (-not $deptNumber) { $deptNumber = "" }

    $invNumStr   = "FAKTURA-{0:D4}" -f $invoiceNumber
    $safeDept    = $deptName -replace '[\\/:*?"<>|]', '_'
    $outFile     = Join-Path $OutputFolder "$invNumStr`_$safeDept.xlsx"

    # ---- Bygg Excel-arbeidsbok ----
    $xl = Open-ExcelPackage -Path $outFile -KillExcel

    $ws = Add-Worksheet -ExcelPackage $xl -WorksheetName "Faktura"

    # Hjelpefunksjon: skriv verdi til celle med valgfri formatering
    function Set-Cell {
        param($sheet, $row, $col, $value, [switch]$Bold, [switch]$Right)
        $sheet.Cells[$row, $col].Value = $value
        if ($Bold) { $sheet.Cells[$row, $col].Style.Font.Bold = $true }
        if ($Right) { $sheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right }
    }

    # --- Topptekst ---
    Set-Cell $ws 1 1 $SellerName -Bold
    Set-Cell $ws 2 1 "Org.nr: $SellerOrgNr"
    Set-Cell $ws 4 1 "FAKTURA" -Bold
    $ws.Cells[4,1].Style.Font.Size = 18

    Set-Cell $ws 6 1 "Fakturanummer:"  -Bold
    Set-Cell $ws 6 2 $invNumStr
    Set-Cell $ws 7 1 "Fakturadato:"    -Bold
    Set-Cell $ws 7 2 ($InvoiceDate.ToString("yyyy-MM-dd"))
    Set-Cell $ws 8 1 "Forfallsdato:"   -Bold
    Set-Cell $ws 8 2 ($InvoiceDate.AddDays(30).ToString("yyyy-MM-dd"))

    # --- Faktureres til ---
    Set-Cell $ws 6 5 "Faktureres til:"  -Bold
    Set-Cell $ws 7 5 $deptName
    if ($deptNumber) {
        Set-Cell $ws 8 5 "Avd.nr: $deptNumber"
    }

    # --- Kolonneoverskrifter (rad 11) ---
    $tableStart = 11
    $headers = @("Beskrivelse", "Bruker", "E-postadresse", "Stilling")
    for ($c = 1; $c -le $headers.Count; $c++) {
        Set-Cell $ws $tableStart $c $headers[$c - 1] -Bold
        $ws.Cells[$tableStart, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ws.Cells[$tableStart, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
        $ws.Cells[$tableStart, $c].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    }

    # --- Brukerrader ---
    $row = $tableStart + 1
    foreach ($user in $users) {
        $ws.Cells[$row, 1].Value = $Product
        $ws.Cells[$row, 2].Value = if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName }
        $ws.Cells[$row, 3].Value = $user.EmailAddress
        $ws.Cells[$row, 4].Value = $user.Title

        # Annenhver rad får lys bakgrunn
        if ($row % 2 -eq 0) {
            for ($c = 1; $c -le 4; $c++) {
                $ws.Cells[$row, $c].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[$row, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217, 225, 242))
            }
        }
        $row++
    }

    # --- Oppsummering ---
    $summaryRow = $row + 1
    Set-Cell $ws $summaryRow 1 "Antall lisenser:"   -Bold
    Set-Cell $ws $summaryRow 2 $quantity

    $summaryRow++
    Set-Cell $ws $summaryRow 1 "Pris per lisens:"   -Bold
    $ws.Cells[$summaryRow, 2].Value = $PricePerUnit
    $ws.Cells[$summaryRow, 2].Style.Numberformat.Format = "#,##0 kr"

    $summaryRow++
    Set-Cell $ws $summaryRow 1 "Totalt eks. mva.:"  -Bold
    $ws.Cells[$summaryRow, 2].Value = $totalAmount
    $ws.Cells[$summaryRow, 2].Style.Numberformat.Format = "#,##0 kr"
    $ws.Cells[$summaryRow, 2].Style.Font.Bold = $true
    $ws.Cells[$summaryRow, 2].Style.Font.Size = 13

    # Tynn kant rundt oppsummeringsblokken
    $sumRange = $ws.Cells[$summaryRow - 2, 1, $summaryRow, 2]
    $sumRange.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)

    # --- Kolonnebredder ---
    $ws.Column(1).Width = 42
    $ws.Column(2).Width = 30
    $ws.Column(3).Width = 34
    $ws.Column(4).Width = 28
    $ws.Column(5).Width = 22

    # --- Frys øverste rader ---
    $ws.View.FreezePanes($tableStart + 1, 1)

    Close-ExcelPackage $xl -Show:$false

    Write-Host "Opprettet: $outFile  ($quantity brukere, totalt $('{0:N0}' -f $totalAmount) kr)" -ForegroundColor Green
    $invoiceNumber++
}
#endregion

Write-Host ""
Write-Host "Ferdig. $($byDepartment.Count) faktura(er) lagret i: $OutputFolder" -ForegroundColor Cyan
