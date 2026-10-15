# Recalculate a valuation workbook in real Excel (Windows + Excel required) and report:
#   - formula error count per sheet (#REF!, #DIV/0!, #NAME?, ...)
#   - the Summary sheet, cross-sheet consistency checks
#   - live-model checks: terminal growth 3% -> 4% and the terminal CapEx switch both move the Summary
# The workbook is opened read-only and never saved.
#
#   powershell -File tools/excel_verify.ps1 -Path output/AEHR_valuation_2026-10-14.xlsx
param([string]$Path)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path, 0, $true)   # read-only
    $xl.CalculateFull()
    $total = 0
    foreach ($ws in $wb.Worksheets) {
        $errs = $null
        try { $errs = $ws.UsedRange.SpecialCells(-4123, 16) } catch { }   # formulas that evaluate to errors
        $nForm = 0
        try { $nForm = $ws.UsedRange.SpecialCells(-4123).Count } catch { }
        if ($errs) {
            $total += $errs.Count
            Write-Output ("{0,-16} formulas={1,-4} ERRORS={2}: {3}" -f $ws.Name, $nForm, $errs.Count, $errs.Address(0,0))
        } else {
            Write-Output ("{0,-16} formulas={1,-4} errors=0" -f $ws.Name, $nForm)
        }
    }
    Write-Output "TOTAL_ERRORS=$total"

    $s = $wb.Worksheets.Item("Summary")
    foreach ($r in 6..26) {
        $lbl = $s.Range("B$r").Text
        if ($lbl) { Write-Output ("  {0,-34} {1,14} {2}" -f $lbl, $s.Range("C$r").Text, $s.Range("D$r").Text) }
    }

    # Cross-sheet consistency checks
    $dcf = $wb.Worksheets.Item("DCF"); $sens = $wb.Worksheets.Item("Sensitivity"); $sc = $wb.Worksheets.Item("Scenarios")
    $dcfv = $s.Range("C7").Value2
    Write-Output ("Sensitivity center {0:N4} vs DCF {1:N4}" -f $sens.Range("E10").Value2, $dcfv)
    Write-Output ("Scenarios base check: {0} | probability check: {1}" -f $sc.Range("B20").Text, $sc.Range("C12").Text)

    # Live-model tests: change inputs, recalc, watch the Summary target (C9) move.
    $a = $wb.Worksheets.Item("Assumptions")
    $t0 = $s.Range("C9").Value2
    $a.Range("B17").Value2 = 0.04; $xl.CalculateFull()
    Write-Output ("TerminalGrowth 3% -> 4%: price target {0:N2} -> {1:N2}" -f $t0, $s.Range("C9").Value2)
    $a.Range("B17").Value2 = 0.03; $a.Range("B18").Value2 = "No"; $xl.CalculateFull()
    Write-Output ("Terminal CapEx=D&A Yes -> No: DCF value {0:N2} -> {1:N2}" -f $dcfv, $s.Range("C7").Value2)
    $wb.Close($false)
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
