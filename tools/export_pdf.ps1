# Export every sheet of a workbook to one PDF via Excel (each sheet fit to one landscape page).
# Used by tools/render_sheets.py to make README screenshots. Opens read-only; never saves.
param([string]$Path, [string]$Out)
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path, 0, $true); $xl.CalculateFull()
  foreach ($ws in $wb.Worksheets) { $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1 }
  $wb.ExportAsFixedFormat(0, $Out); $wb.Close($false)
} finally { $xl.Quit() }
