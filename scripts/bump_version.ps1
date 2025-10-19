param(
  [string]$WorkbookPath = "template/CapacityPlanner_template.xlsm",
  [string]$NewVersion = "0.1.0"
)

Write-Host "Stamping TemplateVersion='$NewVersion' in '$WorkbookPath'..."

if (-not (Test-Path $WorkbookPath)) {
  throw "Workbook not found at '$WorkbookPath'."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
  $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath))
  # Expect a named range 'TemplateVersion' to exist
  $nr = $wb.Names.Item('TemplateVersion')
  $nr.RefersToRange.Value2 = $NewVersion
  $wb.Save()
} catch {
  Write-Error "Version bump failed: $($_.Exception.Message). Ensure 'TemplateVersion' named range exists."
  throw
} finally {
  if ($wb) { $wb.Close($true) }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Version updated."

