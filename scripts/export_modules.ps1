param(
  [string]$WorkbookPath = "template/CapacityPlanner_template.xlsm",
  [string]$DstPath = "src"
)

Write-Host "Exporting VBA modules from '$WorkbookPath' to '$DstPath'..."

if (-not (Test-Path $WorkbookPath)) {
  throw "Workbook not found at '$WorkbookPath'."
}

New-Item -ItemType Directory -Force -Path (Join-Path $DstPath 'modules') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DstPath 'classes') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DstPath 'forms')   | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
  $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath))
  $vbproj = $wb.VBProject

  foreach ($comp in $vbproj.VBComponents) {
    switch ($comp.Type) {
      1 { $target = Join-Path (Join-Path $DstPath 'modules') ($comp.Name + '.bas') }
      2 { $target = Join-Path (Join-Path $DstPath 'classes') ($comp.Name + '.cls') }
      3 { $target = Join-Path (Join-Path $DstPath 'forms')   ($comp.Name + '.frm') }
      Default { continue }
    }
    $comp.Export($target)
  }
} catch {
  Write-Error "Export failed: $($_.Exception.Message). Ensure Excel Trust Center allows VBOM access."
  throw
} finally {
  if ($wb) { $wb.Close($false) }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Export complete."

