param(
  [string]$WorkbookPath = "template/CapacityPlanner_template.xlsm",
  [string]$SrcPath = "src"
)

Write-Host "Importing VBA modules from '$SrcPath' into '$WorkbookPath'..."

if (-not (Test-Path $WorkbookPath)) {
  throw "Workbook not found at '$WorkbookPath'. Place your .xlsm there first."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
  $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath))
  $vbproj = $wb.VBProject

  # Remove existing standard modules, classes, and forms to avoid duplicates
  $toRemove = @()
  foreach ($c in $vbproj.VBComponents) {
    if ($c.Type -in 1,2,3) { $toRemove += $c }
  }
  foreach ($c in $toRemove) { $vbproj.VBComponents.Remove($c) }

  function ImportFolder($folder) {
    if (-not (Test-Path $folder)) { return }
    Get-ChildItem -Path $folder -Filter *.bas -ErrorAction SilentlyContinue | ForEach-Object {
      $vbproj.VBComponents.Import($_.FullName) | Out-Null
    }
    Get-ChildItem -Path $folder -Filter *.cls -ErrorAction SilentlyContinue | ForEach-Object {
      $vbproj.VBComponents.Import($_.FullName) | Out-Null
    }
    Get-ChildItem -Path $folder -Filter *.frm -ErrorAction SilentlyContinue | ForEach-Object {
      $vbproj.VBComponents.Import($_.FullName) | Out-Null
    }
  }

  ImportFolder (Join-Path $SrcPath 'modules')
  ImportFolder (Join-Path $SrcPath 'classes')
  ImportFolder (Join-Path $SrcPath 'forms')

  $wb.Save()
} catch {
  Write-Error "Import failed: $($_.Exception.Message). Ensure Excel Trust Center allows VBOM access."
  throw
} finally {
  if ($wb) { $wb.Close($true) }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Import complete."

