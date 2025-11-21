# ===============================
# Merge two Excel workbooks using COM Automation
# Workbook A sheets: X, Y,... Z
# Workbook B sheets: T (possibly multiple)
# Sheet contents are copied verbatim.
# ===============================
# param (
  $wbAPath    = "${env:TEMP}\A.xls"
  $wbBPath    = "${env:TEMP}\B.xls"
  $outputPath = "${env:TEMP}\merged.xls"
#)
# Start Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open both workbooks
    $wbA = $excel.Workbooks.Open($wbAPath)
    $wbB = $excel.Workbooks.Open($wbBPath)

    # Loop through all sheets in workbook B

    foreach ($sheet in $wbB.Worksheets) {
        # evaluate if the sheet is blank
$used = $sheet.UsedRange

# Sheet is empty if UsedRange exists but has no usable cells
$hasData = $used -and ($used.Rows.Count -gt 1 -or $used.Columns.Count -gt 1 -or $used.Value2)

if (-not $hasData) {
    continue
}
	
        # Copy the sheet into workbook A at the end
        $sheet.Copy(
            [Type]::Missing,
            $wbA.Worksheets.Item($wbA.Worksheets.Count)
        )
    }

    # Save as new merged file
    $wbA.SaveAs($outputPath)

}
finally {
    # Clean up properly
    if ($wbB) { $wbB.Close($false) }
    if ($wbA) { $wbA.Close($true) }
    $excel.Quit()

    # Release COM objects fully
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbB) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbA) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Merge completed: $outputPath"
