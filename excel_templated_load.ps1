# Using Excel.Application COM: Be??st Practices for Preserving Formatting relying on Excel .Copy() 

# param (
  $templatePath = "${env:TEMP}\template.xls"
  $htmlPath = "${env:TEMP}\data.htm"
  $outputPath = "${env:TEMP}\data.xls"
#)
$data = New-Object System.Collections.ArrayList
$html = new-object -com 'HTMLFile'
$raw = get-content -literalpath $htmlPath -raw
$html.IHTMLDocument2_write($raw)
$html.close()

try {
  $title = $html.getElementsByTagName('title')[0].innerText
} catch {
  $title = 'Untitled'
}

$rows = $html.getElementsByTagName('tr')

$rows | 
foreach-object {
  $element = $_
  $cells = $element.getElementsByTagName('td')
  $row_data = @()

  $cells | 
  foreach-object {
    $row_data += $_.innerText
    write-host ('read {0}' -f $row_data[$row_data.count -1])
  }

  if ($row_data.Count -gt 0) {
    [void]$data.Add($row_data)
  }
}
# cleanup
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($html) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$excel = new-object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($templatePath)
$worksheet = $workbook.Sheets.Item('template')
# NOTE - fatal if 'template' is not found
# 0x8002000B Invalid index. 
# NOTE: no header row in template
$sampleRow = 1
$startRow  = 2

# NOTE:  Exception from HRESULT: 0x800A01A8 unknown HRESULT

foreach ($item in $data) {
  $worksheet.Rows.Item($sampleRow).Copy() |out-null
  write-host('copied {0} to {1}' -f $sampleRow, $startRow)
  $worksheet.Rows.Item($startRow).PasteSpecial(-4163) |out-null
	# 4163 xlPasteFormats
  # https://learn.microsoft.com/en-us/office/vba/api/excel.xlpastetype
  # https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial

  1..($item.count) | foreach-object { 
    $index = $_
    write-host ('insert "{0}" into {1},{2}' -f $item[$index-1], $startRow, $index)
    # $worksheet.Cells.Item($startRow,$index).Value2 = $item[$index-1] | out-null
		$worksheet.Cells.Item($startRow,$index).Value2 = $item[$index-1] | out-null
  }
  $startRow++	
}
$worksheet.Name = $title
# Save as new file
$workbook.SaveAs($outputPath)
# Clean up properly
# 
if ($workbook) { 
  $workbook.Close($true) 
}
$excel.Quit()

# Release COM objects fully
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

<# this preserves:

* Border style
* Background color
* Font and size
* Column width (because it exists already)
* Merged cells
*Text alignment
.Interior, .Font, .Borders, .ColumnWidth — all that formatting stays from the template.
    and is
    * less error-prone than enumerating cell style attributes manually
    * simpler to write
    * formatting fidelity nearly 100%)
#>
