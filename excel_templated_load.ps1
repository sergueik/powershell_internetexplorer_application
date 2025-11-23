# Using Excel.Application COM: Be??st Practices for Preserving Formatting relying on Excel .Copy() 

# param (
  $templatePath = "${env:TEMP}\template.xls"
  $htmlPath = "${env:TEMP}\data.htm"
  $outputPath = "${env:TEMP}\data.xls"
#)

function Apply-CellFormat {
    param(
        [Parameter(Mandatory)]
        $ExcelCell,
        [Parameter(Mandatory)]
        [hashtable]$Format
    )

    # Font properties
    if ($Format.ContainsKey('FontName'))      { $ExcelCell.Font.Name       = $Format['FontName'] }
    if ($Format.ContainsKey('FontSize'))      { $ExcelCell.Font.Size       = $Format['FontSize'] }
    if ($Format.ContainsKey('FontBold'))      { $ExcelCell.Font.Bold       = $Format['FontBold'] }
    if ($Format.ContainsKey('FontItalic'))    { $ExcelCell.Font.Italic     = $Format['FontItalic'] }
    if ($Format.ContainsKey('FontUnderline')) { $ExcelCell.Font.Underline  = $Format['FontUnderline'] }
    if ($Format.ContainsKey('FontColor'))     { $ExcelCell.Font.Color      = $Format['FontColor'] }

    # Alignment
    if ($Format.ContainsKey('HorizontalAlignment')) { $ExcelCell.HorizontalAlignment = $Format['HorizontalAlignment'] }
    if ($Format.ContainsKey('VerticalAlignment'))   { $ExcelCell.VerticalAlignment   = $Format['VerticalAlignment'] }
    if ($Format.ContainsKey('WrapText'))            { $ExcelCell.WrapText            = $Format['WrapText'] }
    if ($Format.ContainsKey('Orientation'))         { $ExcelCell.Orientation         = $Format['Orientation'] }

    # Number format / merging
    if ($Format.ContainsKey('NumberFormat')) { $ExcelCell.NumberFormat = $Format['NumberFormat'] }
    if ($Format.ContainsKey('MergeCells'))   { $ExcelCell.MergeCells   = $Format['MergeCells'] }

    # Interior / fill color
    if ($Format.ContainsKey('InteriorColor')) { $ExcelCell.Interior.Color = $Format['InteriorColor'] }

    # Optional: Borders (flattened hashtable for each border)
    if ($Format.ContainsKey('Borders')) {
        foreach ($borderKey in $Format['Borders'].Keys) {
            $b = $ExcelCell.Borders.Item([int]$borderKey)
            $b.LineStyle = $Format['Borders'][$borderKey].LineStyle
            $b.Color     = $Format['Borders'][$borderKey].Color
            $b.Weight    = $Format['Borders'][$borderKey].Weight
        }
    }
}

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

$format = @()
# $lastCol = $worksheet.Cells.Item(1, $worksheet.Columns.Count).End(-4159).Column
# xlToLeft = -4159
# https://learn.microsoft.com/en-us/office/vba/api/excel.xldirection
# intended to help find the last non-blank cell to the left of a specified range
# NOTE: is unreliable and if returns 1, 
# Poweshell silently downcasts the 
# $format from Array of HashMaps to a HashMAp which ruins the later processing 
# $lastCol = 3
$lastCol = $data[0].Count
write-host ('Loading attributes of {0} cells in row 1' -f $lastcol)
for ($col = 1; $col -le $lastCol; $col++) {
    $cell = $worksheet.Cells.Item(1, $col)
    $format += @{
        'FontName' = $cell.Font.Name
        'FontSize' = $cell.Font.Size
        'Bold' = $cell.Font.Bold
        'Color' = $cell.Font.Color
        'InteriorColor' = $cell.Interior.Color
        'Borders' = $cell.Borders.LineStyle
        'HorizontalAlignment' = $cell.HorizontalAlignment
        'VerticalAlignment' = $cell.VerticalAlignment
        'MergeCells' = $cell.MergeCells
    }
}
$format |convertto-json -Depth 5 | write-host
# TODO: assert
# NOTE - fatal if 'template' is not found
# 0x8002000B Invalid index. 
# NOTE: no header row in template
# NOTE:  Exception from HRESULT: 0x800A01A8 unknown HRESULT

$startRow  = 1

foreach ($item in $data) {
  1..($item.count) | foreach-object { 
    $index = $_
    write-host ('insert "{0}" into {1},{2}' -f $item[$index-1], $startRow, $index)
    # $worksheet.Cells.Item($startRow,$index).Value2 = $item[$index-1] | out-null
		# avoid subtle COM/PowerShell interaction issue with Excel
		# fill data w/o formatting 
		# $worksheet.Cells.Item($startRow,$index).Value2 = $item[$index-1]  }
		$cell = $worksheet.Cells.Item($startRow,$index)
		# put data 
    $cell.Value2 = $item[$index-1]
		$cell.WrapText = $true
		# apply formatting from template (flattened hashtable)
		Apply-CellFormat -ExcelCell $cell -Format $format[$index-1]
		# $cell.Select() # Select cell
    # $selection = $Excel.Selection # Get the selected cell object
		# $range = $worksheet.Range($cell.Address)
		$cell.Borders.LineStyle = 1 
		# $xlContinuous
		# https://learn.microsoft.com/en-us/office/vba/api/excel.xllinestyle
    $cell.Borders.Weight = 2 
		# $xlThin
		# https://learn.microsoft.com/en-us/office/vba/api/excel.xlborderweight

  }
  $startRow++	
}
$worksheet.Name = $title
$workbook.SaveAs($outputPath)
if ($workbook) { 
  $workbook.Close($true) 
}
$excel.Quit()

if($range -ne $null){
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
}
try {
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
} catch {}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
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
