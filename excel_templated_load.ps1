# Using Excel.Application COM: Be??st Practices for Preserving formatting relying on Excel .Copy() 

# param (
  $templatePath = "${env:TEMP}\template.xls"
  $htmlPath = "${env:TEMP}\data.htm"
  $outputPath = "${env:TEMP}\data.xls"
#)

function Apply-Cellformat {
    param(
        [Parameter(Mandatory)]
        $ExcelCell,
        [Parameter(Mandatory)]
        [hashtable]$format
    )

    # Font properties
    if ($format.ContainsKey('FontName'))      { $ExcelCell.Font.Name       = $format['FontName'] }
    if ($format.ContainsKey('FontSize'))      { $ExcelCell.Font.Size       = $format['FontSize'] }
    if ($format.ContainsKey('FontBold'))      { $ExcelCell.Font.Bold       = $format['FontBold'] }
    if ($format.ContainsKey('FontItalic'))    { $ExcelCell.Font.Italic     = $format['FontItalic'] }
    if ($format.ContainsKey('FontUnderline')) { $ExcelCell.Font.Underline  = $format['FontUnderline'] }
    if ($format.ContainsKey('FontColor'))     { $ExcelCell.Font.Color      = $format['FontColor'] }

    # Alignment
    if ($format.ContainsKey('HorizontalAlignment')) { $ExcelCell.HorizontalAlignment = $format['HorizontalAlignment'] }
    if ($format.ContainsKey('VerticalAlignment'))   { $ExcelCell.VerticalAlignment   = $format['VerticalAlignment'] }
    if ($format.ContainsKey('WrapText'))            { $ExcelCell.WrapText            = $format['WrapText'] }
    if ($format.ContainsKey('Orientation'))         { $ExcelCell.Orientation         = $format['Orientation'] }

    # Number format / merging
    if ($format.ContainsKey('Numberformat')) { $ExcelCell.Numberformat = $format['Numberformat'] }
    if ($format.ContainsKey('MergeCells'))   { $ExcelCell.MergeCells   = $format['MergeCells'] }

    # Interior / fill color
    if ($format.ContainsKey('InteriorColor')) { $ExcelCell.Interior.Color = $format['InteriorColor'] }

    # Optional: Borders 
		# NOTE: Excel border COM semantics are infamously fragile
    $apply_borders_format = $false
		if ($format.ContainsKey('Borders')) {
			if ( $apply_borders_format ){		
				write-host ('Applying {0} orders formats' -f  $format['Borders'].Count)
				for ($col = 1; $col -lt $format['Borders'].Count; $col++) {
					$b = $ExcelCell.Borders.Item([int]$col)
					# 0x800A03EC 
					# NOTE: orders 11 and 12 almost always throw 0x800A03EC because those indices correspond to:
					# xlDiagonalDown, xlDiagonalUp
					write-host ('LineStyle :{0}' -f $format['Borders'][$col].LineStyle )
					try { $b.LineStyle = $format['Borders'][$col].LineStyle} catch {
						write-host ('Exception: {0} with Borders #{1}' -f $_.Exception.Message, $col)
					}
					write-host ('Weight :{0}' -f $format['Borders'][$col].Weight )
					try { $b.Weight    = $format['Borders'][$col].Weight } catch {
						write-host ('Exception: {0} with Borders #{1}' -f $_.Exception.Message, $col)
					}
					write-host ('Color :{0}' -f $format['Borders'][$col].Color )
					try { $b.Color     = $format['Borders'][$col].Color } catch {
						write-host ('Exception: {0} with Borders #{1}' -f $_.Exception.Message, $col)
					}
					}
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
		# $borderHash = @($null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null,$null)
		$borderHash = New-Object 'System.Object[]' 8
		$borderHash[0] = $null
		for ($cnt = 1; $cnt -lt $borderHash.Count; $cnt++) {
      $border = $cell.Borders.Item($cnt)
			$borderHash[$cnt] = @{
				LineStyle = $border.LineStyle
				Color     = $border.Color
				Weight    = $border.Weight
			}
		}
    $format += @{
        'FontName' = $cell.Font.Name
        'FontSize' = $cell.Font.Size
        'Bold' = $cell.Font.Bold
        'Color' = $cell.Font.Color
        'InteriorColor' = $cell.Interior.Color
        # 'Borders' = $cell.Borders.LineStyle
				'Borders' = $borderHash       # now rich
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
		Apply-Cellformat -ExcelCell $cell -format $format[$index-1]
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
