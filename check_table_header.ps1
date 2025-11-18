# DOM API-conservative version
#
function Get-ComTableIndexByHeadersGeneric {
  param(
    [string]$filename = 'ch09s33.html'
        [string[]]$expectedHeaders,
        [int]$minColumns = 1
  )
  $rawcontent = Get-Content -Path $filename -Raw
  $rawcontent = (($rawcontent -join '') - replace '\n','' ) -replace '<head>.*</head>' ,''

  $htmlfile = New-Object -ComObject 'HTMLFile'
  $htmlfile.IHTMLDocument2_write($rawcontent )
    $tables = $htmlfile.getElementsByTagName('table')
    for ($i=0; $i -lt $tables.length; $i++) {
        $table = $tables.item($i)
        $rows = $table.rows
        if ($rows.length -eq 0) { continue }

        $headerRow = $rows.item(0)
        $cells = $headerRow.cells
        if ($cells.length -lt $minColumns) { continue }

        # Collect normalized text
        $cellTexts = @()
        for ($c=0; $c -lt $cells.length; $c++) {
            $cellTexts += ($cells.item($c).innerText -replace '\s+','').ToLower()
        }

        # Check all expected headers are present
        $match = $true
        foreach ($h in $expectedHeaders) {
            if (-not ($cellTexts -contains ($h -replace '\s+','').ToLower())) {
                $match = $false
                break
            }
        }

        if ($match) { return $i }
    }

    return -1
}

# Usage example:
$index = Get-ComTableIndexByHeadersGeneric -htmlfile $htmlfile -expectedHeaders $expected -minColumns 3
Write-Host "Matching table index (COM): $index"

