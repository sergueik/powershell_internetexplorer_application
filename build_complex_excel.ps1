<#
Excel ingestion skeleton
Author: Serguei Kouzmine
Date: 2025-11-20
Notes:
- Fully modular [ref] design
- Dummy HTML parsing
- Logging to file
- COM cleanup
- Powershell ISE friendly Monolithic 
- Safe for testing in PowerShell ISE
#>
<#
Excel ingestion skeleton
Author: Serguei Kouzmine
Date: 2025-11-20
#>

# ============================================================
# Version 2: Pipeline-based Excel Aggregator (refined conventions)
# ============================================================

# ---------- configuration ----------
$log_file = "${env:TEMP}\excel_aggregator.log"

function log {
    param([string]$msg)
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    "$timestamp $msg" | out-file -filepath $log_file -append -encoding utf8
    write-host $msg
}

# ============================================================
# Excel / COM creation
# ============================================================

function create_excel {
	log 'Creating Excel.Application COM object...'

	try {
		$excel = new-object -comobject Excel.Application
		$excel.Visible = $false
		$excel.DisplayAlerts = $false

		$excel_ref = ([ref] $excel)

		log ($excel.getType().Name + ' created')
		# log 'Excel and workbook created'
	}
	catch {
		#
		log ( 'ERROR creating Excel: ' + $_.Exception.Message )
		throw
	}
	return $excel_ref
}

function close_html  { 
    param(
        [string]$htmlfile_ref
    )

    log 'Closing HTMLFile...'
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($htmlfile_ref.Value) | out-null
    } catch {
    }
    try {
        $htmlfile_ref.Value = $null
        # NOTE if COM Object is  released, the ref value will no longer be present:
        # The property 'Value' cannot be found on this object. 
        # Verify that the property exists and can be set
    } catch {
    }

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    log 'HTMLFile cleanup completed.'

}
function create_html {
    param(
        [string]$filename = 'ch09s33.html'
    )
 
  $rawcontent = Get-Content -Path $filename -Raw
  $rawcontent = (($rawcontent -join '') -replace '\n','' ) -replace '<head>.*</head>' ,''

  $htmlfile = New-Object -ComObject 'HTMLFile'
  $htmlfile.IHTMLDocument2_write($rawcontent )
  log ($htmlfile.getType().Name + ' created')
  return ([ref] $htmlfile )
 }


# ============================================================
# HTML → array_of_arrays  (trivial DOM using HTMLFile)
# ============================================================
<#
function read_html_table {
    param(
        [string]$filepath,
        [ref]$title_ref,
        [ref]$data_ref
    )

    log "Reading HTML: $filepath"

    $data = New-Object System.Collections.ArrayList

    try {
        $html = new-object -com 'HTMLFile'
        $raw = get-content -literalpath $filepath -raw
        $html.IHTMLDocument2_write($raw)
        $html.close()

        # ---- title ----
        try {
            $title = $html.getElementsByTagName('title')[0].innerText
        } catch {
            $title = 'Untitled'
        }
        $title_ref.Value = $title

        # ---- rows ----
        $rows = $html.getElementsByTagName('tr')

        $rows | foreach-object {
            $r = $_
            $cells = $r.getElementsByTagName('td')
            $row_data = @()

            $cells | foreach-object {
                $row_data += $_.innerText
            }

            if ($row_data.Count -gt 0) {
                [void]$data.Add($row_data)
            }
        }

        $data_ref.Value = $data
    }
    catch {
        log "ERROR parsing HTML: $($_.Exception.Message)"
        throw
    }
}
#>
# ============================================================
# Add sheet and populate
# ============================================================
<#
function add_datasheet {
	param(
		[ref]$workbook_ref,
		[ref]$excel_ref,
		[ref]$data_ref,
		[string]$title
	)

    if (-not $workbook_ref.Value) { throw 'Workbook reference is null' }
    if (-not $excel_ref.Value)    { throw 'Excel ref null' }

    if ( -not $data_ref.Value ) {
			log "No data – skipping sheet $title"
			return
    }

    log "Adding sheet: $title"

    try {
			$sheet = $workbook_ref.Value.Worksheets.Add()
			$sheet.Name = $title
    }
    catch {
			log "WARNING: Failed to set sheet name ($title). Using default."
			$sheet = $workbook_ref.Value.Worksheets.Add()
    }

    $row_index = 1

    $data_ref.Value | 
			foreach-object {
        $row = $_
        $col_index = 1

        $row | foreach-object {
					$sheet.Cells.Item($row_index, $col_index).Value2 = $_
					$col_index++
        }

        $row_index++
    }
}
#>
# ============================================================
# Cleanup / Save
# ============================================================

function close_excel {
	param(
		[ref]$workbook_ref = ([ref]$null),
		[ref]$excel_ref,
		[string]$filepath = $null
	)
	if ($workbook_ref -ne $null){
		$c1 = $true 
	} else {
		$c1 = $false
	}


	if ($c1){
		try {
			log ('Saving workbook to {0}' -f $filepath)
			$workbook_ref.Value.SaveAs($filepath)
		} catch {
			log  ('ERROR saving workbook: ' + $_.Exception.Message )
		}

	}
	log 'Closing Excel...'
	if ($c1){
		try { 
			$workbook_ref.Value.Close() 
		} catch {
		}
	}
	try { $excel_ref.Value.Quit() } catch {}

	if ($c1){
		try { 
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook_ref.Value) | out-null
		} catch {
		}
	}
	try { 
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel_ref.Value) | out-null 
	} catch {
	}

	if ($c1){
		$workbook_ref.Value = $null
	}
	$excel_ref.Value     = $null

	[gc]::Collect()
	[gc]::WaitForPendingFinalizers()

	log 'Excel cleanup completed.'
}

# ============================================================
# MAIN (Pipeline version)
# ============================================================

$workdir = 'C:\developer\sergueik\powershell_internetexplorer_application'
cd $workdir

$htmlfile_ref = create_html -filename 'ch09s33.html'
$htmlfile = $htmlfile_ref.value
close_html -htmlfile_ref $htmlfile_ref
# write-output $htmlfile

log '=== Excel Aggregation Start (Version 2) ==='
$excel_ref = create_excel
$excel = $excel_ref.Value
$excel_filepath = "${env:TEMP}\output.xlsx"

$excel.Visible = $false
if ($excel.workbooks.Count -eq 0 ){ 
	log 'adding workbook'
  ($workbook = $excel.workbooks.Add()) | out-null
} else {
	log ('loading workbook from {0}' -f $excel_filepath )
	$workbook = $excel.workbooks.Item(0)
}
if (test-path -path $excel_filepath) {
	log ('Opening workbook from {0}' -f $excel_filepath )
	$excel.workbooks.Open($excel_filepath) | out-null
}
# TODO: conditional
0..4 |
foreach-object {
	$cnt = $_
	$index = $workbook.Worksheets.Count
	# NOTE: this area does not work
	write-host('count of sheets: {0}' -f $index)
	$worksheet = $workbook.Worksheets.Add()
	$index = $workbook.Worksheets.Count
	$worksheet = $workbook.Worksheets.Item($index)
	$worksheet.Activate()
	$title = 'sheet ' + $cnt
	$worksheet.Name = 'MyData'
	$index = $workbook.Worksheets.Count
	write-host('count of sheets: {0}' -f $index)


	$worksheet.Cells.Item(1,1) = 'test'
}

$workbook_ref = [ref]$workbook

close_excel -excel_ref $excel_ref -workbook_ref $workbook_ref -filePath $excel_filepath
exit

get-childitem $datadir -filter '*.html' |
foreach-object {
	$file_path = $_.FullName
	log "Processing file: $file_path"

	$title_ref = [ref]$null
	$data_ref  = [ref]$null

	if ($file_path -match 'STOP_ALL') {
		log 'Fatal condition detected. Exiting.'
		return
	}

	if ($file_path -match 'SKIP') {
		log 'Skipping due to SKIP marker'
		continue
	}

	read_html_table -filepath $file_path -title_ref $title_ref -data_ref $data_ref
	add_datasheet -workbook_ref $workbook_ref -excel_ref $excel_ref -data_ref $data_ref -title $title_ref.Value
}


log '=== Completed ==='
