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
$log_file = "$PSScriptRoot\excel_aggregator.log"

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
    param(
        [ref]$excel_ref,
        [ref]$workbook_ref
    )

    log 'Creating Excel.Application COM object...'

    try {
        $excel = new-object -comobject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()

        $excel_ref.Value = $excel
        $workbook_ref.Value = $workbook

        log 'Excel and workbook created'
    }
    catch {
        log "ERROR creating Excel: $($_.Exception.Message)"
        throw
    }
}

# ============================================================
# HTML → array_of_arrays  (trivial DOM using HTMLFile)
# ============================================================

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

# ============================================================
# Add sheet and populate
# ============================================================

function add_datasheet {
    param(
        [ref]$workbook_ref,
        [ref]$excel_ref,
        [ref]$data_ref,
        [string]$title
    )

    if (-not $workbook_ref.Value) { throw 'Workbook reference is null' }
    if (-not $excel_ref.Value)    { throw 'Excel ref null' }

    if (-not $data_ref.Value) {
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

    $data_ref.Value | foreach-object {
        $row = $_
        $col_index = 1

        $row | foreach-object {
            $sheet.Cells.Item($row_index, $col_index).Value2 = $_
            $col_index++
        }

        $row_index++
    }
}

# ============================================================
# Cleanup / Save
# ============================================================

function close_excel {
    param(
        [ref]$workbook_ref,
        [ref]$excel_ref,
        [string]$filepath
    )

    log "Saving workbook to $filepath"

    try {
        $workbook_ref.Value.SaveAs($filepath)
    }
    catch {
        log "ERROR saving workbook: $($_.Exception.Message)"
    }

    log 'Closing Excel...'

    try { $workbook_ref.Value.Close() } catch {}
    try { $excel_ref.Value.Quit() } catch {}

    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook_ref.Value) | out-null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel_ref.Value)     | out-null } catch {}

    $workbook_ref.Value = $null
    $excel_ref.Value     = $null

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    log 'Excel cleanup completed.'
}

# ============================================================
# MAIN (Pipeline version)
# ============================================================

log '=== Excel Aggregation Start (Version 2) ==='

$excel_ref    = [ref]$null
$workbook_ref = [ref]$null

create_excel -excel_ref $excel_ref -workbook_ref $workbook_ref

get-childitem "$PSScriptRoot\input" -filter '*.html' |
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

$output = "$PSScriptRoot\output.xlsx"
close_excel -workbook_ref $workbook_ref -excel_ref $excel_ref -filepath $output

log '=== Completed ==='
