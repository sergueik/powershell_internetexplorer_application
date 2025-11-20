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

# ---------------------------
# Logging function
# ---------------------------
function Log {
    param(
        [string]$Message,
        [string]$FilePath = "C:\Temp\ExcelIngest.log",
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Message"
    Add-Content -Path $FilePath -Value $line
}

# ---------------------------
# 1️⃣ Constructor / COM handler
# ---------------------------
function New-ExcelComHandler {
    param([ref]$excel_ref)
    try {
        $excel_ref.Value = New-Object -ComObject Excel.Application
        $excel_ref.Value.Visible = $false
        Log "Excel COM handler created."
    } catch {
        Log "Failed to create Excel COM handler: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# 
# ---------------------------
# 2️⃣ Read HTML → nested array + title
# ---------------------------
function Get-DataFromHtml {
    param(
        [string]$FilePath,
        [ref]$data_ref,
        [ref]$title_ref
    )
    try {
        # Dummy parsing simulating table extraction
        $title_ref.Value = [IO.Path]::GetFileNameWithoutExtension($FilePath)
        $data_ref.Value = @(
            @("Header1","Header2"),
            @("Row1Col1","Row1Col2"),
            @("Row2Col1","Row2Col2")
        )
        Log "Parsed HTML file '$FilePath', title: $($title_ref.Value)"
    } catch {
        Log "Failed to parse HTML file '$FilePath': $($_.Exception.Message)" -Level "ERROR"
        $data_ref.Value = @()
        $title_ref.Value = "SheetError"
    }
}

# ---------------------------
# 3️⃣ Add sheet and populate data
# ---------------------------
function Add-DataSheet {
    param(
        [ref]$workbook_ref,
        [ref]$excel_ref,
        [ref]$data_ref,
        [string]$Title
    )
    try {
        if (-not $workbook_ref.Value) { throw "Workbook reference is null" }
        if (-not $excel_ref.Value) { throw "Excel reference is null" }
        if (-not $data_ref.Value -or $data_ref.Value.Count -eq 0) { 
            Log "Data array is empty. Skipping sheet '$Title'." 
            return 
        }

        # Add new sheet
        if ($workbook_ref.Value.Worksheets.Count -eq 1 -and $workbook_ref.Value.Worksheets.Item(1).UsedRange.Rows.Count -eq 0) {
            $sheet = $workbook_ref.Value.Worksheets.Item(1)
        } else {
            $sheet = $workbook_ref.Value.Worksheets.Add()
        }
        $sheet.Name = $Title

        # Populate cells
        for ($r = 0; $r -lt $data_ref.Value.Count; $r++) {
            for ($c = 0; $c -lt $data_ref.Value[$r].Count; $c++) {
                $sheet.Cells.Item($r+1, $c+1) = $data_ref.Value[$r][$c]
            }
        }
        Log "Sheet '$Title' populated with $($data_ref.Value.Count) rows."
        return $sheet
    } catch {
        Log "Failed to add/populate sheet '$Title': $($_.Exception.Message)" -Level "ERROR"
    }
}

# ---------------------------
# 4️⃣ Close / flush / COM cleanup
# ---------------------------
function Close-ExcelComHandler {
    param(
        [ref]$workbook_ref,
        [ref]$excel_ref,
        [string]$FilePath
    )
    try {
        if ($workbook_ref.Value) {
            $workbook_ref.Value.SaveAs($FilePath)
            $workbook_ref.Value.Close($false)
            Log "Workbook saved as '$FilePath'."
        }
        if ($excel_ref.Value) { $excel_ref.Value.Quit() }
    } catch {
        Log "Error during workbook save/close: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Release COM objects
        if ($workbook_ref.Value) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook_ref.Value) | Out-Null }
        if ($excel_ref.Value) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel_ref.Value) | Out-Null }
        $workbook_ref.Value = $null
        $excel_ref.Value = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Log "COM objects released and garbage collected."
    }
}

# ---------------------------
# Main script block
# ---------------------------

# [ref] variables
$excel_ref = [ref]$null
$workbook_ref = [ref]$null
$data_ref = [ref]$null
$title_ref = [ref]$null

# 1. Open Excel
New-ExcelComHandler -excel_ref $excel_ref

# 2. Add workbook
$workbook_ref.Value = $excel_ref.Value.Workbooks.Add()

# 3. Loop through test files (dummy HTML)
$files = @("C:\Temp\file1.html", "C:\Temp\file2.html")
foreach ($file in $files) {
    Get-DataFromHtml -FilePath $file -data_ref $data_ref -title_ref $title_ref
    Add-DataSheet -workbook_ref $workbook_ref -excel_ref $excel_ref -data_ref $data_ref -Title $title_ref.Value
}

# 4. Save & cleanup
Close-ExcelComHandler -workbook_ref $workbook_ref -excel_ref $excel_ref -FilePath "C:\Temp\output.xlsx"

