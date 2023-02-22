#TMC Inital Commit 2/10/23
#Rev 5 2/20/23
class ExcelFileAnalyzer {
    [string]$Directory
    [string[]]$FileExtensions = @('*.xlsx')
    [bool]$Recurse
    [System.Collections.Generic.List[PSObject]]$Report = @()

    [void] Analyze() {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false

        try {
            $excelFiles = Get-ChildItem -Path $this.Directory -Filter $this.FileExtensions -Recurse:$this.Recurse

            foreach ($file in $excelFiles) {
                Write-Host "Analyzing $($file.FullName)"

                $workbook = $excel.Workbooks.Open($file.FullName)

                try {
                    $reportRow = $this.AnalyzeWorkbook($workbook, $file.FullName)
                    $this.Report.Add($reportRow)
                }
                finally {
                    $workbook.Close($false)
                }
            }
        }
        finally {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Remove-Variable excel
        }
    }

    [PSObject] AnalyzeWorkbook($workbook, $fileName) {
        $sheet = $workbook.ActiveSheet
        $usedRange = $sheet.UsedRange
        $usedRows = $usedRange.Rows.Count
        $usedColumns = $usedRange.Columns.Count
        $totalRows = $sheet.Rows.Count
        $totalColumns = $sheet.Columns.Count

        $excessiveBlankRows = $totalRows - $usedRows - 1 -gt 0
        $excessiveBlankColumns = $totalColumns - $usedColumns - 1 -gt 0

        $sheetCount = $workbook.Sheets.Count
        $usedSheetCount = 0
        for ($i = 1; $i -le $sheetCount; $i++) {
            $sheet = $workbook.Sheets.Item($i)
            if ($sheet.Visible -eq -1) {
                $usedSheetCount++
            }
        }
        $unusedWorksheets = $sheetCount - $usedSheetCount -gt 0

        $formulaCount = 0
        $range = $sheet.UsedRange
        for ($row = 1; $row -le $usedRows; $row++) {
            for ($column = 1; $column -le $usedColumns; $column++) {
                $cell = $range.Cells.Item($row, $column)
                if ($cell.HasFormula) {
                    $formulaCount++
                }
            }
        }
        $excessiveFormulas = $formulaCount -gt 5000

        $conditionalFormattingCount = $sheet.Cells.FormatConditions.Count
        $excessiveConditionalFormatting = $conditionalFormattingCount -gt 100

        $dataValidationCount = 0
        $range = $sheet.UsedRange
        for ($row = 1; $row -le $usedRows; $row++) {
            for ($column = 1; $column -le $usedColumns; $column++) {
                $cell = $range.Cells.Item($row, $column)
                if ($cell.Validation -and $cell.Validation.Type -ne 0) {
                    $dataValidationCount++
                }
            }
        }
        $excessiveDataValidation = $dataValidationCount -gt 100

        # Check the workbook's CompatibilityChecker.CheckCompatibility property
        $compatibilityChecker = $workbook.CheckCompatibility
        $needsCompatibilityUpdate = !$compatibilityChecker.CheckCompatibility

       
