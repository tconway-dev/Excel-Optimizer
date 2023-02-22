#Init Commit 2/10/23
#Rev 6 2/21/23

class ExcelFileAnalyzer {
  [string]$FilePath
  [bool]$ExcessiveBlankRows
  [bool]$ExcessiveBlankColumns
  [bool]$UnusedWorksheets
  [bool]$ExcessiveFormulas
  [bool]$ExcessiveConditionalFormatting
  [bool]$ExcessiveDataValidation

  ExcelFileAnalyzer($filePath) {
    $this.FilePath = $filePath
  }

  [void]Analyze() {
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($this.FilePath)

    $usedRange = $workbook.ActiveSheet.UsedRange
    $usedRows = $usedRange.Rows.Count
    $usedColumns = $usedRange.Columns.Count
    $totalRows = $workbook.ActiveSheet.Rows.Count
    $totalColumns = $workbook.ActiveSheet.Columns.Count

    if ($totalRows - $usedRows - 1 -gt 0) {
      $this.ExcessiveBlankRows = $true
    }

    if ($totalColumns - $usedColumns - 1 -gt 0) {
      $this.ExcessiveBlankColumns = $true
    }

    $sheetCount = $workbook.Sheets.Count
    $usedSheetCount = 0

    for ($i = 1; $i -le $sheetCount; $i++) {
      $sheet = $workbook.Sheets.Item($i)

      if ($sheet.Visible -eq -1) {
        $usedSheetCount++
      }
    }

    if ($sheetCount -gt $usedSheetCount) {
      $this.UnusedWorksheets = $true
    }

    $formulaCount = 0
    $range = $workbook.ActiveSheet.UsedRange

    for ($row = 1; $row -le $usedRows; $row++) {
      for ($column = 1; $column -le $usedColumns; $column++) {
        $cell = $range.Cells.Item($row, $column)

        if ($cell.HasFormula) {
          $formulaCount++
        }
      }
    }

    if ($formulaCount -gt 5000) {
      $this.ExcessiveFormulas = $true
    }

    $conditionalFormattingCount = $workbook.ActiveSheet.Range("A1").FormatConditions.Count

    if ($conditionalFormattingCount -gt 100) {
      $this.ExcessiveConditionalFormatting = $true
    }

    $dataValidationCount = 0

    for ($row = 1; $row -le $usedRows; $row++) {
      for ($column = 1; $column -le $usedColumns; $column++) {
        $cell = $range.Cells.Item($row, $column)

        if ($cell.Validation.Type -ne 0) {
          $dataValidationCount++
        }
      }
    }

    if ($dataValidationCount -gt 100) {
      $this.ExcessiveDataValidation = $true
    }
  }
}

$directory = "C:\ExcelFiles"
$excelFiles = Get-ChildItem -Path $directory -Filter *.xlsx -Recurse
$analyzers = @()

foreach ($file in $excelFiles) {
  $analyzer = [ExcelFileAnalyzer]::new($file.FullName)
  $analyzer.Analyze()
  $analyzers += $analyzer
}

$report = $analyzers | Select-Object FilePath, ExcessiveBlankRows, ExcessiveBlankColumns, UnusedWorksheets, ExcessiveFormulas, ExcessiveConditionalFormatting, ExcessiveDataValidation

$report | Export-Csv -Path "C:\ExcelFilesReport.csv" -NoTypeInformation

