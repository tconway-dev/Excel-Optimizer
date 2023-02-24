#Init Commit 2/10/23
#Rev 8 2/23/23
# Define the directory to be scanned
$directory = "C:\ExcelFiles"

# Get a list of all Excel files in the directory and subdirectories
$excelFiles = Get-ChildItem -Path $directory -Filter *.xlsx -Recurse

# Initialize the report array
$report = @()

# Loop through each Excel file
foreach ($file in $excelFiles) {

  # Open the Excel file
  $excel = New-Object -ComObject Excel.Application
  $workbook = $excel.Workbooks.Open($file.FullName)

  # Initialize the report for this file
  $fileReport = [PSCustomObject]@{
    'FileName' = $file.FullName
    'ExcessiveBlankRows' = $false
    'ExcessiveBlankColumns' = $false
    'UnusedWorksheets' = $false
    'ExcessiveFormulas' = $false
    'ExcessiveConditionalFormatting' = $false
    'ExcessiveDataValidation' = $false
  }

  # Check for common issues

  # Check for excessive blank rows or columns
  $usedRange = $workbook.ActiveSheet.UsedRange
  $usedRows = $usedRange.Rows.Count
  $usedColumns = $usedRange.Columns.Count
  $totalRows = $workbook.ActiveSheet.Rows.Count
  $totalColumns = $workbook.ActiveSheet.Columns.Count
  if ($totalRows - $usedRows - 1 -gt 0) {
    $fileReport.ExcessiveBlankRows = $true
  }
  if ($totalColumns - $usedColumns - 1 -gt 0) {
    $fileReport.ExcessiveBlankColumns = $true
  }

  # Check for unused worksheets
  $sheetCount = $workbook.Sheets.Count
  $usedSheetCount = 0
  for ($i = 1; $i -le $sheetCount; $i++) {
    $sheet = $workbook.Sheets.Item($i)
    if ($sheet.Visible -eq -1) {
      $usedSheetCount++
    }
  }
  if ($sheetCount -gt $usedSheetCount) {
    $fileReport.UnusedWorksheets = $true
  }

  # Check for optimization issues

  # Check for excessive formulas
    # Check for excessive formulas
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
      $report += "`nExcessive formulas detected."
    }
  
    # Check for excessive conditional formatting
    $conditionalFormattingCount = $workbook.ActiveSheet.Range("A1").FormatConditions.Count
    if ($conditionalFormattingCount -gt 100) {
      $report += "`nExcessive conditional formatting detected."
    }
  
    # Check for excessive data validation
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
      $report += "`nExcessive data validation detected."
    }
  
    # Check for compatibility issues
    $compatibilityChecker = $workbook.CheckCompatibility
    $compatibilityIssues = $compatibilityChecker.GetCompatibilityIssues()
    if ($compatibilityIssues.Count -gt 0) {
      $report += "`nCompatibility issues detected:"
      foreach ($issue in $compatibilityIssues) {
        $report += "- $($issue.Name)"
      }
    }
  
    # Close the workbook and Excel application
    $workbook.Close($false)
    $excel.Quit()
  }
  
  # Export the report to a pipe-delimited CSV file
  $report | Export-Csv -Path 'report.csv' -NoTypeInformation -Delimiter '|'
  
