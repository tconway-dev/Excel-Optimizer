# Define the directory to be scanned
$directory = "C:\ExcelFiles"

# Get a list of all Excel files in the directory and subdirectories
$excelFiles = Get-ChildItem -Path $directory -Filter *.xlsx -Recurse

# Loop through each Excel file
foreach ($file in $excelFiles) {

  # Open the Excel file
  $excel = New-Object -ComObject Excel.Application
  $workbook = $excel.Workbooks.Open($file.FullName)

  # Initialize the report for this file
  $report = "Report for Excel file: $($file.FullName)"

  # Check for common issues

  # Check for excessive blank rows or columns
  $usedRange = $workbook.ActiveSheet.UsedRange
  $usedRows = $usedRange.Rows.Count
  $usedColumns = $usedRange.Columns.Count
  $totalRows = $workbook.ActiveSheet.Rows.Count
  $totalColumns = $workbook.ActiveSheet.Columns.Count
  if ($totalRows - $usedRows - 1 -gt 0) {
    $report += "`nExcessive blank rows detected."
  }
  if ($totalColumns - $usedColumns - 1 -gt 0) {
    $report += "`nExcessive blank columns detected."
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
    $report += "`nUnused worksheets detected."
  }

  # Check for optimization issues

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
#Check for excessive data validation 
  $dataValidationCount = 0
  for ($row = 1; $row -le $usedRows; $row++) {
    for ($column = 1; $column -le $usedColumns; $column++) {
      $cell = $range.Cells.Item($row, $column)
      if ($cell.Validation.Type -ne 0) {
        $dataValidationCount++
      }
    }
  }
  $excessiveDataValidation = $false
  if ($dataValidationCount -gt 100) {
    $excessiveDataValidation = $true
  }
    # Check the workbook's CompatibilityChecker.CheckCompatibility property
    $compatibilityChecker = $workbook.CheckCompatibility
    $result = $compatibilityChecker.CheckCompatibility

    if ($result -eq $false) {
      Write-Output "$filePath is in need of a compatibility update."
    } else {
      Write-Output "$filePath is up-to-date and does not require a compatibility update."
    }

  # Add the results to the report
  $report += [PSCustomObject]@{
    'FileName' = $file.FullName
    'ExcessiveBlankRows' = $excessiveBlankRows
    'ExcessiveBlankColumns' = $excessiveBlankColumns
    'UnusedWorksheets' = $unusedWorksheets
    'ExcessiveFormulas' = $excessiveFormulas
    'ExcessiveConditionalFormatting' = $excessiveConditionalFormatting
    'ExcessiveDataValidation' = $excessiveDataValidation
  }
}  


$report = @()

# For each file in the directory
foreach ($file in $files) {
  # ... (existing code to check for excessive blank rows, columns, unused worksheets, excessive formulas, conditional formatting, and data validation) ...

  # Add the results to the report
  $report += [PSCustomObject]@{
    'FileName' = $file.FullName
    'ExcessiveBlankRows' = $excessiveBlankRows
    'ExcessiveBlankColumns' = $excessiveBlankColumns
    'UnusedWorksheets' = $unusedWorksheets
    'ExcessiveFormulas' = $excessiveFormulas
    'ExcessiveConditionalFormatting' = $excessiveConditionalFormatting
    'ExcessiveDataValidation' = $excessiveDataValidation
  }
}

# Export the report to a pipe-delimited CSV file
$report | Export-Csv -Path 'report.csv' -NoTypeInformation -Delimiter '|'
