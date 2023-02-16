Function Optimize-UnusedWorksheets {
  param(
    [string]$FilePath
  )

  # Load the Excel Assembly
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $workbook = $excel.Workbooks.Open($FilePath)

  # Get all the worksheets in the workbook
  $worksheets = $workbook.Worksheets

  # Loop through each worksheet and check if it is unused
  for ($i = $worksheets.Count; $i -gt 0; $i--) {
    $worksheet = $worksheets.Item($i)
    $usedRange = $worksheet.UsedRange
    $usedRows = $usedRange.Rows.Count
    $usedColumns = $usedRange.Columns.Count

    # If the worksheet is unused, delete it
    if ($usedRows -eq 0 -and $usedColumns -eq 0) {
      $worksheet.Delete()
    }
  }

  # Save the changes and close the workbook
  $workbook.Save()
  $workbook.Close()
  $excel.Quit()

  # Release the Excel objects
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheets) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
