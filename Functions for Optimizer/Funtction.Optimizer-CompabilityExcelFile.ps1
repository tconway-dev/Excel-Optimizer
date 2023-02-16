Function Upgrade-ExcelFile {
  param (
    [string]$filePath
  )
  
  # Open the Excel application
  $excel = New-Object -ComObject Excel.Application

  # Open the old Excel file
  $workbook = $excel.Workbooks.Open($filePath)

  # Save the file as a new file in the latest file format
  $newFilePath = [System.IO.Path]::ChangeExtension($filePath, ".xlsx")
  $workbook.SaveAs($newFilePath, 51)

  # Close the workbook and the Excel application
  $workbook.Close()
  $excel.Quit()

  # Release the COM object
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
