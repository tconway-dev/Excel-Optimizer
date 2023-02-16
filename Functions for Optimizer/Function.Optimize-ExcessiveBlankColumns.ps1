Function Optimize-ExcessiveBlankColumns {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Office.Interop.Excel.Workbook]$workbook,
        [Parameter(Mandatory = $true)]
        [int]$maxBlankColumns
    )

    # Get the first worksheet in the workbook
    $worksheet = $workbook.Sheets.Item(1)

    # Get the used range of the worksheet
    $usedRange = $worksheet.UsedRange

    # Get the total number of rows in the used range
    $totalRows = $usedRange.Rows.Count

    # Get the total number of columns in the used range
    $totalColumns = $usedRange.Columns.Count

    # Loop through each column
    for ($column = $totalColumns; $column -gt 1; $column--) {

        # Check if the column is entirely blank
        $isBlank = $true
        for ($row = 1; $row -le $totalRows; $row++) {
            if ($usedRange.Cells.Item($row, $column).Value2 -ne $null) {
                $isBlank = $false
                break
            }
        }

        # If the column is blank and the number of blank columns is greater than the maximum allowed, delete the column
    }
}