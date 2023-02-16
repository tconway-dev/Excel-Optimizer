Function Optimize-ExcessiveBlankRows {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Office.Interop.Excel.Workbook]$workbook,
        [Parameter(Mandatory = $true)]
        [int]$maxBlankRows
    )

    # Get the first worksheet in the workbook
    $worksheet = $workbook.Sheets.Item(1)

    # Get the used range of the worksheet
    $usedRange = $worksheet.UsedRange

    # Get the total number of rows in the used range
    $totalRows = $usedRange.Rows.Count

    # Get the total number of columns in the used range
    $totalColumns = $usedRange.Columns.Count

    # Loop through each row
    for ($row = $totalRows; $row -gt 1; $row--) {

        # Check if the row is entirely blank
        $isBlank = $true
        for ($column = 1; $column -le $totalColumns; $column++) {
            if ($usedRange.Cells.Item($row, $column).Value2 -ne $null) {
                $isBlank = $false
                break
            }
        }

        # If the row is blank and the number of blank rows is greater than the maximum allowed, delete the row
        if ($isBlank -and $totalBlankRows + 1 -gt $maxBlankRows) {
            $usedRange.Rows.Item($row).Delete()
            $totalBlankRows++
        } else {
            # If the row is not blank, reset the count of blank rows
            $totalBlankRows = 0
        }
    }
}
