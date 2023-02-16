#TMC 2/10/23
#Rev 5 2/15/23
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False

$SearchPath = "C:\Users\<Username>\<Path to directory>"
$CSVFile = "C:\Users\<Username>\<Path to output file>.csv"

try {
    # Check if the output file already exists and if so, delete it.
    if (Test-Path $CSVFile) {
        Remove-Item $CSVFile
    }

    Get-ChildItem -Path $SearchPath -Filter *.xlsx -Recurse | ForEach-Object {
        $Workbook = $Excel.Workbooks.Open($_.FullName)

        $Links = @()
        $Links += ($Workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks)) | Select-Object -ExpandProperty Text

        if ($Links.Count -gt 0) {
            [PSCustomObject]@{
                FileName = $_.Name
                FilePath = $_.FullName
                Links = ($Links -Join ";")
            } | Export-Csv -Path $CSVFile -Append -NoTypeInformation -Delimiter ","
        }

        $Workbook.Close($False)
        Remove-Variable Workbook
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel
}
