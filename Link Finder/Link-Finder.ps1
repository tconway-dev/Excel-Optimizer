#TMC 2/10/23
#Rev 10 2/18/23
#OOP Change 
class ExcelFile {
    [string] $Name
    [string] $Path
    [string[]] $LinkedWorksheets
    [string[]] $LinkedWorkbooks
    [string[]] $OleDbConnections

    ExcelFile([string] $name, [string] $path) {
        $this.Name = $name
        $this.Path = $path
        $this.LinkedWorksheets = @()
        $this.LinkedWorkbooks = @()
        $this.OleDbConnections = @()
    }

    [void] AddLinkedWorksheet([string] $worksheet) {
        $this.LinkedWorksheets += $worksheet
    }

    [void] AddLinkedWorkbook([string] $workbook) {
        $this.LinkedWorkbooks += $workbook
    }

    [void] AddOleDbConnection([string] $name, [string] $connection) {
        $this.OleDbConnections += "$name: $connection"
    }
}

class ExcelFileProcessor {
    [Microsoft.Office.Interop.Excel.Application] $Excel
    [string] $SearchPath
    [string] $CSVFile

    ExcelFileProcessor([string] $searchPath, [string] $csvFile) {
        # Validate input
        if ([string]::IsNullOrWhiteSpace($searchPath)) {
            throw "Search path cannot be empty or null."
        }
        if ([string]::IsNullOrWhiteSpace($csvFile)) {
            throw "CSV file path cannot be empty or null."
        }

        # Initialize Excel application
        $this.Excel = New-Object -ComObject Excel.Application
        $this.Excel.Visible = $false

        $this.SearchPath = $searchPath
        $this.CSVFile = $csvFile
    }

    [ExcelFile] GetExcelFileInfo([string] $filePath) {
        $file = New-Object ExcelFile -ArgumentList (Split-Path $filePath -Leaf), $filePath

        try {
            # Open Excel workbook
            $workbook = $this.Excel.Workbooks.Open($filePath)

            # Get linked worksheets
            $linkedWorksheets = ($workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks)) | Select-Object -ExpandProperty Text
            foreach ($linkedWorksheet in $linkedWorksheets) {
                $file.AddLinkedWorksheet($linkedWorksheet)
            }

            # Get linked workbooks
            $linkedWorkbooks = ($workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks)) | Select-Object -ExpandProperty FullName
            foreach ($linkedWorkbook in $linkedWorkbooks) {
                $file.AddLinkedWorkbook($linkedWorkbook)
            }

            # Get OLE DB connections
            $oleDbConnections = $workbook.Connections
            foreach ($oleDbConnection in $oleDbConnections) {
                $file.AddOleDbConnection($oleDbConnection.Name, $oleDbConnection.Connection)
            }
        }
        catch {
            Write-Error "Error processing file: $($_.FullName)"
        }
