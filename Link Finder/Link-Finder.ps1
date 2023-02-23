#TMC 2/10/23
#Rev 11 2/22/23
#OOP Change 
class ExcelFileAnalyzer {
    [System.Collections.ArrayList]$results

    ExcelFileAnalyzer() {
        $this.results = New-Object System.Collections.ArrayList
    }

    [void]AnalyzeDirectory([string]$path) {
        $files = Get-ChildItem -Path $path -Include *.xls,*.xlsx -Recurse
        foreach ($file in $files) {
            $workbook = $this.OpenWorkbook($file.FullName, $true)
            if ($workbook) {
                $linkedWorksheets = $this.GetLinkedWorksheets($workbook)
                $linkedWorkbooks = $this.GetLinkedWorkbooks($workbook)
                $connections = $this.GetOleDbConnections($workbook)

                if ($linkedWorksheets.Count -gt 0 -or $linkedWorkbooks.Count -gt 0 -or $connections.Count -gt 0) {
                    $result = [PSCustomObject]@{
                        FileName = $file.Name
                        FilePath = $file.FullName
                        LinkedWorksheets = $linkedWorksheets -join ";"
                        LinkedWorkbooks = $linkedWorkbooks -join ";"
                        Connections = $connections -join ";"
                    }
                    $this.results.Add($result) | Out-Null
                }
                $workbook.Close($false)
            }
        }
    }

    [Microsoft.Office.Interop.Excel.Workbook]OpenWorkbook([string]$path, [bool]$readOnly = $false) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false

        if ($readOnly) {
            $workbook = $excel.Workbooks.Open($path, 3, $true) # 3 = read-only mode
        } else {
            $workbook = $excel.Workbooks.Open($path)
        }

        if ($workbook -eq $null) {
            Write-Error "Failed to open workbook: $path"
            return $null
        } else {
            return $workbook
        }
    }

    [string[]]GetLinkedWorksheets([Microsoft.Office.Interop.Excel.Workbook]$workbook) {
        $links = $workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks)
        if ($links) {
            return $links | ForEach-Object { $workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks, $_) } | Select-Object -Unique
        } else {
            return @()
        }
    }

    [string[]]GetLinkedWorkbooks([Microsoft.Office.Interop.Excel.Workbook]$workbook) {
        $links = $workbook.LinkSources([Microsoft.Office.Interop.Excel.XlLinkType]::xlExcelLinks)
        if ($links) {
            return $links | ForEach-Object { Split-Path $_ } | Select-Object -Unique
        } else {
            return @()
        }
    }

    [string[]]GetOleDbConnections([Microsoft.Office.Interop.Excel.Workbook]$workbook) {
        $connections = @()
        foreach ($connection in $workbook.Connections) {
            $connectionStrings = $connection.OLEDBConnection.Connection
            if ($connectionStrings -is [System.Array]) {
                $connectionStrings | ForEach-Object { $connection.Name + ":" + $_ }
            } else {
                $connection.Name + ":" + $connectionStrings
            }
            $connections += $connectionStrings
        }
        return $connections
    }
}
   
# Initialize the ExcelFileAnalyzer class
$analyzer = New-Object ExcelFileAnalyzer

# Analyze the directory and get the results
$analyzer.AnalyzeDirectory("C:\path\to\directory")

# Output the results to a CSV file
$analyzer.results | Export-Csv -Path "C:\path\to\output.csv" -NoTypeInformation
