#TMC 2/10/23
#Rev 13 3/10/23
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Collections.Generic

class ExcelFileAnalyzer {
    [System.Collections.Generic.List[PSObject]]$results

    ExcelFileAnalyzer() {
        $this.results = [System.Collections.Generic.List[PSObject]]::new()
    }

    [void]AnalyzeDirectory([string]$path) {
        $files = Get-ChildItem -Path $path -Filter *.xls,*.xlsx -Recurse
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
                    $this.results.Add($result)
                }
                $workbook.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($file) | Out-Null
        }
    }

   [Microsoft.Office.Interop.Excel.Workbook]OpenWorkbook([string]$path, [bool]$readOnly = $false) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $options = @{
        FilePath = $path
        ReadOnly = $readOnly
        UpdateLinks = 0
    }
    $workbook = $excel.Workbooks.Open($options)

    if(!$workbook) {
        Write-Error "Failed to open workbook: $path"
    } elseif ($workbook.Password) {
        Write-Warning "Skipping password-protected workbook: $path"
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        $workbook = $null
    }
    return $workbook
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
                $connectionStrings | ForEach-Object { $connections += $connection.Name + ":" + $_ }
            } else {
                $connections += $connection.Name + ":" + $connectionStrings
            }
        }
        return $connections
    }
    }
    # Example usage:
    $analyzer = New-Object ExcelFileAnalyzer
    $analyzer.AnalyzeDirectory("C:\ExcelFiles")
    $analyzer.results | Export-Csv "C:\ExcelFiles\Results.csv" -NoTypeInformation
    

