function pxl {
    $sheet = @($args)
    $xl = New-Object -comobject Excel.Application
    $wb = $null
    try {
        $xl.Visible = $false
        $xl.DisplayAlerts = $false

        Get-ChildItem *.xls* | ForEach-Object {
            try {
                $wb = $xl.Workbooks.Open($_.FullName)
                if ($sheet -eq "all") {
                    $wb.Worksheets | Where-Object { $_.Visible -eq $true } | ForEach-Object {
                        $_.PrintOut()
                    }
                }
                elseif ($sheet -eq $null) {
                    $wb.Worksheets.Item(1).PrintOut()
                }
                else {
                    $sheet | ForEach-Object {
                        $ws = $wb.Worksheets.Item($_)
                        if ($ws -ne $null) {
                            $ws.PrintOut()
                        } else {
                            throw "Sheet $_ does not exist in workbook $($_.FullName)"
                        }
                    }
                }
            } catch {
                Write-Error "Error processing workbook $($_.FullName): $($_.Exception.Message)"
                return
            } finally {
                if ($null -ne $wb) {
                    $wb.Close($false)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                }
            }
        }
    } finally {
        $xl.DisplayAlerts = $true
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}