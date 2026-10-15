# Opens the workbook in Excel, forces a full recalculation, dumps every sheet's values to JSON for independent checking
# (src/verify.py compares them with the Python mirror), and optionally saves the recalculated file so cached values are stored.
#
#   powershell -File src\recalc_and_dump.ps1 -Path <xlsx> -Dump <json> [-SaveAs <xlsx>] [-Scenario Base|Upside|Downside] [-Pdf <pdf>]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Dump,
    [string]$SaveAs = "",
    [string]$Scenario = "",
    [string]$Pdf = ""
)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    if ($Scenario -ne "") { $wb.Worksheets.Item("Dashboard").Range("O2").Value2 = $Scenario }
    $xl.CalculateFull()
    $out = @{}
    foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        $rows = $used.Rows.Count; $cols = $used.Columns.Count
        $vals = $used.Value2
        $grid = New-Object System.Collections.ArrayList
        for ($r = 1; $r -le $rows; $r++) {
            $line = New-Object System.Collections.ArrayList
            for ($c = 1; $c -le $cols; $c++) {
                $v = $vals[$r, $c]
                if ($v -is [int] -and $v -lt 0 -and $v -le -2146826246) { $v = "#ERR" }   # Excel error codes arrive as negative ints
                [void]$line.Add($v)
            }
            [void]$grid.Add($line)
        }
        $out[$ws.Name] = @{ first_row = $used.Row; first_col = $used.Column; values = $grid }
    }
    ($out | ConvertTo-Json -Depth 6 -Compress) | Set-Content -Path $Dump -Encoding UTF8
    if ($Pdf -ne "") {
        $dash = $wb.Worksheets.Item("Dashboard")
        $dash.ExportAsFixedFormat(0, $Pdf)
    }
    if ($SaveAs -ne "") {
        $wb.Worksheets.Item("Cover").Activate()
        $wb.SaveAs($SaveAs, 51)
    }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
}
"ok"
