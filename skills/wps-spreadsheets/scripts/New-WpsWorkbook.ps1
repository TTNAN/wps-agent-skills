# New-WpsWorkbook.ps1 — 新建 WPS 表格工作簿并批量写入
#
# 用法：
#   .\New-WpsWorkbook.ps1 -OutputPath "$env:TEMP\销售.xlsx" -SheetName "销售数据" `
#       -Headers @("产品", "销量", "金额") `
#       -Rows @(@("产品A", 1500, 300000), @("产品B", 900, 180000))
#
# 批量写：先转成真正的二维数组再一次性赋给 Value2（PowerShell 锯齿数组直接赋值经常翻车）。
param(
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$SheetName = "Sheet1",
    [string[]]$Headers = @(),
    [object[]]$Rows = @()
)
$ErrorActionPreference = "Stop"

function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

$sb = {
    param($et)
    $wb = $et.Workbooks.Add()
    try {
        $ws = $wb.Worksheets.Item(1)   # WPS 集合统一用 .Item(n)
        $ws.Name = $SheetName

        # —— 组装真正的二维数组（0-based，COM 互操作会自动处理）——
        $colCount = $Headers.Count
        foreach ($r in $Rows) { if ($r.Count -gt $colCount) { $colCount = $r.Count } }
        $rowCount = $Rows.Count + $(if ($Headers.Count -gt 0) { 1 } else { 0 })
        if ($rowCount -gt 0 -and $colCount -gt 0) {
            $arr = New-Object 'object[,]' $rowCount, $colCount
            $ri = 0
            if ($Headers.Count -gt 0) {
                for ($ci = 0; $ci -lt $Headers.Count; $ci++) { $arr[$ri, $ci] = $Headers[$ci] }
                $ri++
            }
            foreach ($r in $Rows) {
                for ($ci = 0; $ci -lt $r.Count; $ci++) {
                    $v = $r[$ci]
                    # 以 = 开头的纯文本会被当成公式：加英文单引号强制文本
                    if ($v -is [string] -and $v.StartsWith("=")) { $v = "'" + $v }
                    $arr[$ri, $ci] = $v
                }
                $ri++
            }
            $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rowCount, $colCount)).Value2 = $arr
            # 表头加粗
            if ($Headers.Count -gt 0) {
                $hdr = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $colCount))
                $hdr.Font.Bold = $true
                $hdr.Font.NameFarEast = "黑体"
                Release-WpsObject $hdr
            }
        }
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $wb.SaveAs($OutputPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        Write-Output "OK: saved $OutputPath"
        Release-WpsObject $ws
    } finally {
        $wb.Close()
        Release-WpsObject $wb
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
