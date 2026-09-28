# Export-WpsRangePng.ps1 — 把工作表指定区域导出为 PNG
#
# ⚠️ 未验证：基于 Excel 通用对象模型的写法（Range.CopyPicture → Chart.Paste → Chart.Export），
#    在 WPS 真机上尚未确认可用。先在目标机器跑一遍，看到 OK: 才算数；失败就改用 Export-WpsPdf.ps1 整本转 PDF。
#
# 用法：
#   .\Export-WpsRangePng.ps1 -InputPath "D:\in\销售.xlsx" -Range "A1:E10" -OutputPath "$env:TEMP\chart.png"
param(
    [Parameter(Mandatory)][string]$InputPath,
    [string]$Range = "A1:E10",
    [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference = "Stop"
if (-not (Test-Path $InputPath)) { Write-Output "FAIL: input-not-found $InputPath"; exit 1 }

function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

$sb = {
    param($et)
    $wb = $et.Workbooks.Open($InputPath)
    try {
        $ws = $wb.Worksheets.Item(1)
        # 经典套路：区域复制为图片 → 粘贴到临时图表 → 图表导出 PNG → 删掉临时图表
        $ws.Range($Range).CopyPicture(1, 2)   # 1=xlScreen, 2=xlPicture
        $chartObj = $ws.ChartObjects().Add(0, 0, 800, 600)
        $chartObj.Chart.Paste()
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $okExport = $chartObj.Chart.Export($OutputPath, "PNG")
        $chartObj.Delete()
        if (-not $okExport -or -not (Test-Path $OutputPath)) { throw "Chart.Export failed on this machine" }
        Write-Output "OK: exported $OutputPath"
        Release-WpsObject $chartObj
        Release-WpsObject $ws
    } finally {
        $wb.Close($false)   # 丢弃临时图表，不保存
        Release-WpsObject $wb
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
