# Export-WpsPdf.ps1 — WPS 表格工作簿转 PDF
#
# 用法：
#   .\Export-WpsPdf.ps1 -InputPath "D:\in\销售.xlsx" -OutputPath "D:\out\销售.pdf"
#
# ⚠️ 未验证：免费版导出的 PDF 可能带 "WPS Office" 水印，承诺前先在目标机器验证。
param(
    [Parameter(Mandatory)][string]$InputPath,
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
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $wb.ExportAsFixedFormat(0, $OutputPath)  # 0 = xlTypePDF
        Write-Output "OK: exported $OutputPath"
    } finally {
        $wb.Close()
        Release-WpsObject $wb
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
