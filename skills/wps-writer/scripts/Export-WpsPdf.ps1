# Export-WpsPdf.ps1 — WPS 文字文档转 PDF
#
# 用法：
#   .\Export-WpsPdf.ps1 -InputPath "D:\in\报告.docx" -OutputPath "D:\out\报告.pdf"
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
    param($app)
    $doc = $app.Documents.Open($InputPath)
    try {
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $doc.ExportAsFixedFormat($OutputPath, 17)  # 17 = wdExportFormatPDF
        Write-Output "OK: exported $OutputPath"
    } finally {
        $doc.Close()
        Release-WpsObject $doc
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
