# Export-WpsPdf.ps1 — WPS 演示文稿转 PDF
#
# 用法：
#   .\Export-WpsPdf.ps1 -InputPath "D:\in\分享.pptx" -OutputPath "D:\out\分享.pdf"
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
    param($wpp)
    $pres = $wpp.Presentations.Open($InputPath)
    try {
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $pres.SaveAs($OutputPath, 32)  # 32 = ppSaveAsPDF
        Write-Output "OK: exported $OutputPath"
    } finally {
        $pres.Close()
        Release-WpsObject $pres
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
