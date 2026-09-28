# New-WpsDocument.ps1 — 新建 WPS 文字文档并保存
#
# 用法：
#   .\New-WpsDocument.ps1 -OutputPath "$env:TEMP\报告.docx" -Title "季度报告" -Paragraphs @("第一段……", "第二段……")
param(
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$Title = "",
    [string[]]$Paragraphs = @()
)
$ErrorActionPreference = "Stop"

function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

$sb = {
    param($app)
    $doc = $app.Documents.Add()
    try {
        if ($Title) {
            $p = $doc.Content.Paragraphs.Add()
            $p.Range.Text = $Title
            $p.Style = $doc.Styles.Item("标题 1")   # 中文样式名带空格："标题 1" 不是 "标题1"
            $p.Range.Font.NameFarEast = "黑体"
            Release-WpsObject $p
        }
        foreach ($text in $Paragraphs) {
            $p = $doc.Content.Paragraphs.Add()
            $p.Range.Text = $text
            $p.Style = $doc.Styles.Item("正文")
            $p.Range.Font.NameFarEast = "宋体"
            $p.Range.Font.Size = 12
            Release-WpsObject $p
        }
        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $doc.SaveAs($OutputPath, 16)   # 16 = wdFormatDocumentDefault (.docx)
        Write-Output "OK: saved $OutputPath"
    } finally {
        $doc.Close()
        Release-WpsObject $doc
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
