# New-WpsDeck.ps1 — 新建 WPS 演示文稿
#
# 用法：
#   .\New-WpsDeck.ps1 -OutputPath "$env:TEMP\分享.pptx" -Title "未来已来" -Subtitle "2026 年度技术分享" `
#       -Slides @("核心观点一", "核心观点二", "谢谢")
#
# 版式策略：标题页用版式 1，其余用空白版式(12)+文本框——WPS 模板版式编号可能和 PowerPoint 不一致，
# 空白+手写形状永远可用，是最稳的回退。
param(
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$Title = "",
    [string]$Subtitle = "",
    [string[]]$Slides = @()
)
$ErrorActionPreference = "Stop"

function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

$sb = {
    param($wpp)
    $pres = $wpp.Presentations.Add()
    try {
        $pres.PageSetup.SlideWidth = 960     # 16:9（单位：磅，72 磅 = 1 英寸）
        $pres.PageSetup.SlideHeight = 540

        # —— 标题页（版式 1 = ppLayoutTitle）——
        $s1 = $pres.Slides.Add(1, 1)
        if ($Title) { $s1.Shapes.Title.TextFrame.TextRange.Text = $Title }
        if ($Subtitle) {
            try { $s1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $Subtitle }
            catch { }   # 模板占位符编号可能不一致，失败就跳过
        }
        if ($Title) {
            $tf = $s1.Shapes.Title.TextFrame.TextRange
            $tf.Font.NameFarEast = "微软雅黑"
            $tf.Font.Size = 44
        }
        Release-WpsObject $s1

        # —— 内容页：空白版式 + 文本框（最稳）——
        $i = 2
        foreach ($text in $Slides) {
            $s = $pres.Slides.Add($i, 12)   # 12 = ppLayoutBlank
            $tb = $s.Shapes.AddTextbox(1, 72, 72, 816, 200)   # 1 = msoTextBox
            $tb.TextFrame.TextRange.Text = $text
            $tb.TextFrame.TextRange.Font.Size = 28
            $tb.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
            $tb.TextFrame.WordWrap = $true
            Release-WpsObject $tb
            Release-WpsObject $s
            $i++
        }

        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        $pres.SaveAs($OutputPath)   # 默认就是 .pptx
        Write-Output "OK: saved $OutputPath"
    } finally {
        $pres.Close()
        Release-WpsObject $pres
    }
}.GetNewClosure()

& "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
