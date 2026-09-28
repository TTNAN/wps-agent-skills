# WPS 演示常用 recipe

## 插入图片

`AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)`：

```powershell
$s.Shapes.AddPicture("D:\img\chart.png", $false, $true, 72, 300, 400, 225) | Out-Null
```

用本地绝对路径；WPS 不会自动保持宽高比，自己算好。

## 备注页

```powershell
$notes = $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange
$notes.Text = "讲到这里停顿，抛问题给观众"
Release-WpsObject $notes
```

## 按页导出 PNG（⚠️ 未验证）降级方案

`Slide.Export` 在 PowerPoint 对象模型里存在，WPS 是否可用**未在真机确认**。
先试单页，失败则整份转 PDF：

```powershell
try {
    $s.Export("$env:TEMP\slide1.png", "PNG", 1920, 1080)
    Write-Output "OK: slide-export works on this machine"
} catch {
    # 降级：整份 SaveAs PDF
    $pres.SaveAs("$env:TEMP\deck.pdf", 32)  # 32 = ppSaveAsPDF
    Write-Output "WARN: Slide.Export unsupported, fell back to full-deck PDF"
}
```

## 跨应用组合：表格出图 → 贴进 PPT

办公自动化最常见的组合拳。表格端把区域导出成图片（`Range.CopyPicture` 后另存，或先转 PDF 再转图），
演示端用上面的 `AddPicture` 贴进来。两个 skill 的 `Invoke-WpsSession.ps1` 各管各的会话，
不要试图复用同一个 COM 实例。
