# WPS 文字常用 recipe

## 查找替换（全部替换）

`Find.Execute` 参数顺序：
`FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace`
（Wrap `1` = wdFindContinue，Replace `2` = wdReplaceAll）

```powershell
$find = $doc.Content.Find
$find.Execute("十五%", $false, $false, $false, $false, $false, $true, 1, $false, "15%", 2) | Out-Null
Release-WpsObject $find
```

## 插入表格

```powershell
$range = $doc.Content.Paragraphs.Add().Range
$table = $doc.Tables.Add($range, 3, 4)   # 3 行 4 列
$table.Cell(1, 1).Range.Text = "姓名"
$table.Cell(1, 2).Range.Text = "部门"
$table.Cell(2, 1).Range.Text = "张三"
$headerRange = $table.Rows.Item(1).Range
$headerRange.Font.Bold = $true
$headerRange.Font.NameFarEast = "黑体"
Release-WpsObject $headerRange
Release-WpsObject $table
Release-WpsObject $range
```

⚠️ 单元格边框 API（`.Borders` / `.BorderAround()`）在 WPS 脚本环境有崩溃报告：
先 smoke-test 再用，优先用表格级边框属性。

## 页眉页脚 + 页码

```powershell
$section = $doc.Sections.Item(1)
$header = $section.Headers.Item(1).Range
$header.Text = "公司机密"
$header.Font.NameFarEast = "宋体"
$header.Font.Size = 9
$footer = $section.Footers.Item(1).Range
$footer.Fields.Add($footer, -1, "PAGE")   # -1 = wdFieldEmpty, "PAGE" = 页码域
Release-WpsObject $footer
Release-WpsObject $header
Release-WpsObject $section
```
