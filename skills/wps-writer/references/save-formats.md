# WPS 文字存盘格式常数

| 调用 | 值 | 含义 |
|---|---|---|
| `$doc.SaveAs($path, 16)` | 16 | wdFormatDocumentDefault → `.docx` |
| `$doc.SaveAs($path, 0)` | 0 | wdFormatDocument → `.doc` |
| `$doc.ExportAsFixedFormat($path, 17)` | 17 | wdExportFormatPDF → `.pdf` |
| `$wps.Documents.Open($path, $false, $true)` | — | 第三个参数 `$true` = 只读打开 |

打开前备份原则：COM 修改基本不进撤销栈，重要文档先 `Copy-Item` 备份再改。
用本地绝对路径；UNC 路径 / 符号链接被报告不可靠，依赖前先在目标机器验证。
