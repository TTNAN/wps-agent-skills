# WPS 表格存盘格式常数

| 调用 | 值 | 含义 |
|---|---|---|
| `$wb.SaveAs($path, 51)` | 51 | xlOpenXMLWorkbook → `.xlsx` |
| `$wb.SaveAs($path, 56)` | 56 | xlExcel8 → `.xls` |
| `$wb.ExportAsFixedFormat(0, $path)` | 0 | xlTypePDF → `.pdf` |
| `$et.Workbooks.Open($path, $false, $true)` | — | 第三个参数 `$true` = 只读打开 |

颜色是 BGR 长整型：黄 `65535`、红 `255`、蓝 `16711680`。
