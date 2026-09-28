# WPS 表格常用 recipe

## 多 sheet

```powershell
$wb = $et.Workbooks.Add()
$ws1 = $wb.Worksheets.Item(1); $ws1.Name = "汇总"
$ws2 = $wb.Worksheets.Add(); $ws2.Name = "明细"   # Add() 追加到末尾
# 按名取：$wb.Worksheets.Item("明细")
Release-WpsObject $ws2
Release-WpsObject $ws1
```

## 按列名写（而不是 A1:C3）

```powershell
$used = $ws.UsedRange
$cols = $used.Columns.Count
$headerRow = @()
for ($c = 1; $c -le $cols; $c++) { $headerRow += $used.Cells.Item(1, $c).Value2 }
$idx = [array]::IndexOf($headerRow, "金额") + 1   # 1-based 列号
if ($idx -gt 0) { $ws.Cells.Item(2, $idx).Value2 = 999 }
Release-WpsObject $used
```

## 冻结窗格 / 打印区域

```powershell
$ws.Range("B2").Select() | Out-Null
$et.ActiveWindow.FreezePanes = $true          # 冻结首行首列
$ws.PageSetup.PrintArea = "`$A`$1:`$E`$50"     # 打印区域（$ 需转义）
```

## 读整表

```powershell
$all = $ws.UsedRange.Value2   # 二维数组，$all[1,1] 对应 A1（1-based）
$rowCount = $ws.UsedRange.Rows.Count
```

## 图表

```powershell
# AddChart2(Style, ChartType, Left, Top, Width, Height)
# Style 必须是 0（默认）：传 -1 会返回 null。51 = xlColumnClustered（柱形图）
$chart = $ws.Shapes.AddChart2(0, 51, 350, 20, 480, 300).Chart
$chart.SetSourceData($ws.Range("A1:C3"))
$chart.HasTitle = $true
$chart.ChartTitle.Text = "季度销量"
$chart.ChartTitle.Font.NameFarEast = "微软雅黑"
Release-WpsObject $chart
```

## 公式与数字格式

```powershell
$ws.Range("C4").Formula = "=SUM(C2:C3)"
$ws.Range("C2:C4").NumberFormat = "#,##0"
```

⚠️ WPS 函数约 400+，不是完整 Excel 集合：`LAMBDA` / `WEBSERVICE` 缺失，
动态数组函数（`FILTER` / `SORT` / `XLOOKUP`）支持不完整，嵌套场景有已知 bug。
复杂公式先在目标机器实测。透视表弱于 Excel，能避则避。
