---
name: wps-spreadsheets
description: >
  在目标机器是 WPS（金山办公）而不是 Microsoft Office 时，用 PowerShell + COM 自动处理 WPS 表格。
  触发词：WPS、金山、金山表格、WPS表格、xlsx、et文件、求和、做表、批量填表、图表、转PDF。
  不要用 Excel.Application —— 在同时装了 Office 和 WPS 的机器上，它可能指向任意一方。
version: 0.1.0
compatibility: windows + wps-office
---

# wps-spreadsheets

用 PowerShell + COM 驱动 WPS 表格。零依赖：除了 WPS Office 本身什么都不需要，免费个人版可用。

前置：先读 `powershell-windows` skill（5.1 / BOM / 32-64 位 COM 规则）。

## 何时用 / 何时不用

用：目标机器装的是 WPS（金山办公），要读写 `.xlsx` / `.et`、批量填表、公式、图表、转 PDF。
不用：目标是 Microsoft Excel（用通用 excel skill）；云文档（COM 只碰本地文件）；Windows 服务 / Session 0（WPS 需要已登录的交互式桌面）；复杂透视表（WPS 弱于 Excel，能避则避）。

## 先跑脚本，不要手拼 COM

`scripts/` 里是可直接运行的脚本，输出 `OK:` / `FAIL:` 行。优先调它们：

| 脚本 | 用途 | 关键参数 |
|---|---|---|
| `scripts/Invoke-WpsSession.ps1` | 会话包装：连 COM → 跑你的逻辑 → Quit → 倒序释放 → 清残留（批量写自动关 ScreenUpdating 提速） | `-Script { param($et) ... }.GetNewClosure()` |
| `scripts/New-WpsWorkbook.ps1` | 新建工作簿：表头 + 多行数据一次性写入 | `-OutputPath`、`-SheetName`、`-Headers`、`-Rows` |
| `scripts/Export-WpsPdf.ps1` | 工作簿转 PDF | `-InputPath`、`-OutputPath` |
| `scripts/Export-WpsRangePng.ps1` | ⚠️ 指定区域导出 PNG（真机未验证，先跑一遍看 OK） | `-InputPath`、`-Range`（如 `A1:E10`）、`-OutputPath` |

**路径规则**：Agent 的 cwd 通常是用户项目，不是 skill 目录——永远用**绝对路径**调脚本。
先按 Claude → Codex → Cursor 的顺序探测实际安装位置（不要用相对路径 `scripts/...`，会找不到文件）：

```powershell
$skillDir = @(
    "$env:USERPROFILE\.claude\skills\wps-spreadsheets",
    "$env:USERPROFILE\.codex\skills\wps-spreadsheets",
    "$env:USERPROFILE\.cursor\skills\wps-spreadsheets"
) | Where-Object { Test-Path "$_\scripts\Invoke-WpsSession.ps1" } | Select-Object -First 1
if (-not $skillDir) { throw "wps-spreadsheets skill 未安装：请先跑仓库里的 install.ps1" }
```

自定义逻辑示例：

```powershell
$sb = {
    param($et)
    $wb = $et.Workbooks.Add()
    try {
        $ws = $wb.Worksheets.Item(1)   # WPS 集合统一用 .Item(n)
        $ws.Range("A1").Value2 = "验证"
        Release-WpsObject $ws          # 倒序释放：先业务对象
        $wb.SaveAs("$env:TEMP\a.xlsx", 51)
    } finally {
        $wb.Close()
        Release-WpsObject $wb          # …最后释放工作簿；$et 由包装器释放
    }
}.GetNewClosure()
$skillDir = @(
    "$env:USERPROFILE\.claude\skills\wps-spreadsheets",
    "$env:USERPROFILE\.codex\skills\wps-spreadsheets",
    "$env:USERPROFILE\.cursor\skills\wps-spreadsheets"
) | Where-Object { Test-Path "$_\scripts\Invoke-WpsSession.ps1" } | Select-Object -First 1
if (-not $skillDir) { throw "wps-spreadsheets skill 未安装：请先跑仓库里的 install.ps1" }
& "$skillDir\scripts\Invoke-WpsSession.ps1" -Script $sb
```

`Release-WpsObject` 由包装器提供。释放顺序永远是**获取的逆序**（Range → Worksheet → Workbook → App），最后 `[GC]::Collect()`。

## 参考表

- `references/save-formats.md` — SaveAs / ExportAsFixedFormat 常数、BGR 颜色值
- `references/recipes.md` — 多 sheet、按列名写、冻结窗格、打印区域、图表、公式

## 真机验证过的坑（2026-09-28，Win11 + WPS 个人版）

- ProgID 用 `KET.Application`，老版本回退 `Et.Application`；**永远不要用 `Excel.Application` 碰运气**。
- `Visible=$false` + `DisplayAlerts=0`：表格支持后台跑。
- 集合**统一**用 `.Item(1)`，`Worksheets(1)` 这种直接索引在 WPS 里不可靠。
- 批量写：先拼真正的二维数组（`New-Object 'object[,]'`）再一次性赋给 `Value2`，不要逐格循环（慢一个数量级）；锯齿数组直接赋值经常翻车。
- 以 `=` 开头的纯文本会被当成公式：前面加英文单引号 `"'=不是公式"`。
- 图表用 `AddChart2(0, 51, ...)`：Style 必须是 `0`，传 `-1` 返回 null。
- `Font.Bold` 不要 `-eq $true` 比较（WPS 可能返回 0 / -1 / $true / $false），用真值判断：`if ($cell.Font.Bold) { ... }`。
- `et.exe` 在 `Quit()` 后可能残留：包装器用"启动前 PID 快照 → 只杀新增 PID"，绝不 `Stop-Process -Name et`。
- 用户已经开着 WPS 表格时：跳过 `Quit()`，只关自己打开的工作簿。

## 安全硬规则

- 默认只写 `$env:TEMP` 或用户点名的路径；不扫描、不遍历 `C:\Windows`、桌面等用户没提到的位置。
- `Stop-Process` 只杀 PID 快照差集。

## 待验证 ⚠️

- 免费版导出 PDF 是否带 "WPS Office" 水印（社区传闻，未第一手确认）。
- `Borders` / `BorderAround()` 在 WPS 脚本环境有崩溃报告，先 smoke-test。
- `LAMBDA` / `WEBSERVICE` 缺失；`FILTER` / `SORT` / `XLOOKUP` 等动态数组函数支持不完整——复杂公式先在目标机器实测。
