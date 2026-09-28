---
name: wps-writer
description: >
  在目标机器是 WPS（金山办公）而不是 Microsoft Office 时，用 PowerShell + COM 自动处理 WPS 文字文档。
  触发词：WPS、金山、WPS文字、写报告、转PDF、docx、wps文件、查找替换、页眉页脚。
  不要用 Word.Application —— 在同时装了 Office 和 WPS 的机器上，它可能指向任意一方。
version: 0.1.0
compatibility: windows + wps-office
---

# wps-writer

用 PowerShell + COM 驱动 WPS 文字。零依赖：除了 WPS Office 本身什么都不需要，免费个人版可用。

前置：先读 `powershell-windows` skill（5.1 / BOM / 32-64 位 COM 规则）。

## 何时用 / 何时不用

用：目标机器装的是 WPS（金山办公），要新建 / 改写 / 转 PDF 文字文档。
不用：目标是 Microsoft Word（用通用 word skill）；云文档（COM 只碰本地文件）；Windows 服务 / Session 0（WPS 需要已登录的交互式桌面）。

## 先跑脚本，不要手拼 COM

`scripts/` 里是可直接运行的脚本，输出 `OK:` / `FAIL:` 行。优先调它们，不要从 recipe 现场拼 `New-Object` / `Quit`：

| 脚本 | 用途 | 关键参数 |
|---|---|---|
| `scripts/Invoke-WpsSession.ps1` | 会话包装：连 COM → 跑你的逻辑 → Quit → 倒序释放 → 清残留 | `-Script { param($app) ... }.GetNewClosure()` |
| `scripts/New-WpsDocument.ps1` | 新建文档：标题 + 正文 + 存盘 | `-OutputPath`、`-Title`、`-Paragraphs` |
| `scripts/Export-WpsPdf.ps1` | 文档转 PDF | `-InputPath`、`-OutputPath` |

自定义逻辑时，把 scriptblock 传给 `Invoke-WpsSession.ps1`（注意 `.GetNewClosure()`，否则外层参数传不进去）：

```powershell
$sb = {
    param($app)
    $doc = $app.Documents.Add()
    try {
        $p = $doc.Content.Paragraphs.Add()
        $p.Range.Text = "你好"
        $p.Style = $doc.Styles.Item("标题 1")   # 中文样式名带空格
        $p.Range.Font.NameFarEast = "黑体"
        Release-WpsObject $p          # 倒序释放：先业务对象
        $doc.SaveAs("$env:TEMP\a.docx", 16)
    } finally {
        $doc.Close()
        Release-WpsObject $doc        # …最后释放文档；$app 由包装器释放
    }
}.GetNewClosure()
& "scripts/Invoke-WpsSession.ps1" -Script $sb
```

`Release-WpsObject` 由包装器提供。释放顺序永远是**获取的逆序**（Range → Table → Document → App），最后 `[GC]::Collect()`。

## 参考表

- `references/style-names.md` — 中文样式名（`"标题 1"` 带空格）、`NameFarEast` 双设
- `references/save-formats.md` — SaveAs / ExportAsFixedFormat 常数、只读打开
- `references/recipes.md` — 查找替换、插入表格、页眉页脚 + 页码

## 真机验证过的坑（2026-09-28，Win11 + WPS 个人版）

- ProgID 用 `KWPS.Application`，老版本回退 `WPS.Application`；**永远不要用 `Word.Application` 碰运气**。
- `Visible=$false` + `DisplayAlerts=0`：文字支持后台跑；弹窗（文件已存在 / 兼容性警告）不屏蔽会直接卡死自动化。
- `wps.exe` 在 `Quit()` 后可能残留：包装器用"启动前 PID 快照 → 只杀新增 PID"，绝不 `Stop-Process -Name wps`（会误杀用户已打开的窗口）。
- 用户已经开着 WPS 时：`New-Object` 会附着到他的实例，**跳过 `Quit()`**，只关自己打开的文档。

## 安全硬规则

- 默认只写 `$env:TEMP` 或用户点名的路径；不扫描、不遍历 `C:\Windows`、桌面等用户没提到的位置。
- 重要文档先 `Copy-Item` 备份再改（COM 修改基本不进撤销栈）。
- `Stop-Process` 只杀 PID 快照差集。

## 待验证 ⚠️

- 免费版导出 PDF 是否带 "WPS Office" 水印（社区传闻，未第一手确认）。
- `Borders` / `BorderAround()` 在 WPS 脚本环境有崩溃报告，表格边框先 smoke-test。
- UNC 路径 / 符号链接被报告不可靠，用本地绝对路径。
