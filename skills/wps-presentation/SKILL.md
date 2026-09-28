---
name: wps-presentation
description: >
  在目标机器是 WPS（金山办公）而不是 Microsoft Office 时，用 PowerShell + COM 自动做 WPS 演示文稿。
  触发词：WPS、金山、WPS演示、做PPT、pptx、dps文件、16:9、转PDF。
  不要用 PowerPoint.Application —— 在同时装了 Office 和 WPS 的机器上，它可能指向任意一方。
  注意：WPS 演示不支持后台运行（Visible=$false 会 E_FAIL），运行时会有窗口弹出来。
version: 0.1.0
compatibility: windows + wps-office
---

# wps-presentation

用 PowerShell + COM 驱动 WPS 演示。零依赖：除了 WPS Office 本身什么都不需要，免费个人版可用。

前置：先读 `powershell-windows` skill（5.1 / BOM / 32-64 位 COM 规则）。

## 何时用 / 何时不用

用：目标机器装的是 WPS（金山办公），要新建 / 改写 `.pptx` / `.dps`、插图、转 PDF。
不用：目标是 Microsoft PowerPoint（用通用 powerpoint skill）；云文档（COM 只碰本地文件）；Windows 服务 / Session 0（WPS 需要已登录的交互式桌面）。

## ⚠️ 先看这条：演示必须亮着窗口跑

`$wpp.Visible = $false` 在 WPS 演示上会直接报 `HRESULT E_FAIL`（2026-09-28 真机验证）。
没有隐藏运行的办法——动手前先告诉用户："WPS 会弹出来一下，跑完自动关。"
`scripts/Invoke-WpsSession.ps1` 已经处理了这点（它永远不碰 `Visible`）。

## 先跑脚本，不要手拼 COM

`scripts/` 里是可直接运行的脚本，输出 `OK:` / `FAIL:` 行。优先调它们：

| 脚本 | 用途 | 关键参数 |
|---|---|---|
| `scripts/Invoke-WpsSession.ps1` | 会话包装：连 COM → 跑你的逻辑 → Quit → 倒序释放 → 清残留（**不设置 Visible**） | `-Script { param($wpp) ... }.GetNewClosure()` |
| `scripts/New-WpsDeck.ps1` | 新建文稿：标题页 + N 个内容页（空白版式+文本框，最稳） | `-OutputPath`、`-Title`、`-Subtitle`、`-Slides` |
| `scripts/Export-WpsPdf.ps1` | 文稿转 PDF（`SaveAs($path, 32)`） | `-InputPath`、`-OutputPath` |

**路径规则**：Agent 的 cwd 通常是用户项目，不是 skill 目录——永远用**绝对路径**调脚本，
先定位 skill 的安装目录（不要用相对路径 `scripts/...`，会找不到文件）：

```powershell
$skillDir = "$env:USERPROFILE\.claude\skills\wps-presentation"   # Codex: .codex\skills\wps-presentation；Cursor: .cursor\skills\wps-presentation
```

自定义逻辑示例：

```powershell
$sb = {
    param($wpp)
    $pres = $wpp.Presentations.Add()
    try {
        $pres.PageSetup.SlideWidth = 960    # 16:9（单位：磅）
        $pres.PageSetup.SlideHeight = 540
        $s = $pres.Slides.Add(1, 1)
        $s.Shapes.Title.TextFrame.TextRange.Text = "你好"
        Release-WpsObject $s               # 倒序释放：先业务对象
        $pres.SaveAs("$env:TEMP\a.pptx")
    } finally {
        $pres.Close()
        Release-WpsObject $pres            # …最后释放文稿；$wpp 由包装器释放
    }
}.GetNewClosure()
$skillDir = "$env:USERPROFILE\.claude\skills\wps-presentation"   # 按实际安装位置改
& "$skillDir\scripts\Invoke-WpsSession.ps1" -Script $sb
```

`Release-WpsObject` 由包装器提供。释放顺序永远是**获取的逆序**（Shape → Slide → Presentation → App），最后 `[GC]::Collect()`。

## 参考表

- `references/layout-ids.md` — 版式 ID、空白版式回退策略
- `references/recipes.md` — 插图、备注页、按页导出 PNG 降级、表格出图→贴进 PPT

## 真机验证过的坑（2026-09-28，Win11 + WPS 个人版）

- ProgID 用 `KWPP.Application`，老版本回退 `Wpp.Application`；**永远不要用 `PowerPoint.Application` 碰运气**。
- **禁止** `$wpp.Visible = $false`（E_FAIL）。想低调可以*试* `$wpp.WindowState = 2` 最小化（⚠️ 未验证）。
- `DisplayAlerts=0` 照设，弹窗（图片路径不存在 / 保存冲突）不屏蔽会卡死。
- 版式 ID 在 WPS 模板下可能和 PowerPoint 不一致：拿不准就用空白版式 + 手写形状。
- `Placeholders.Item(2)` 在某些模板上会抛：try/catch 包住，失败回退到文本框。
- `wpp.exe` 退出干净（同机器上 `wps.exe` / `et.exe` 会残留，`wpp.exe` 不会），但 PID 快照清理照做。
- 用户已经开着 WPS 演示时：跳过 `Quit()`，只关自己打开的文稿。

## 安全硬规则

- 默认只写 `$env:TEMP` 或用户点名的路径；不扫描、不遍历 `C:\Windows`、桌面等用户没提到的位置。
- `AddPicture` 前先 `Test-Path` 确认图片存在。
- `Stop-Process` 只杀 PID 快照差集。

## 待验证 ⚠️

- `Slide.Export` 导出 PNG（对象模型里有，真机未确认）：先试单页，失败则整份 `SaveAs` PDF 降级。
- 免费版导出 PDF 是否带 "WPS Office" 水印（社区传闻，未第一手确认）。
