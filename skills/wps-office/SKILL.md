---
name: wps-office
description: >
  用户说"用 WPS"但没说是文字/表格/演示时，先用这个路由 skill 判断。
  触发词：WPS、金山办公、金山、wps、et、dps、做文档（未指明类型时）。
  只做路由和跨应用编排，不直接操作 COM。
version: 0.1.0
compatibility: windows + wps-office
---

# wps-office（路由）

很薄的一层：判断该用哪个业务 skill，以及处理跨应用任务。

## 路由规则

| 用户说什么 | 用哪个 skill |
|---|---|
| 文字、报告、合同、docx、wps文件、转PDF（文档） | `wps-writer` |
| 表格、xlsx、et文件、求和、填表、图表 | `wps-spreadsheets` |
| 演示、PPT、pptx、dps文件、幻灯片 | `wps-presentation` |
| 没说清是哪种（"帮我做个文档"） | 先问；或按文件扩展名判断：`.docx/.wps`→writer，`.xlsx/.et`→spreadsheets，`.pptx/.dps`→presentation |
| 目标是 Microsoft Office 不是 WPS | 不用本仓库的任何 skill，用通用 office skill |

## 跨应用任务

表格出图 → 贴进 PPT 这类组合：两个 skill 的 `Invoke-WpsSession.ps1` 各跑各的会话，
不要试图复用同一个 COM 实例。中间文件放 `$env:TEMP`，例如：

1. `wps-spreadsheets`：区域导出图片到 `$env:TEMP\chart.png`
2. `wps-presentation`：`AddPicture("$env:TEMP\chart.png", ...)` 贴进幻灯片

## 用户已经开着 WPS 时

三个业务 skill 的包装器都会检测：附着到用户已开的实例时**跳过 `Quit()`**，
只关闭自己打开的文档/工作簿/文稿。跨应用任务里每个会话独立判断，不互相干扰。
