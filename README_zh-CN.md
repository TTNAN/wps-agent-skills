# WPS Agent Skills

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Stars](https://img.shields.io/github/stars/TTNAN/wps-agent-skills)](https://github.com/TTNAN/wps-agent-skills/stargazers)
[![English](https://img.shields.io/badge/docs-English-blue)](README.md)

教你的 AI 编程助手（Claude Code、Codex 等）用纯 PowerShell 通过 COM 自动化 **WPS Office**——文字、表格、演示全覆盖。**零依赖**：不要 Python、不要 Node、不要装加载项，机器上有 WPS 就能跑。

市面上的"办公自动化" skill 基本都是给 Microsoft Office 写的。但在国内，大家用的其实是 WPS——而 WPS 的 COM 坑（带空格的中文样式名、32 位个人版、`AddChart2` 的样式陷阱）正是这套 skill 要解决的，每一条都来自真机实测。

## Skill 一览

| Skill | 教会 Agent 什么 |
|---|---|
| [`powershell-windows`](skills/powershell-windows/SKILL.md) | 地基：PS 5.1 与 7 的区别、BOM/编码坑、执行策略、提权、带空格路径、COM 基础 |
| [`wps-writer`](skills/wps-writer/SKILL.md) | 新建/编辑 `.docx`、带样式的段落（`标题 1`）、查找替换、插入表格、导出 PDF |
| [`wps-spreadsheets`](skills/wps-spreadsheets/SKILL.md) | 批量读写单元格、公式、数字格式、图表、导出 PDF |
| [`wps-presentation`](skills/wps-presentation/SKILL.md) | 从零搭幻灯片、版式、文本框、插图、导出 PDF |

## 环境要求

- Windows 10/11 + 安装了 **WPS Office**（免费个人版就行，COM 自动化不收费）
- Windows PowerShell 5.1+（系统自带）
- 支持 skill 的 AI 编程助手：Claude Code（`~/.claude/skills/`）、Codex 等

> WPS 自动化需要**已登录的交互式桌面会话**，Windows 服务（Session 0）里跑不起来。这是 Office COM 的先天限制，不是 skill 的 bug。

## 安装

一行命令（PowerShell）：

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

会自动把每个 skill 拷到 `~/.claude/skills/<skill-name>/`。也可以手动安装：

```powershell
git clone https://github.com/TTNAN/wps-agent-skills.git
Copy-Item -Recurse .\wps-agent-skills\skills\wps-writer "$env:USERPROFILE\.claude\skills\wps-writer\"
# wps-spreadsheets、wps-presentation、powershell-windows 同理
```

验证 WPS COM 是否可用：

```powershell
Get-ItemProperty "HKLM:\Software\Classes\KWPS.Application" -ErrorAction SilentlyContinue
```

用 Codex 或其他助手？把 skill 文件夹拷到对应助手的 skills 目录即可，`SKILL.md` 格式是通用的。

## 用法

装好之后直接说人话就行：

> "用 WPS 表格打开 D:\data\销售.xlsx，把 C 列求和写到 C4，表头加粗，再导出一份 PDF"

Agent 会自动调用 `wps-spreadsheets`，照着 skill 里的 recipe 写 PowerShell 脚本、运行、汇报结果。文字（"把这份报告转成 PDF"）和演示（"做 5 页 16:9 的分享 PPT"）同理。

## 诚实声明：已验证 vs 待验证

Skill 里标了 ⚠️"请在你的机器上验证"的条目，都来自我们未能第一手确认的社区经验（比如免费版导出 PDF 是否带水印、WPS 下 `Slide.Export` 是否可用）。我们宁可标出不确定，也不编造"看起来很确定"的内容。如果你验证了某一条，欢迎提 PR，我们会把它升级为已确认。

**已在真机验证通过**（Windows 11、PowerShell 5.1 64 位、WPS 个人版，2026-09-28）：三个 ProgID（`KWPS.Application`、`KET.Application`、`KWPP.Application`）全部可用；文字和表格支持后台运行（`Visible=$false`）并正常存盘 `.docx`/`.xlsx`；演示可以建幻灯片、导出 `.pptx`/PDF——但**设置 `Visible=$false` 会直接报 E_FAIL**，必须亮着窗口跑；`wpp.exe` 退出干净，而 `wps.exe`/`et.exe` 在 `Quit()` 后可能残留。

## 目录结构

```
wps-agent-skills/
├── README.md / README_zh-CN.md
├── install.ps1                 # 一键安装脚本（UTF-8 带 BOM，兼容 PS 5.1）
├── LICENSE                     # MIT
└── skills/
    ├── powershell-windows/     # SKILL.md — Windows/PowerShell 地基
    ├── wps-writer/             # SKILL.md — WPS 文字
    ├── wps-spreadsheets/       # SKILL.md — WPS 表格
    └── wps-presentation/       # SKILL.md — WPS 演示
```

## 贡献

发现了我们没写到的 WPS 坑？在你的机器上验证了某条 ⚠️？欢迎提 PR——特别欢迎真机测试报告（WPS 版本 + 版本类型 + 成功/失败情况）。

## Star 趋势

如果这套 skill 帮你省了一个下午，点个 ⭐ 能让更多人找到它。

[![Star History Chart](https://api.star-history.com/svg?repos=TTNAN/wps-agent-skills&type=Date)](https://star-history.com/#TTNAN/wps-agent-skills&Date)

## 许可证

MIT — 见 [LICENSE](LICENSE)。
