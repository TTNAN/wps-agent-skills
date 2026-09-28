**English** | [简体中文](README_zh-CN.md)

# WPS Agent Skills

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Stars](https://img.shields.io/github/stars/TTNAN/wps-agent-skills)](https://github.com/TTNAN/wps-agent-skills/stargazers)

Teach your AI coding agent (Claude Code, Codex, etc.) to automate **WPS Office** — Writer, Spreadsheets, and Presentation — through COM, using plain PowerShell. **Zero dependencies**: no Python, no Node, no add-ins. If the machine has WPS, the skills work.

Most "office automation" skills target Microsoft Office. In China, most people run WPS — and WPS's COM quirks (Chinese style names with spaces, 32-bit personal edition, `AddChart2` style traps) are exactly what these skills cover, learned from real machines.

## Skills

| Skill | What it teaches the agent |
|---|---|
| [`powershell-windows`](skills/powershell-windows/SKILL.md) | The foundation: PS 5.1 vs 7, BOM/encoding traps, execution policy, admin elevation, paths with spaces, COM basics |
| [`wps-writer`](skills/wps-writer/SKILL.md) | Create/edit `.docx`, styled paragraphs (`标题 1`), find-replace, tables, export PDF |
| [`wps-spreadsheets`](skills/wps-spreadsheets/SKILL.md) | Batch read/write ranges, formulas, formatting, charts, export PDF |
| [`wps-presentation`](skills/wps-presentation/SKILL.md) | Build decks, layouts, text boxes, pictures, export PDF |

## Requirements

- Windows 10/11 with **WPS Office** installed (free personal edition is fine — COM automation is not paywalled)
- Windows PowerShell 5.1+ (ships with Windows)
- An AI coding agent that supports skills: Claude Code (`~/.claude/skills/`), Codex, or similar

> WPS automation needs a **logged-in interactive desktop session**. It won't work from Windows services (Session 0). That's an Office-COM fact of life, not a bug in these skills.

## Install

One-liner (PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

This copies each skill into `~/.claude/skills/<skill-name>/`. Or install manually:

```powershell
git clone https://github.com/TTNAN/wps-agent-skills.git
Copy-Item -Recurse .\wps-agent-skills\skills\wps-writer "$env:USERPROFILE\.claude\skills\wps-writer\"
# repeat for wps-spreadsheets, wps-presentation, powershell-windows
```

Verify WPS COM is reachable:

```powershell
Get-ItemProperty "HKLM:\Software\Classes\KWPS.Application" -ErrorAction SilentlyContinue
```

Using Codex or another agent? Copy the skill folders into that agent's skills directory instead — the `SKILL.md` format is the same.

## Usage

Once installed, just ask in plain language:

> "用 WPS 表格打开 D:\data\销售.xlsx，把 C 列求和写到 C4，表头加粗，导出一份 PDF"

The agent picks up `wps-spreadsheets`, writes a PowerShell script from the skill's recipes, runs it, and reports back. Same for Writer ("把这份报告转成 PDF") and Presentation ("做 5 页 16:9 的分享 PPT").

## Honesty note: verified vs. unverified

Every recipe marked ⚠️ **verify on your machine** in the skills comes from community reports we could not confirm first-hand (e.g. whether the free edition watermarks exported PDFs, whether `Slide.Export` works under WPS). We'd rather flag uncertainty than ship confident-sounding fiction. If you verify one, open a PR — we'll promote it to confirmed.

**Already verified on a real machine** (Windows 11, PowerShell 5.1 64-bit, WPS personal edition, 2026-09-28): all three ProgIDs (`KWPS.Application`, `KET.Application`, `KWPP.Application`); Writer and Spreadsheets run headless (`Visible=$false`) and save `.docx`/`.xlsx` correctly; Presentation builds slides and exports `.pptx`/PDF — but **rejects `Visible=$false` with E_FAIL**, so it must run with its window visible; `wpp.exe` exits cleanly while `wps.exe`/`et.exe` may linger after `Quit()`.

## Project structure

```
wps-agent-skills/
├── README.md / README_zh-CN.md
├── install.ps1                 # one-click installer (UTF-8 with BOM, PS 5.1-safe)
├── LICENSE                     # MIT
└── skills/
    ├── powershell-windows/     # SKILL.md — the Windows/PowerShell baseline
    ├── wps-writer/             # SKILL.md — WPS 文字
    ├── wps-spreadsheets/       # SKILL.md — WPS 表格
    └── wps-presentation/       # SKILL.md — WPS 演示
```

## Contributing

Found a WPS quirk we missed? Verified a ⚠️ item on your machine? PRs welcome — especially real-machine test reports (WPS version + edition + what worked/failed).

## Star history

If this saved you an afternoon, a ⭐ helps others find it.

[![Star History Chart](https://api.star-history.com/svg?repos=TTNAN/wps-agent-skills&type=Date)](https://star-history.com/#TTNAN/wps-agent-skills&Date)

## License

MIT — see [LICENSE](LICENSE).
