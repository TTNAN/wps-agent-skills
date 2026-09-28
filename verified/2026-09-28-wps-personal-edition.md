# 真机验证报告 — 2026-09-28

> 以下为用户在自有 Windows 机器上实际运行的结果（关键结果转述，非逐字粘贴）。
> 欢迎用 `verify-wps.ps1` 在你的机器上复现，并把输出按此格式追加到本目录。

## 环境

- 系统：Windows 11
- PowerShell：5.1.26100.8875，64 位
- WPS：个人版（32 位进程）

## verify-wps.ps1 运行结果

| 检查项 | 结果 |
|---|---|
| `KWPS.Application` 注册 | OK |
| `KET.Application` 注册 | OK |
| `KWPP.Application` 注册 | OK |
| writer：新建 + 保存 .docx | OK |
| spreadsheets：新建 + 保存 .xlsx | OK |
| presentation：新建 + 保存 | 待重跑（见下） |
| 残留进程检查 | `wps.exe` / `et.exe` 在 `Quit()` 后可能残留 → 标 WARN（预期内），`wpp.exe` 退出干净 |

### 待用户用新脚本重跑

旧版 `verify-wps.ps1` 对三个应用统一设置 `Visible=$false`，导致 presentation 报 `E_FAIL`
（见附录）。新版已加 `-KeepVisible` 分支——请在同一台机器上重跑并把原文贴到这里：

```powershell
powershell -ExecutionPolicy Bypass -File verify-wps.ps1
```

```text
（粘贴 OK: / WARN: / FAIL: 全文输出）
```

## 结论（已写入 skill）

1. writer / spreadsheets：可用 `Visible=$false` 后台运行。
2. presentation：**禁止**设置 `Visible=$false`，必须以可见窗口运行；跑之前先告诉用户"WPS 会弹出来一下，跑完自动关"。
3. `wps.exe` / `et.exe` 退出后可能残留：用"启动前 PID 快照 → 结束后只杀新增 PID"的方式清理，**禁止**无脑 `Stop-Process -Name wps/et`（会误杀用户原有窗口）。
4. 僵尸判定：启动前存在但**没有可见主窗口**的进程视为上轮崩溃残留，直接清理，不算"用户的实例"。

## 附录：presentation E_FAIL 根因（2026-09-28，diag-wpp.ps1）

| 步骤 | 结果 |
|---|---|
| `New-Object KWPP.Application` | OK |
| `$wpp.Visible = $false` | **FAIL** — `HRESULT E_FAIL` ← 根因 |
| `$wpp.DisplayAlerts = 0` | OK |
| `Presentations.Add()` | OK |
| `Slides.Add(1, 1)`（不可见） | OK |
| 设置标题文本 | OK |
| 保存 .pptx | OK |
| 导出 PDF（format 32） | OK |
| `Quit()` | OK |
| 无新增残留进程 | OK |
