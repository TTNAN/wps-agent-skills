# 真机验证报告 — 2026-09-28

> 以下为用户在自有 Windows 机器上实际运行的结果。欢迎用 `verify-wps.ps1`
> 在你的机器上复现，并把输出按此格式追加到本目录。

## 环境

- 系统：Windows 11
- PowerShell：5.1.26100.8875，64 位
- WPS：个人版（32 位进程）

## verify-wps.ps1 运行结果（原文，2026-09-28 晚，修复后重跑）

```text
OK: powershell-version 5.1.26100.8875
OK: powershell-bitness 64-bit
OK: progid-registered writer -> KWPS.Application
OK: progid-registered presentation -> KWPP.Application
OK: progid-registered spreadsheets -> KET.Application
OK: app-writer create+save works
OK: app-spreadsheets create+save works
OK: app-presentation create+save works
OK: no-leftover-process wps
OK: no-leftover-process et
OK: no-leftover-process wpp
DONE. 临时文件在 D:\DSH-TE~2\wps-verify
```

presentation 一项用的是新脚本的 `-KeepVisible` 分支（可见窗口跑，不再设置 `Visible=$false`），**绿了**。

| 检查项 | 结果 |
|---|---|
| `KWPS.Application` / `KET.Application` / `KWPP.Application` 注册 | OK |
| writer：新建 + 保存 .docx | OK |
| spreadsheets：新建 + 保存 .xlsx | OK |
| presentation：新建 + 保存（可见窗口） | OK |
| 残留进程检查 | OK（本次无残留；`wpp.exe` 退出干净） |

## 插曲：当天 COM 注册损坏又修复（已解决）

同一天 17:50 重跑时曾出现：三个 `New-Object` 全部报 `80040154`，
且错误里的 CLSID 是 `{00000000-0000-0000-0000-000000000000}`（空值）——
ProgID 的注册表键还在，但它指向的 CLSID 链断了。早上同一台机器还一切正常，
头号嫌疑是 WPS 个人版当天静默自动更新把注册表写坏了。

修复：用户直接用 WPS 官方最新安装包**更新**了一遍 WPS，之后重跑 verify-wps.ps1，全部回到 OK。
（备选：WPS 自带"配置工具"→ 高级 → 重置修复，也可能修好。）

教训已写入脚本：`verify-wps.ps1` 的 ProgID 检查不再只看"键存在"，
而是跟完 **ProgID → CLSID → LocalServer32 → exe 文件存在** 整条链，
链断会直接报 `FAIL: progid-registered …（CLSID 为空：COM 注册已损坏，需修复/重装 WPS）`。
另：用户环境是 64 位 PowerShell + 32 位 WPS，32 位 COM 经常只注册在
`HKLM:\Software\Wow6432Node\Classes` 下，所以检查的查找根已把 Wow6432Node 加进去，
避免在"只写了 Wow6432Node"的机器上误报"CLSID 为空"。

## 结论（已写入 skill）

1. writer / spreadsheets：可用 `Visible=$false` 后台运行。
2. presentation：**禁止**设置 `Visible=$false`，必须以可见窗口运行；跑之前先告诉用户"WPS 会弹出来一下，跑完自动关"。
3. `wps.exe` / `et.exe` 退出后可能残留：用"启动前 PID 快照 → 结束后只杀新增 PID"的方式清理，**禁止**无脑 `Stop-Process -Name wps/et`（会误杀用户原有窗口）。
4. 僵尸判定：无可见主窗口的预存进程 ≠ 一定是僵尸（可能是托盘驻留/云同步/还没退完），开局不杀；只有 `New-Object` 全部失败时才清理它们并重试。
5. WPS 个人版可能静默自动更新并破坏 COM 注册；如果某天脚本突然全挂 80040154，先跑 `verify-wps.ps1` 看链，再修 WPS。

## 附录：presentation E_FAIL 根因（2026-09-28 早，diag-wpp.ps1）

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
