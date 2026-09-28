---
name: powershell-windows
description: Use when writing, running, or debugging PowerShell scripts on Windows. Covers Windows PowerShell 5.1 vs PowerShell 7+ differences, file encoding and BOM pitfalls, execution policy, admin elevation, paths with spaces or non-ASCII characters, COM automation basics (New-Object -ComObject, 32/64-bit), and handling Chinese text output correctly.
---

# PowerShell on Windows — Baseline

Rules for writing PowerShell that actually runs on real Windows machines — not just in theory. Assume **Windows PowerShell 5.1** unless the user explicitly says they have PowerShell 7+. Everything below was learned from real failures.

## 1. Know your host

```powershell
$PSVersionTable.PSVersion   # 5.1.x = Windows PowerShell (ships with Windows), 7.x = PowerShell Core (separate install)
```

- **Never assume PowerShell 7.** On a random user's machine, `powershell.exe` = 5.1. Write 5.1-compatible code.
- 5.1 does **not** support `if` as an expression: `$x = if ($a) {1} else {2}` is a syntax error. Use `$x = $(if ($a) {1} else {2})` or a plain if/else block.
- 5.1 does **not** support ternary `$a ? $b : $c`, null-coalescing `??`, or `?.` chains in all positions the way 7 does. Avoid them.

## 2. Encoding: the #1 silent killer

Windows PowerShell 5.1 decodes `.ps1` files **without BOM as the system ANSI codepage** (GBK on Chinese Windows). UTF-8 Chinese characters become mojibake and can break syntax.

Rules:

| File type | Rule |
|---|---|
| `.ps1` containing non-ASCII (Chinese) | **Save with UTF-8 BOM.** 5.1 needs the BOM to detect UTF-8. |
| `.json` read by your own code | **Must be BOM-less.** Many parsers choke on BOM. In Python, open with `encoding="utf-8-sig"` to tolerate both. |
| Writing UTF-8 without BOM from 5.1 | `Out-File -Encoding utf8` writes **with** BOM in 5.1. For BOM-less: `[IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))` |
| Chinese output garbled in console | `[Console]::OutputEncoding = [Text.Encoding]::UTF8` at the top of the script. |

## 3. Execution policy

Scripts are blocked by default (`Restricted`). For a setup/install script, either:

```powershell
# One-time, current user only (no admin needed):
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
```

or instruct the user to run with `powershell -ExecutionPolicy Bypass -File script.ps1`.

## 4. Admin elevation

Check, don't assume:

```powershell
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # Re-launch self elevated
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
```

## 5. Paths with spaces and non-ASCII usernames

- Always quote paths: `"C:\Program Files\..."`. When building argument strings for native exes, quote each path separately.
- **icacls vs usernames with spaces** (e.g. `Haonan Tong`): `icacls "C:\path" /grant "Haonan Tong:(OI)(CI)F"` fails to parse. Fix: resolve to SID first —
  ```powershell
  $sid = (New-Object Security.Principal.NTAccount($env:USERNAME)).Translate(
      [Security.Principal.SecurityIdentifier]).Value
  icacls "C:\path" /grant "${sid}:(OI)(CI)F"
  ```
- Prefer `$env:USERPROFILE`, `$env:APPDATA`, `$env:TEMP` over hardcoded `C:\Users\...`.

## 6. COM automation basics

```powershell
$obj = New-Object -ComObject "Some.ProgID"
try {
    # ... work ...
} finally {
    # ALWAYS release, otherwise the host process lingers
    [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
```

- **32-bit vs 64-bit**: many desktop apps (WPS Office personal edition, older Office) register COM only in the 32-bit registry view. If 64-bit PowerShell says "cannot find ProgID", retry with 32-bit PowerShell: `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`.
- Set `$app.DisplayAlerts = 0` (or `$false`) before doing anything — modal dialogs hang headless automation forever.
- COM method calls with optional parameters: use `[Type]::Missing` or `[System.Reflection.Missing]::Value` for skipped args.

## 7. Error handling template

```powershell
$ErrorActionPreference = "Stop"   # turn non-terminating errors into terminating ones
try {
    # ... main logic ...
} catch {
    Write-Error "FAILED: $($_.Exception.Message)"
    exit 1
}
```

For scripts an agent runs unattended: print machine-readable status lines (`OK: ...` / `FAIL: ...`) so the agent can parse results without reading Chinese console output.

## 8. Quick checklist before shipping a script

- [ ] Runs on 5.1 (no 7-only syntax)?
- [ ] `.ps1` saved **with** UTF-8 BOM if it contains Chinese?
- [ ] JSON/data files BOM-less?
- [ ] No hardcoded `C:\Users\...`?
- [ ] COM objects released in `finally`?
- [ ] `DisplayAlerts` silenced for automation?
- [ ] Tested on a real machine, not just reviewed?
