---
name: wps-writer
description: Use when automating WPS Writer documents via COM from PowerShell. Covers creating and editing .docx files, writing paragraphs with styles, find-and-replace, inserting tables, saving, and exporting to PDF. Use instead of the generic word/office skill whenever the target machine runs WPS Office rather than Microsoft Office.
---

# WPS Writer Automation (PowerShell + COM)

Teach the agent to drive WPS 文字 (Writer) through COM. Zero dependencies beyond WPS Office itself — no Python, no Node, no add-ins. Works with the **free personal edition**.

Requires: Windows 10/11, WPS Office installed, Windows PowerShell 5.1+. Also read the `powershell-windows` skill first — encoding, 32/64-bit, and COM release rules all apply here.

## 1. Connecting

ProgID is **`KWPS.Application`** (verified across multiple sources; case-insensitive). Fallback for old installs: `WPS.Application`.

```powershell
function Connect-WpsWriter {
    foreach ($progId in @("KWPS.Application", "WPS.Application")) {
        try { return New-Object -ComObject $progId }
        catch { }
    }
    throw "WPS Writer COM not available. Is WPS Office installed? (ProgIDs tried: KWPS.Application, WPS.Application)"
}
```

**Do NOT use `Word.Application` and hope it lands on WPS.** On machines with both MS Office and WPS installed, `Word.Application` may resolve to either one depending on registry view and install order. Always use the explicit WPS ProgID.

Check registration without launching WPS:

```powershell
Get-ItemProperty "HKLM:\Software\Classes\KWPS.Application" -ErrorAction SilentlyContinue
```

## 2. Session pattern (use this shell for every task)

```powershell
$wps = $null
try {
    $wps = Connect-WpsWriter
    $wps.Visible = $false      # run headless — WPS respects this
    $wps.DisplayAlerts = 0     # wdAlertsNone: suppress ALL modal dialogs, or automation hangs

    $doc = $wps.Documents.Add()
    # ... do work (see recipes) ...
    $doc.Close()
}
finally {
    if ($wps) {
        $wps.Quit()   # MANDATORY — otherwise wps.exe lingers and the next run hangs on Documents.Open
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wps) | Out-Null
    }
}
```

**Session etiquette:** WPS is single-instance. If the user already has WPS open, `New-Object` may attach to their instance — calling `Quit()` would kill *their* window. Before automating, check:

```powershell
$alreadyRunning = Get-Process -Name "wps" -ErrorAction SilentlyContinue
```

If WPS was already running: skip `Quit()`, only close the documents you opened. If you launched it: `Quit()` in `finally`.

## 3. Recipes

### 3.1 Create a document with styled paragraphs

```powershell
$doc = $wps.Documents.Add()

$p1 = $doc.Content.Paragraphs.Add()
$p1.Range.Text = "2026 年第三季度工作报告"
$p1.Style = $doc.Styles.Item("标题 1")   # NOTE: Chinese style names contain a SPACE: "标题 1", not "标题1"
$p1.Range.Font.Name = "Calibri"
$p1.Range.Font.NameFarEast = "黑体"       # Chinese text needs NameFarEast, or it falls back to the default font

$p2 = $doc.Content.Paragraphs.Add()
$p2.Range.Text = "本季度核心指标同比增长 15%……"
$p2.Style = $doc.Styles.Item("正文")
$p2.Range.Font.NameFarEast = "宋体"
$p2.Range.Font.Size = 12
```

### 3.2 Open an existing document

```powershell
$doc = $wps.Documents.Open("D:\in\合同.docx")
# Read-only: $wps.Documents.Open("D:\in\合同.docx", $false, $true)
```

**Safety rule:** COM edits don't reliably enter the undo stack. For important documents, save a backup copy before modifying.

Use local absolute paths only. ⚠️ UNC paths and symlinks have been reported as unreliable — verify on your machine before relying on them.

### 3.3 Find and replace (replace-all)

Parameter order for `Find.Execute`:
`FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace`
(Wrap `1` = wdFindContinue, Replace `2` = wdReplaceAll)

```powershell
$find = $doc.Content.Find
$find.Execute("十五%", $false, $false, $false, $false, $false, $true, 1, $false, "15%", 2) | Out-Null
```

### 3.4 Insert a table

```powershell
$range = $doc.Content.Paragraphs.Add().Range
$table = $doc.Tables.Add($range, 3, 4)   # 3 rows, 4 columns
$table.Cell(1, 1).Range.Text = "姓名"
$table.Cell(1, 2).Range.Text = "部门"
$table.Cell(2, 1).Range.Text = "张三"
# Style the header row
$headerRange = $table.Rows.Item(1).Range
$headerRange.Font.Bold = $true
$headerRange.Font.NameFarEast = "黑体"
```

⚠️ Cell border APIs (`.Borders` / `.BorderAround()`) have crash reports in WPS scripting environments. Prefer table-level border properties, and smoke-test borders on your machine before shipping a script that sets them.

### 3.5 Save and export

```powershell
$doc.SaveAs("D:\out\报告.docx", 16)                  # 16 = wdFormatDocumentDefault (.docx)
$doc.ExportAsFixedFormat("D:\out\报告.pdf", 17)      # 17 = wdExportFormatPDF
```

⚠️ **PDF watermark (unverified):** community reports say the free edition exports PDF without page limits but *may* stamp a "WPS Office" watermark. Verify on your machine before promising watermark-free PDFs.

## 4. Gotchas

| Problem | Cause / fix |
|---|---|
| `wps.exe` lingers after script ends | `Quit()` wasn't reached — always put it in `finally`. Verified 2026-09-28: `wps.exe` can linger even after `Quit()`; clean up with the PID-snapshot pattern (kill only PIDs you started, never the user's pre-existing instance): `$before=(Get-Process wps -EA SilentlyContinue).Id; ... Start-Sleep 3; Get-Process wps -EA SilentlyContinue | Where-Object {$before -notcontains $_.Id} | Stop-Process -Force` |
| Script hangs with no output | A modal dialog is open (file-exists prompt, compatibility warning). You forgot `DisplayAlerts = 0`, or WPS showed a login/ad popup. Run with `Visible=$false`; if login popups persist, they must be dismissed in WPS settings once by the user. |
| "Cannot create ActiveX component" / ProgID not found | 32-bit WPS vs 64-bit PowerShell. Retry from 32-bit PowerShell: `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`. |
| Chinese text renders in the wrong font | You set `Font.Name` but not `Font.NameFarEast`. Always set both. |
| Style assignment silently does nothing | You used `"标题1"`. WPS Chinese style names have a space: `"标题 1"`, `"标题 2"`, `"正文"`. |
| Next run hangs on `Documents.Open` | A zombie `wps.exe` from a crashed run is holding the file. Kill it first. |
| Works on your machine, fails on user's | Their edition (教育版/政务版) may register different ProgIDs — that's why the connect function tries the fallback. |

## 5. What NOT to promise

- **Server/headless use:** WPS needs a logged-in interactive desktop session. Don't claim it works from Windows services (Session 0) or headless scheduled tasks — treat "run only when user is logged on" as the supported setup.
- **Cloud documents:** COM only touches local files. Sync WPS cloud docs to local first.
- **`KPDF.Application`:** do not touch it via COM — it pops a blocking dialog. For PDF merge/split, use a pure-Python library instead.

## 6. Checklist before shipping

- [ ] Uses `KWPS.Application` explicitly (never `Word.Application`)?
- [ ] `Visible=$false` + `DisplayAlerts=0` set before any document work?
- [ ] `Quit()` in `finally`, with already-running-instance check?
- [ ] Chinese styles use `"标题 1"` (with space)?
- [ ] `NameFarEast` set wherever Chinese text appears?
- [ ] Tested on a real machine with WPS installed?
