---
name: wps-presentation
description: Use when automating WPS Presentation (.pptx/.dps) via COM from PowerShell. Covers building slide decks from scratch, title and blank layouts, text boxes, inserting pictures, and exporting to PDF or PNG. Use instead of the generic powerpoint skill whenever the target machine runs WPS Office rather than Microsoft Office.
---

# WPS Presentation Automation (PowerShell + COM)

Drive WPS 演示 (Presentation) through COM from PowerShell. Zero dependencies beyond WPS Office itself. Works with the **free personal edition**.

Requires: Windows 10/11, WPS Office installed, Windows PowerShell 5.1+. Read the `powershell-windows` skill first — encoding, 32/64-bit, and COM release rules all apply.

## 1. Connecting

ProgID is **`KWPP.Application`** (verified; case-insensitive). Fallback for old installs: `Wpp.Application`.

```powershell
function Connect-WpsPresentation {
    foreach ($progId in @("KWPP.Application", "Wpp.Application")) {
        try { return New-Object -ComObject $progId }
        catch { }
    }
    throw "WPS Presentation COM not available. Is WPS Office installed?"
}
```

Do NOT use `PowerPoint.Application` — on dual-install machines it may resolve to MS PowerPoint or WPS unpredictably. Always use the explicit WPS ProgID.

## 2. Session pattern

```powershell
$wpp = $null
try {
    $wpp = Connect-WpsPresentation
    # ⚠️ DO NOT set Visible=$false — WPS Presentation throws HRESULT E_FAIL on it
    # (verified 2026-09-28). Unlike Writer/Spreadsheets, Presentation MUST run
    # with its window visible. Warn the user: a WPS window will pop up.
    $wpp.DisplayAlerts = 0

    $pres = $wpp.Presentations.Add()
    # ... do work ...
    $pres.Close()
}
finally {
    if ($wpp) {
        $wpp.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wpp) | Out-Null
    }
}
```

**Expect a visible window.** There is no supported way to hide WPS Presentation during automation — tell the user upfront ("WPS 会弹出来一下，跑完自动关") instead of letting them think something broke. If you want it less intrusive, you may *try* minimizing (unverified — test on the target machine):

```powershell
# ⚠️ unverified: may not be honored by WPS
$wpp.WindowState = 2   # 2 = minimized (ppWindowMinimized)
```

Session etiquette — never kill the user's own WPS. Snapshot PIDs before you start, and after `Quit()` only terminate processes *you* started that are still lingering:

```powershell
$before = (Get-Process -Name "wpp" -ErrorAction SilentlyContinue).Id
# ... run automation, Quit() in finally ...
Start-Sleep -Seconds 3
$after = Get-Process -Name "wpp" -ErrorAction SilentlyContinue |
    Where-Object { $before -notcontains $_.Id }
if ($after) { $after | Stop-Process -Force }   # only our leftovers, never the user's
```

Verified 2026-09-28: `wpp.exe` exits cleanly after `Quit()` — no leftovers in practice. Still keep the guard; Writer (`wps.exe`) and Spreadsheets (`et.exe`) have been observed lingering on the same machine.

## 3. Recipes

### 3.1 Page setup (16:9)

PowerPoint-family dimensions are in **points** (72 pt = 1 inch):

```powershell
$pres.PageSetup.SlideWidth = 960    # 16:9 widescreen
$pres.PageSetup.SlideHeight = 540
# 4:3 classic would be 960 x 720
```

### 3.2 Title slide

Layout `1` = ppLayoutTitle:

```powershell
$s1 = $pres.Slides.Add(1, 1)
$s1.Shapes.Title.TextFrame.TextRange.Text = "未来已来"
$s1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "2026 年度技术分享"
$s1.Shapes.Title.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
$s1.Shapes.Title.TextFrame.TextRange.Font.Size = 44
```

### 3.3 Content slide (blank layout + text box)

Layout `12` = ppLayoutBlank. Shape type `1` = msoTextBox. `AddTextbox(Type, Left, Top, Width, Height)`:

```powershell
$s2 = $pres.Slides.Add(2, 12)
$tb = $s2.Shapes.AddTextbox(1, 72, 72, 816, 200)
$tb.TextFrame.TextRange.Text = "核心观点一二三……"
$tb.TextFrame.TextRange.Font.Size = 28
$tb.TextFrame.TextRange.Font.Name = "Calibri"
$tb.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"   # always set both for Chinese text
$tb.TextFrame.WordWrap = $true
```

Common layout IDs: `1` title, `2` title+content, `5` title only, `6` blank (varies), `12` blank. If a layout ID misbehaves on WPS, fall back to blank + manual shapes — it always works.

### 3.4 Insert a picture

`AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)`:

```powershell
$s2.Shapes.AddPicture("D:\img\chart.png", $false, $true, 72, 300, 400, 225) | Out-Null
```

Use local absolute paths. Keep aspect ratio in mind — WPS won't auto-fit for you.

### 3.5 Save and export

```powershell
$pres.SaveAs("D:\out\分享.pptx")       # default format is already .pptx
$pres.SaveAs("D:\out\分享.pdf", 32)    # 32 = ppSaveAsPDF
```

⚠️ **Slide thumbnails / PNG export (unverified):** `Slide.Export("D:\out\slide1.png", "PNG", 1920, 1080)` exists in the PowerPoint object model and *should* work in WPS, but this has not been confirmed on a real WPS install. Verify before using:

```powershell
# ⚠️ verify on your machine first
$s1.Export("D:\out\slide1.png", "PNG", 1920, 1080)
```

⚠️ **PDF watermark (unverified):** community reports suggest the free edition may stamp exported PDFs. Verify on your machine before promising clean PDFs.

## 4. Gotchas

| Problem | Cause / fix |
|---|---|
| `wpp.exe` lingers / next run hangs | `Quit()` not in `finally`. Clean up with the PID-snapshot pattern from §2 — never blanket-kill `wpp`. |
| "Cannot create ActiveX component" | Usually 32-bit WPS vs 64-bit host — retry with `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`. (Verified working on 64-bit PowerShell 5.1 + 32-bit WPS personal edition, 2026-09-28, so try 64-bit first.) |
| Chinese text in wrong font | Set `Font.NameFarEast` in addition to `Font.Name`. |
| `Placeholders.Item(2)` throws on some templates | Template placeholder indexing varies — wrap in try/catch and fall back to a textbox. |
| A layout ID produces the wrong layout | WPS template layouts can differ from PowerPoint's — prefer blank layout + explicit shapes for full control. |
| Script hangs with no output | Modal dialog (missing image path, save conflict). `DisplayAlerts = 0` + verify all file paths exist before calling. |

## 5. What NOT to promise

- **Server/headless use:** needs a logged-in interactive desktop session.
- **Cloud presentations:** COM only touches local files.
- **Slide.Export PNG:** unverified — test first, then promise.

## 6. Checklist before shipping

- [ ] Uses `KWPP.Application` explicitly (never `PowerPoint.Application`)?
- [ ] `Visible=$false` is **absent** (it throws E_FAIL on Presentation)?
- [ ] User warned that a WPS window will visibly open and close?
- [ ] `DisplayAlerts=0` set before any work?
- [ ] PID-snapshot cleanup in place (never kill pre-existing `wpp.exe`)?
- [ ] File paths verified to exist before `AddPicture`?
- [ ] `NameFarEast` set wherever Chinese text appears?
- [ ] `Slide.Export` only used after verifying it works on the target machine?
- [ ] Tested on a real machine with WPS installed?
