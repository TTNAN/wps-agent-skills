---
name: wps-spreadsheets
description: Use when automating WPS Spreadsheets (.xlsx/.et) via COM from PowerShell. Covers batch reading and writing cell ranges, formulas, number formats, styling, charts, and exporting to PDF. Use instead of the generic excel skill whenever the target machine runs WPS Office rather than Microsoft Office.
---

# WPS Spreadsheets Automation (PowerShell + COM)

Drive WPS 表格 (Spreadsheets) through COM from PowerShell. Zero dependencies beyond WPS Office itself. Works with the **free personal edition**.

Requires: Windows 10/11, WPS Office installed, Windows PowerShell 5.1+. Read the `powershell-windows` skill first — encoding, 32/64-bit, and COM release rules all apply.

## 1. Connecting

ProgID is **`KET.Application`** (verified; case-insensitive). Fallback for old installs: `Et.Application`.

```powershell
function Connect-WpsSpreadsheets {
    foreach ($progId in @("KET.Application", "Et.Application")) {
        try { return New-Object -ComObject $progId }
        catch { }
    }
    throw "WPS Spreadsheets COM not available. Is WPS Office installed?"
}
```

Do NOT use `Excel.Application` — on dual-install machines it may resolve to MS Excel or WPS unpredictably. Always use the explicit WPS ProgID.

## 2. Session pattern

```powershell
$et = $null
try {
    $et = Connect-WpsSpreadsheets
    $et.Visible = $false
    $et.DisplayAlerts = 0
    $et.ScreenUpdating = $false   # big speedup for batch writes

    $wb = $et.Workbooks.Add()
    # ... do work ...
    $wb.Close()
}
finally {
    if ($et) {
        $et.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($et) | Out-Null
    }
}
```

Same session etiquette as Writer: check `Get-Process -Name "et"` first — if the user already has WPS Spreadsheets open, don't `Quit()` their instance, only close your own workbooks.

## 3. Recipes

### 3.1 Batch write (the fast way)

Never write cell-by-cell in a loop — it's an order of magnitude slower. Build a 2D array and assign it once:

```powershell
$wb = $et.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)   # always .Item(n) — direct indexing like Worksheets(1) is unreliable in WPS
$ws.Name = "销售数据"

$data = @(
    @("产品", "销量", "金额"),
    @("产品A", 1500, 300000),
    @("产品B", 900, 180000)
)
$ws.Range("A1:C3").Value2 = $data
```

### 3.2 Read a whole sheet at once

```powershell
$all = $ws.UsedRange.Value2   # 2D array; $all[1,1] is A1 (1-based)
$rowCount = $ws.UsedRange.Rows.Count
```

### 3.3 Formulas and number formats

```powershell
$ws.Range("C4").Formula = "=SUM(C2:C3)"
$ws.Range("C2:C4").NumberFormat = "#,##0"
```

⚠️ WPS ships ~400+ functions but **not** the full Excel set: `LAMBDA` / `WEBSERVICE` are missing, and dynamic-array functions (`FILTER`, `SORT`, `XLOOKUP`) have incomplete support with known bugs in nested cases. Test any non-trivial formula on the target machine before relying on it. PivotTables are weaker than Excel's — mark them "limited support" and avoid when possible.

**Text that starts with `=`:** WPS parses it as a formula and throws a COM exception on invalid syntax. Prefix a literal leading apostrophe:

```powershell
$ws.Range("A5").Value2 = "'=不是公式"   # the ' forces text mode, like typing it in the UI
```

### 3.4 Styling: header row

```powershell
$header = $ws.Range("A1:C1")
$header.Font.Bold = $true
$header.Font.NameFarEast = "黑体"
$header.Interior.Color = 65535   # BGR long integer: yellow = 65535, red = 255, blue = 16711680
```

Two WPS quirks to respect:
- **Bold check:** WPS may return `0`, `-1`, `$true`, or `$false` for `Font.Bold`. Never compare with `== $true` — use truthiness: `if ($cell.Font.Bold) { ... }`.
- **Borders:** `.Borders` / `.BorderAround()` have crash reports in WPS scripting environments. Smoke-test border code on your machine before shipping; when in doubt, skip borders.

### 3.5 Chart

```powershell
# AddChart2(Style, ChartType, Left, Top, Width, Height)
# Style MUST be 0 (default) — Style = -1 returns null in WPS. 51 = xlColumnClustered (bar chart).
$chart = $ws.Shapes.AddChart2(0, 51, 350, 20, 480, 300).Chart
$chart.SetSourceData($ws.Range("A1:C3"))
$chart.HasTitle = $true
$chart.ChartTitle.Text = "季度销量"
$chart.ChartTitle.Font.NameFarEast = "微软雅黑"
```

### 3.6 Save and export

```powershell
$wb.SaveAs("D:\out\销售.xlsx", 51)              # 51 = xlOpenXMLWorkbook (.xlsx)
$wb.ExportAsFixedFormat(0, "D:\out\销售.pdf")  # 0 = xlTypePDF
```

⚠️ **PDF watermark (unverified):** community reports suggest the free edition may stamp exported PDFs. Verify on your machine before promising clean PDFs.

## 4. Gotchas

| Problem | Cause / fix |
|---|---|
| `et.exe` lingers / next run hangs | `Quit()` not in `finally`. Verified 2026-09-28: `et.exe` can linger even after `Quit()` — clean up with the PID-snapshot pattern (kill only PIDs you started): `$before=(Get-Process et -EA SilentlyContinue).Id; ... Start-Sleep 3; Get-Process et -EA SilentlyContinue | Where-Object {$before -notcontains $_.Id} | Stop-Process -Force` |
| "Cannot create ActiveX component" | 32-bit WPS vs 64-bit PowerShell — retry with `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`. |
| `Worksheets(1)` throws | Use `.Item(1)` for all collections in WPS. |
| Writing `"=abc"` to a cell throws | Prefix with `'`: `"''=abc"`. |
| Bold detection logic misfires | Don't compare `Font.Bold -eq $true`; use truthiness. |
| Chart creation returns null | You passed Style `-1` to `AddChart2`. Use `0`. |
| Formula works in Excel, fails in WPS | Missing/partial function support — test on WPS, simplify if needed. |
| Slow on large sheets | You forgot `ScreenUpdating = $false`, or you're writing cell-by-cell. Batch it. |

## 5. What NOT to promise

- **Server/headless use:** needs a logged-in interactive desktop session. No Session 0, no "works on a server" claims.
- **Cloud workbooks:** COM only touches local files — sync first.
- **Full Excel parity:** ~400 functions, weaker pivots, partial dynamic arrays. Say what was tested, not what "should" work.

## 6. Checklist before shipping

- [ ] Uses `KET.Application` explicitly (never `Excel.Application`)?
- [ ] Collections accessed via `.Item(n)`?
- [ ] Batch writes via 2D array, not loops?
- [ ] `AddChart2` called with Style `0`, not `-1`?
- [ ] Text starting with `=` gets the `'` prefix?
- [ ] `Font.Bold` checked by truthiness, not `-eq $true`?
- [ ] Tested on a real machine with WPS installed?
