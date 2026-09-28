# diag-wpp.ps1 — WPS 演示组件分步诊断
# Win+R 运行: powershell -NoExit -ExecutionPolicy Bypass -File "D:\Edge Download\diag-wpp.ps1"
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Report($ok, $name, $detail = "") {
    if ($ok) { Write-Output "OK: $name $detail" } else { Write-Output "FAIL: $name $detail" }
}

$before = (Get-Process -Name "wpp" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

$wpp = $null
try { $wpp = New-Object -ComObject "KWPP.Application"; Report $true "new-object" }
catch { Report $false "new-object" $_.Exception.Message; exit 1 }

try { $wpp.Visible = $false; Report $true "visible-false" }
catch { Report $false "visible-false" $_.Exception.Message }

try { $wpp.DisplayAlerts = 0; Report $true "display-alerts-0" }
catch { Report $false "display-alerts-0" $_.Exception.Message }

$pres = $null
try { $pres = $wpp.Presentations.Add(); Report $true "presentations-add" }
catch { Report $false "presentations-add" $_.Exception.Message }

if ($pres) {
    # 先试不可见模式
    try {
        $s = $pres.Slides.Add(1, 1)
        Report $true "slides-add-invisible" "layout=1"
    } catch {
        Report $false "slides-add-invisible" $_.Exception.Message
        # 重试可见模式
        try {
            $wpp.Visible = $true
            Start-Sleep -Seconds 2
            $s = $pres.Slides.Add(1, 1)
            Report $true "slides-add-visible-retry" "layout=1"
        } catch {
            Report $false "slides-add-visible-retry" $_.Exception.Message
            $s = $null
        }
    }

    if ($s) {
        try { $s.Shapes.Title.TextFrame.TextRange.Text = "诊断标题"; Report $true "title-set" }
        catch { Report $false "title-set" $_.Exception.Message }

        try {
            $out = "$env:TEMP\wps-verify\diag.pptx"
            if (Test-Path $out) { Remove-Item $out -Force }
            $pres.SaveAs($out); Report $true "save-pptx"
        } catch { Report $false "save-pptx" $_.Exception.Message }

        try {
            $outPdf = "$env:TEMP\wps-verify\diag.pdf"
            if (Test-Path $outPdf) { Remove-Item $outPdf -Force }
            $pres.SaveAs($outPdf, 32); Report $true "save-pdf-32"
        } catch { Report $false "save-pdf-32" $_.Exception.Message }
    }

    try { $pres.Close() } catch {}
}

try { $wpp.Quit(); Report $true "quit-called" }
catch { Report $false "quit-called" $_.Exception.Message }
[Runtime.InteropServices.Marshal]::ReleaseComObject($wpp) | Out-Null

Start-Sleep -Seconds 3
$after = (Get-Process -Name "wpp" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$newLeft = $after | Where-Object { $before -notcontains $_ }
Report (($null -eq $newLeft) -or ($newLeft.Count -eq 0)) "no-new-leftover-wpp" "before=$($before.Count) after=$($after.Count)"
Write-Output "DONE"
