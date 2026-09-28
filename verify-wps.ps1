# verify-wps.ps1 — WPS COM 真机验证脚本
# 在装有 WPS 的 Windows 上用 powershell -ExecutionPolicy Bypass -File verify-wps.ps1 运行
# 只输出 OK: / FAIL: 行，方便直接贴回结果
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Report($ok, $name, $detail = "") {
    if ($ok) { Write-Output "OK: $name $detail" } else { Write-Output "FAIL: $name $detail" }
}

# 0. 环境
Report $true "powershell-version" $PSVersionTable.PSVersion.ToString()
$is64 = [IntPtr]::Size -eq 8
Report $true "powershell-bitness" $(if ($is64) { "64-bit" } else { "32-bit" })

# 1. ProgID 注册表探测
$progids = @{
    "writer"       = "KWPS.Application"
    "spreadsheets" = "KET.Application"
    "presentation" = "KWPP.Application"
}
foreach ($k in $progids.Keys) {
    $p = $progids[$k]
    $hit = (Test-Path "HKLM:\Software\Classes\$p") -or (Test-Path "HKCU:\Software\Classes\$p")
    Report $hit "progid-registered" "$k -> $p"
}

# 2. 逐个应用：启动 → 新建 → 写内容 → 保存 → 退出
$tmp = Join-Path $env:TEMP "wps-verify"
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

function Test-WpsApp($progid, $label, [scriptblock]$work, [switch]$KeepVisible) {
    # 2026-09-28 真机结论：
    # - writer / spreadsheets 支持 Visible=$false 后台运行
    # - presentation 设置 Visible=$false 会报 HRESULT E_FAIL，必须可见窗口跑
    $app = $null
    try {
        $app = New-Object -ComObject $progid
        if (-not $KeepVisible) { $app.Visible = $false }
        try { $app.DisplayAlerts = 0 } catch {}
        & $work $app
        Report $true "app-$label" "create+save works"
    } catch {
        Report $false "app-$label" $_.Exception.Message
    } finally {
        if ($app) {
            try { $app.Quit() } catch {}
            [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
        }
    }
}

Test-WpsApp "KWPS.Application" "writer" {
    param($wps)
    $doc = $wps.Documents.Add()
    $p = $doc.Content.Paragraphs.Add()
    $p.Range.Text = "WPS COM 验证文档"
    $path = Join-Path $tmp "verify.docx"
    if (Test-Path $path) { Remove-Item $path -Force }
    $doc.SaveAs($path, 16)
    $doc.Close()
    if (-not (Test-Path $path)) { throw "save produced no file" }
}

Test-WpsApp "KET.Application" "spreadsheets" {
    param($et)
    $wb = $et.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Range("A1").Value2 = "验证"
    $ws.Range("B1").Formula = "=1+1"
    $path = Join-Path $tmp "verify.xlsx"
    if (Test-Path $path) { Remove-Item $path -Force }
    $wb.SaveAs($path, 51)
    $wb.Close()
    if (-not (Test-Path $path)) { throw "save produced no file" }
}

Test-WpsApp "KWPP.Application" "presentation" {
    param($wpp)
    $pres = $wpp.Presentations.Add()
    $s = $pres.Slides.Add(1, 1)
    $s.Shapes.Title.TextFrame.TextRange.Text = "验证标题"
    $path = Join-Path $tmp "verify.pptx"
    if (Test-Path $path) { Remove-Item $path -Force }
    $pres.SaveAs($path)
    $pres.Close()
    if (-not (Test-Path $path)) { throw "save produced no file" }
} -KeepVisible

# 3. 残留进程检查
Start-Sleep -Seconds 2
foreach ($proc in @("wps", "et", "wpp")) {
    $left = Get-Process -Name $proc -ErrorAction SilentlyContinue
    Report ($null -eq $left) "no-leftover-process" $proc
}

Write-Output "DONE. 临时文件在 $tmp"
