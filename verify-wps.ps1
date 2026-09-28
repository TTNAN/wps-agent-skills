# verify-wps.ps1 — WPS COM 真机验证脚本
# 在装有 WPS 的 Windows 上用 powershell -ExecutionPolicy Bypass -File verify-wps.ps1 运行
# 只输出 OK: / FAIL: 行，方便直接贴回结果
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Report($ok, $name, $detail = "") {
    if ($ok) { Write-Output "OK: $name $detail" } else { Write-Output "FAIL: $name $detail" }
}

function Warn($name, $detail = "") {
    Write-Output "WARN: $name $detail"
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

# 启动前 PID 快照：残留检查只关心本次新建的进程，不关心用户本来就开着的
$procNames = @("wps", "et", "wpp")
$idsBefore = @{}
foreach ($pn in $procNames) {
    $idsBefore[$pn] = @(Get-Process -Name $pn -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

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

# 3. 残留进程检查：只看本次新建的 PID
# 已知现象（2026-09-28 真机）：wps.exe / et.exe 在 Quit() 后可能残留，这是预期内的，
# 标 WARN 不算失败；只有"本次新建的 PID 在强制清理后还活着"才算 FAIL。
Start-Sleep -Seconds 3
foreach ($pn in $procNames) {
    $now = @(Get-Process -Name $pn -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $new = @($now | Where-Object { $idsBefore[$pn] -notcontains $_ })
    if ($new.Count -eq 0) { Report $true "no-leftover-process" $pn; continue }
    $new | ForEach-Object { try { Stop-Process -Id $_ -Force } catch { } }
    Start-Sleep -Seconds 2
    $still = @(Get-Process -Name $pn -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } |
        Where-Object { $new -contains $_ })
    if ($still.Count -gt 0) {
        Report $false "leftover-process" "$pn 强制清理后仍残留 PID: $($still -join ',')"
    } else {
        Warn "leftover-process" "$pn 有残留（预期内现象），已强制清理"
    }
}

Write-Output "DONE. 临时文件在 $tmp"
