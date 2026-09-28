# Invoke-WpsSession.ps1 — WPS 演示会话包装器
#
# ⚠️ 和文字/表格不同：演示设置 Visible=$false 会报 HRESULT E_FAIL（2026-09-28 真机验证），
#    所以本包装器永远不碰 Visible，WPS 会以可见窗口运行。调用前先告诉用户：
#    "WPS 会弹出来一下，跑完自动关。"
#
# Agent 不要手拼 New-Object / Quit / ReleaseComObject，直接调这个：
#
#   $sb = {
#       param($wpp)
#       $pres = $wpp.Presentations.Add()
#       try {
#           # ……业务操作……
#           $pres.SaveAs("$env:TEMP\out.pptx")
#       } finally {
#           $pres.Close()
#           Release-WpsObject $pres
#       }
#   }.GetNewClosure()
#   & "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
#
# 注意 .GetNewClosure()：否则外层参数传不进 scriptblock。
# 输出 OK:/FAIL: 行；exit code 0=成功，1=失败。
param(
    [Parameter(Mandatory)][scriptblock]$Script
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Report($ok, $name, $detail = "") {
    if ($ok) { Write-Output "OK: $name $detail" } else { Write-Output "FAIL: $name $detail" }
}

# 倒序释放 helper：调用方按"获取的逆序"逐个释放业务对象，$wpp 由 finally 统一释放
function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

Write-Output "NOTE: WPS 演示必须在可见窗口运行，WPS 会弹出来一下，跑完自动关。"

$procName = "wpp"
$before = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$owned = $before.Count -eq 0
$app = $null
$code = 0
try {
    foreach ($progId in @("KWPP.Application", "Wpp.Application")) {
        try { $app = New-Object -ComObject $progId; Report $true "connect" $progId; break }
        catch { }
    }
    if (-not $app) { throw "WPS 演示 COM 不可用：请确认安装了 WPS Office" }
    # 注意：这里故意不设置 $app.Visible（会 E_FAIL）
    if ($owned) { $app.DisplayAlerts = 0 }
    & $Script $app
    Report $true "session" "done"
} catch {
    Report $false "session" $_.Exception.Message
    $code = 1
} finally {
    if ($app) {
        if (-not $owned) {
            Report $true "session" "attached to user's running WPS; skipped Quit"
        } else {
            try { $app.Quit() } catch { }
        }
        Release-WpsObject $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 2
        # 只杀本次新增的 PID，绝不碰用户原有的 wpp.exe
        $leftover = @(Get-Process -Name $procName -ErrorAction SilentlyContinue |
            Where-Object { $before -notcontains $_.Id })
        if ($leftover.Count -gt 0) {
            $leftover | Stop-Process -Force
            Report $true "cleanup" "killed $($leftover.Count) lingering process(es)"
        }
    }
}
exit $code
