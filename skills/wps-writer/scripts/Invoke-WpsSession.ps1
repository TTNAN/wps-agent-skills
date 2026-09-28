# Invoke-WpsSession.ps1 — WPS 文字会话包装器
#
# Agent 不要手拼 New-Object / Quit / ReleaseComObject，直接调这个：
#
#   $sb = {
#       param($app)
#       $doc = $app.Documents.Add()
#       try {
#           # ……业务操作……
#           $doc.SaveAs("$env:TEMP\out.docx", 16)
#       } finally {
#           $doc.Close()
#           Release-WpsObject $doc   # 倒序释放：先业务对象，后 $app（$app 由本包装器释放）
#       }
#   }.GetNewClosure()
#   & "$PSScriptRoot\Invoke-WpsSession.ps1" -Script $sb
#
# 注意 .GetNewClosure()：否则外层参数（$OutputPath 等）传不进 scriptblock。
# 输出 OK:/FAIL: 行；exit code 0=成功，1=失败。
param(
    [Parameter(Mandatory)][scriptblock]$Script
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Report($ok, $name, $detail = "") {
    if ($ok) { Write-Output "OK: $name $detail" } else { Write-Output "FAIL: $name $detail" }
}

# 倒序释放 helper：调用方按"获取的逆序"逐个释放业务对象，$app 由 finally 统一释放
function Release-WpsObject($obj) {
    if ($null -ne $obj) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
    }
}

$procName = "wps"
$before = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$owned = $before.Count -eq 0   # $false = 附着到了用户已开的实例，不要动它的窗口/弹窗设置
$app = $null
$code = 0
try {
    foreach ($progId in @("KWPS.Application", "WPS.Application")) {
        try { $app = New-Object -ComObject $progId; Report $true "connect" $progId; break }
        catch { }
    }
    if (-not $app) { throw "WPS 文字 COM 不可用：请确认安装了 WPS Office" }
    if ($owned) {
        $app.Visible = $false
        $app.DisplayAlerts = 0
    }
    & $Script $app
    Report $true "session" "done"
} catch {
    Report $false "session" $_.Exception.Message
    $code = 1
} finally {
    if ($app) {
        if (-not $owned) {
            # 用户已经开着 WPS：New-Object 附着到了他的实例，跳过 Quit，只关自己打开的文档
            Report $true "session" "attached to user's running WPS; skipped Quit"
        } else {
            try { $app.Quit() } catch { }
        }
        Release-WpsObject $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 2
        # 只杀本次新增的 PID，绝不碰用户原有的 wps.exe
        $leftover = @(Get-Process -Name $procName -ErrorAction SilentlyContinue |
            Where-Object { $before -notcontains $_.Id })
        if ($leftover.Count -gt 0) {
            $leftover | Stop-Process -Force
            Report $true "cleanup" "killed $($leftover.Count) lingering process(es)"
        }
    }
}
exit $code
