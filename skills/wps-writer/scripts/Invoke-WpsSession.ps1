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

# 注意：会话规则（僵尸判定 / PID 快照 / 懒清理 / 差集清理）在三个 skill 的 Invoke-WpsSession.ps1 里各有一份，
# 改一处必须三处同步（wps-writer / wps-spreadsheets / wps-presentation）。
$procName = "wps"
function Get-WpsSnapshot($name) {
    @(Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{ Id = $_.Id; HasWindow = ($_.MainWindowHandle -ne 0) }
    })
}
$before = Get-WpsSnapshot $procName
$beforeIds = @($before | ForEach-Object { $_.Id })
# $owned=$false 仅当：启动前存在带可见主窗口的进程（用户真的在用）。
# 无窗口的预存进程 ≠ 一定是僵尸（可能是托盘驻留/云同步/还没退完），开局不杀——
# 只有 New-Object 全部失败时，才清理它们再重试（见"懒清理"）。
$owned = @($before | Where-Object { $_.HasWindow }).Count -eq 0
$app = $null
$code = 0
try {
    foreach ($progId in @("KWPS.Application", "WPS.Application")) {
        try { $app = New-Object -ComObject $progId; Report $true "connect" $progId; break }
        catch { }
    }
    if (-not $app) {
        # 懒清理：连接失败才动刀。无可见主窗口的预存进程很可能是上轮崩溃的僵尸，
        # 清掉再试一次；带窗口的（用户正在用）绝不碰。
        foreach ($z in @($before | Where-Object { -not $_.HasWindow })) {
            try { Stop-Process -Id $z.Id -Force; Report $true "zombie-cleanup" "killed stale PID $($z.Id) (no main window)" } catch { }
        }
        Start-Sleep -Seconds 1
        foreach ($progId in @("KWPS.Application", "WPS.Application")) {
            try { $app = New-Object -ComObject $progId; Report $true "connect-retry" $progId; break }
            catch { }
        }
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
            Where-Object { $beforeIds -notcontains $_.Id })
        if ($leftover.Count -gt 0) {
            $leftover | Stop-Process -Force
            Report $true "cleanup" "killed $($leftover.Count) lingering process(es)"
        }
    }
}
exit $code
