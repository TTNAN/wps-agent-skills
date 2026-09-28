<#
.SYNOPSIS
    一键安装 WPS Agent Skills 到当前用户的 Claude Code skills 目录。
.DESCRIPTION
    把 skills/*/ 拷到 $env:USERPROFILE\.claude\skills\<skill-name>\。
    可重复运行，已安装的会被覆盖。
    兼容 Windows PowerShell 5.1。
#>
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSCommandPath
$destRoot = Join-Path $env:USERPROFILE ".claude\skills"

$skills = @("powershell-windows", "wps-writer", "wps-spreadsheets", "wps-presentation")
$ok = 0
$fail = 0

foreach ($name in $skills) {
    $src = Join-Path $repoRoot "skills\$name"
    $dst = Join-Path $destRoot $name
    try {
        if (-not (Test-Path $src)) { throw "source not found: $src" }
        if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Path $dst -Force | Out-Null }
        Copy-Item -Recurse -Force (Join-Path $src "*") $dst
        Write-Output "OK: installed $name -> $dst"
        $ok++
    } catch {
        Write-Output "FAIL: $name : $($_.Exception.Message)"
        $fail++
    }
}

Write-Output "----"
Write-Output "Done. OK=$ok FAIL=$fail"
Write-Output "请重启你的 Agent（Claude Code / Codex），让它加载新 skill。"

if ($fail -gt 0) { exit 1 }
