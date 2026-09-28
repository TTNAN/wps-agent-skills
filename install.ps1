<#
.SYNOPSIS
    安装 / 卸载 WPS Agent Skills。
.DESCRIPTION
    把 skills/*/ 拷到各 Agent 的 skills 目录：
      claude -> $env:USERPROFILE\.claude\skills
      codex  -> $env:USERPROFILE\.codex\skills
      cursor -> $env:USERPROFILE\.cursor\skills
    不指定 -Agent 时自动探测已安装的 Agent（看目录是否存在），一个都没装则默认 claude。
    安装成功后自动跑 verify-wps.ps1（COM 探活），把问题挡在 Agent 动手之前。
    兼容 Windows PowerShell 5.1。
.EXAMPLE
    .\install.ps1
    .\install.ps1 -Agent cursor
    .\install.ps1 -Skills wps-writer,wps-spreadsheets
    .\install.ps1 -Uninstall
    .\install.ps1 -SkipVerify
#>
param(
    [string[]]$Agent = @(),
    [string[]]$Skills = @(),
    [switch]$Uninstall,
    [switch]$SkipVerify
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSCommandPath
$allSkills = @("powershell-windows", "wps-office", "wps-writer", "wps-spreadsheets", "wps-presentation")
if ($Skills.Count -eq 0) { $Skills = $allSkills }

$agentRoots = @{
    "claude" = Join-Path $env:USERPROFILE ".claude\skills"
    "codex"  = Join-Path $env:USERPROFILE ".codex\skills"
    "cursor" = Join-Path $env:USERPROFILE ".cursor\skills"
}

# 自动探测：已存在的 Agent 目录
if ($Agent.Count -eq 0) {
    $Agent = @()
    foreach ($k in $agentRoots.Keys) {
        if (Test-Path (Split-Path -Parent $agentRoots[$k])) { $Agent += $k }
    }
    if ($Agent.Count -eq 0) { $Agent = @("claude") }
}

$ok = 0
$fail = 0

foreach ($a in $Agent) {
    if (-not $agentRoots.ContainsKey($a)) {
        Write-Output "FAIL: unknown agent '$a' (claude/codex/cursor)"
        $fail++
        continue
    }
    $destRoot = $agentRoots[$a]
    foreach ($name in $Skills) {
        $src = Join-Path $repoRoot "skills\$name"
        $dst = Join-Path $destRoot $name
        try {
            if (-not (Test-Path $src)) { throw "source not found: $src" }
            if ($Uninstall) {
                if (Test-Path $dst) { Remove-Item -Recurse -Force $dst }
                Write-Output "OK: uninstalled $name from $a"
            } else {
                if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Path $dst -Force | Out-Null }
                Copy-Item -Recurse -Force (Join-Path $src "*") $dst
                Write-Output "OK: installed $name -> $dst"
            }
            $ok++
        } catch {
            Write-Output "FAIL: [$a] $name : $($_.Exception.Message)"
            $fail++
        }
    }
}

Write-Output "----"
Write-Output "Done. OK=$ok FAIL=$fail"
if (-not $Uninstall) {
    Write-Output "请重启你的 Agent（Claude Code / Codex / Cursor），让它加载新 skill。"
    if (-not $SkipVerify -and $fail -eq 0) {
        Write-Output "----"
        Write-Output "自动运行 COM 探活（verify-wps.ps1），演示会弹一下 WPS 窗口："
        & (Join-Path $repoRoot "verify-wps.ps1")
    }
}

if ($fail -gt 0) { exit 1 }
