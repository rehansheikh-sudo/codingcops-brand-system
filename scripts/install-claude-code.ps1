# Make the CodingCops brand skill available in every Claude Code session
# (Windows), and install the brand fonts for the current user.
#
#   .\scripts\install-claude-code.ps1            # skill + fonts
#   .\scripts\install-claude-code.ps1 -NoFonts   # skill only
#
# Uses a directory junction (no admin rights needed), so `git pull` updates it.
param([switch]$NoFonts)
$ErrorActionPreference = "Stop"

$Repo    = Split-Path -Parent $PSScriptRoot
$Src     = Join-Path $Repo ".claude\skills\codingcops-brand"
$DestDir = Join-Path $HOME ".claude\skills"
$Dest    = Join-Path $DestDir "codingcops-brand"

if (-not (Test-Path (Join-Path $Src "SKILL.md"))) { throw "Can't find $Src\SKILL.md. Run this from the cloned repo." }
New-Item -ItemType Directory -Force -Path $DestDir | Out-Null

if (Test-Path $Dest) {
    $item = Get-Item $Dest -Force
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        $item.Delete()
    } else {
        $bk = "$Dest.backup-$(Get-Date -Format yyyyMMddHHmmss)"
        Move-Item $Dest $bk
        Write-Host "Moved an existing copy to $bk"
    }
}
New-Item -ItemType Junction -Path $Dest -Target $Src | Out-Null
Write-Host "Skill linked: $Dest -> $Src"

if (-not $NoFonts) {
    $FontDir = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
    $RegPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    New-Item -ItemType Directory -Force -Path $FontDir | Out-Null
    Get-ChildItem (Join-Path $Src "assets\fonts") -Include *.otf, *.ttf -Recurse | ForEach-Object {
        $target = Join-Path $FontDir $_.Name
        if (-not (Test-Path $target)) { Copy-Item $_.FullName $target }
        $kind = if ($_.Extension -eq ".otf") { "OpenType" } else { "TrueType" }
        New-ItemProperty -Path $RegPath -Name "$($_.BaseName) ($kind)" -Value $target -PropertyType String -Force | Out-Null
    }
    Write-Host "Fonts installed for $env:USERNAME (restart PowerPoint/Word to see them)"
}

Write-Host ""
Write-Host "Done. In any folder, run 'claude' and ask, for example:"
Write-Host "  /codingcops-brand Create a 10-slide pitch deck for Acme. Save it here as acme-pitch.pptx."
Write-Host "Update later with: git -C `"$Repo`" pull --ff-only"
