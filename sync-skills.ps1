<#
.SYNOPSIS
    Syncs selected skills from this repo to ~/.copilot/skills/ for cross-workspace availability.

.DESCRIPTION
    Copies skill folders listed in $SyncSkills from the repo's skills/ directory
    to the user-level Copilot skills folder. Skills in $CustomSkills are always
    synced on first copy (when they don't exist in target yet), but protected
    from overwrite on subsequent syncs unless -Force is specified.

.PARAMETER Force
    Also overwrite custom skills (use with caution).

.PARAMETER DryRun
    Show what would be synced without copying anything.

.EXAMPLE
    .\sync-skills.ps1            # sync upstream skills only
    .\sync-skills.ps1 -DryRun    # preview what would be synced
    .\sync-skills.ps1 -Force     # sync everything, including custom skills
#>
param(
    [switch]$Force,
    [switch]$DryRun
)

$RepoSkillsDir = Join-Path $PSScriptRoot "skills"
$TargetDir = Join-Path $HOME ".copilot" "skills"

# ── Skills to sync from this repo (add more as needed) ──
$SyncSkills = @(
    "doc-coauthoring"
    "docx"
    "pdf"
    "pptx"
    "xlsx"
    "skill-creator"
)

# ── Custom skills: synced only with -Force (your modified copies) ──
$CustomSkills = @(
    "docx-msft"
    "docx-foundry"
)

# ── Workspace skills: synced to specific repos instead of ~/.copilot/skills/ ──
# Key = skill name, Value = array of target repo roots (each gets .github/skills/<name>/)
$WorkspaceSkills = @{
    "cost-proposal-kutlu" = @(
        (Join-Path $HOME "1Repositories" "rth-opportunities")
    )
}

if (-not (Test-Path $TargetDir)) {
    if ($DryRun) {
        Write-Host "Would create: $TargetDir" -ForegroundColor Yellow
    } else {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }
}

$allSkills = $SyncSkills + $CustomSkills

# Helper: get the newest file timestamp in a folder (recursive)
function Get-NewestTimestamp([string]$Path) {
    Get-ChildItem $Path -Recurse -File | Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty LastWriteTime
}

# Helper: show what changed between source (repo master) and destination (modified copy)
function Show-Divergence([string]$Source, [string]$Dest) {
    $srcFiles = Get-ChildItem $Source -Recurse -File | ForEach-Object {
        [PSCustomObject]@{ Rel = $_.FullName.Substring($Source.Length + 1); Full = $_.FullName }
    }
    $dstFiles = Get-ChildItem $Dest -Recurse -File | ForEach-Object {
        [PSCustomObject]@{ Rel = $_.FullName.Substring($Dest.Length + 1); Full = $_.FullName }
    }
    $srcMap = @{}; $srcFiles | ForEach-Object { $srcMap[$_.Rel] = $_.Full }
    $dstMap = @{}; $dstFiles | ForEach-Object { $dstMap[$_.Rel] = $_.Full }

    # Files only in destination (added by agent)
    $added = $dstFiles | Where-Object { -not $srcMap.ContainsKey($_.Rel) }
    # Files only in source (deleted from destination)
    $removed = $srcFiles | Where-Object { -not $dstMap.ContainsKey($_.Rel) }
    # Files in both — check for content differences
    $modified = @()
    foreach ($f in $srcFiles) {
        if ($dstMap.ContainsKey($f.Rel)) {
            $srcHash = (Get-FileHash $f.Full -Algorithm SHA256).Hash
            $dstHash = (Get-FileHash $dstMap[$f.Rel] -Algorithm SHA256).Hash
            if ($srcHash -ne $dstHash) {
                $modified += $f.Rel
            }
        }
    }

    $hasChanges = ($added.Count -gt 0) -or ($removed.Count -gt 0) -or ($modified.Count -gt 0)
    if (-not $hasChanges) {
        Write-Host "    (no content differences — only timestamps differ)" -ForegroundColor DarkGray
        return $false
    }

    if ($added.Count -gt 0) {
        Write-Host "    Files added in destination (will be LOST):" -ForegroundColor Red
        $added | ForEach-Object { Write-Host "      + $($_.Rel)" -ForegroundColor Red }
    }
    if ($removed.Count -gt 0) {
        Write-Host "    Files missing from destination (will be restored):" -ForegroundColor Cyan
        $removed | ForEach-Object { Write-Host "      - $($_.Rel)" -ForegroundColor Cyan }
    }
    if ($modified.Count -gt 0) {
        Write-Host "    Files modified in destination (will be OVERWRITTEN):" -ForegroundColor Yellow
        foreach ($rel in $modified) {
            $dstTime = (Get-Item $dstMap[$rel]).LastWriteTime
            Write-Host "      ~ $rel  (modified $($dstTime.ToString('yyyy-MM-dd HH:mm')))" -ForegroundColor Yellow
        }
        # Show inline diff for text files (first modified file, limited to keep output manageable)
        $firstMod = $modified[0]
        $srcContent = Get-Content $srcMap[$firstMod] -ErrorAction SilentlyContinue
        $dstContent = Get-Content $dstMap[$firstMod] -ErrorAction SilentlyContinue
        if ($srcContent -and $dstContent) {
            Write-Host "    Preview diff for $firstMod (destination lines that differ):" -ForegroundColor DarkGray
            $diff = Compare-Object $srcContent $dstContent -PassThru | Select-Object -First 15
            foreach ($line in $diff) {
                $indicator = $line.SideIndicator  # <= means repo, => means destination
                if ($indicator -eq '=>') {
                    Write-Host "      + $line" -ForegroundColor Green   # destination-only (will be lost)
                } elseif ($indicator -eq '<=') {
                    Write-Host "      - $line" -ForegroundColor DarkRed # repo-only (will be restored)
                }
            }
            $totalDiffs = (Compare-Object $srcContent $dstContent).Count
            if ($totalDiffs -gt 15) {
                Write-Host "      ... and $($totalDiffs - 15) more changed lines" -ForegroundColor DarkGray
            }
            if ($modified.Count -gt 1) {
                Write-Host "      ($($modified.Count - 1) more modified file(s) not shown)" -ForegroundColor DarkGray
            }
        }
    }
    return $true
}

# Helper: prompt user for confirmation on diverged skills (skipped in DryRun)
function Confirm-Overwrite([string]$Skill, [string]$Source, [string]$Dest) {
    Write-Host ""
    Write-Host "  WARNING  $Skill — destination has local changes:" -ForegroundColor Magenta
    $hasChanges = Show-Divergence $Source $Dest
    if (-not $hasChanges) {
        return $true  # timestamp-only divergence, safe to overwrite
    }
    Write-Host ""
    $answer = Read-Host "    Overwrite with repo master? [y/N]"
    return ($answer -eq 'y' -or $answer -eq 'Y')
}

$synced = 0
$skipped = 0
$upToDate = 0

foreach ($skill in $allSkills) {
    $source = Join-Path $RepoSkillsDir $skill
    $dest = Join-Path $TargetDir $skill
    $isCustom = $CustomSkills -contains $skill

    if (-not (Test-Path $source)) {
        Write-Host "  MISSING  $skill (not found in repo)" -ForegroundColor Red
        $skipped++
        continue
    }

    # Check if destination exists and compare timestamps
    $status = "NEW"
    if (Test-Path $dest) {
        $repoTime = Get-NewestTimestamp $source
        $destTime = Get-NewestTimestamp $dest
        if ($repoTime -and $destTime) {
            if ($repoTime -gt $destTime) {
                $status = "UPDATED"
            } elseif ($destTime -gt $repoTime -and -not $isCustom) {
                # Destination was modified (e.g., by an agent) — repo is master, overwrite
                $status = "DIVERGED"
            } else {
                $status = "CURRENT"
            }
        }
    }

    # Custom skills: always sync if new, protect existing unless -Force
    if ($isCustom -and $status -ne "NEW" -and -not $Force) {
        if ($DryRun) {
            if ($status -eq "CURRENT") {
                Write-Host "  UP TO DATE  $skill (custom)" -ForegroundColor DarkGray
            } else {
                Write-Host "  SKIPPED  $skill (custom, repo is newer — use -Force to overwrite)" -ForegroundColor DarkYellow
            }
        } else {
            if ($status -eq "CURRENT") {
                Write-Host "  UP TO DATE  $skill (custom)" -ForegroundColor DarkGray
            } else {
                Write-Host "  SKIPPED  $skill (custom, repo is newer — use -Force to overwrite)" -ForegroundColor DarkYellow
            }
        }
        $upToDate++
        continue
    }

    if ($DryRun) {
        switch ($status) {
            "NEW"      { Write-Host "  WOULD SYNC  $skill (new)" -ForegroundColor Cyan }
            "UPDATED"  { Write-Host "  WOULD SYNC  $skill (repo is newer)" -ForegroundColor Yellow }
            "DIVERGED" {
                Write-Host "  WOULD SYNC  $skill (destination was modified — would reset to repo master)" -ForegroundColor Magenta
                Show-Divergence $source $dest | Out-Null
            }
            "CURRENT"  { Write-Host "  UP TO DATE  $skill" -ForegroundColor DarkGray }
        }
        if ($status -ne "CURRENT") { $synced++ } else { $upToDate++ }
        continue
    }

    if ($status -eq "CURRENT") {
        Write-Host "  UP TO DATE  $skill" -ForegroundColor DarkGray
        $upToDate++
        continue
    }

    # For diverged skills, show diff and confirm before overwriting
    if ($status -eq "DIVERGED" -and -not $Force) {
        if (-not (Confirm-Overwrite $skill $source $dest)) {
            Write-Host "  SKIPPED  $skill (kept local changes)" -ForegroundColor DarkYellow
            $skipped++
            continue
        }
    }

    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Copy-Item $source $dest -Recurse
    switch ($status) {
        "NEW"      { Write-Host "  SYNCED  $skill (new)" -ForegroundColor Green }
        "UPDATED"  { Write-Host "  SYNCED  $skill (updated)" -ForegroundColor Yellow }
        "DIVERGED" { Write-Host "  SYNCED  $skill (reset to repo master)" -ForegroundColor Magenta }
    }
    $synced++
}

Write-Host ""
Write-Host "Done: $synced synced, $upToDate up to date, $skipped missing" -ForegroundColor White

# ── Workspace-level skills ──────────────────────────────────────────
if ($WorkspaceSkills.Count -gt 0) {
    Write-Host ""
    Write-Host "Workspace skills:" -ForegroundColor White

    $wsSynced = 0
    $wsUpToDate = 0
    $wsSkipped = 0

    foreach ($skill in $WorkspaceSkills.Keys) {
        $source = Join-Path $RepoSkillsDir $skill

        if (-not (Test-Path $source)) {
            Write-Host "  MISSING  $skill (not found in repo)" -ForegroundColor Red
            $wsSkipped++
            continue
        }

        foreach ($repoRoot in $WorkspaceSkills[$skill]) {
            $repoName = Split-Path $repoRoot -Leaf
            $dest = Join-Path $repoRoot ".github" "skills" $skill

            if (-not (Test-Path $repoRoot)) {
                Write-Host "  MISSING  $skill -> $repoName (repo not found: $repoRoot)" -ForegroundColor Red
                $wsSkipped++
                continue
            }

            # Ensure .github/skills/ exists
            $skillsDir = Join-Path $repoRoot ".github" "skills"
            if (-not (Test-Path $skillsDir)) {
                if (-not $DryRun) {
                    New-Item -ItemType Directory -Path $skillsDir -Force | Out-Null
                }
            }

            # Compare timestamps
            $status = "NEW"
            if (Test-Path $dest) {
                $repoTime = Get-NewestTimestamp $source
                $destTime = Get-NewestTimestamp $dest
                if ($repoTime -and $destTime) {
                    if ($repoTime -gt $destTime) {
                        $status = "UPDATED"
                    } elseif ($destTime -gt $repoTime) {
                        $status = "DIVERGED"
                    } else {
                        $status = "CURRENT"
                    }
                }
            }

            if ($DryRun) {
                switch ($status) {
                    "NEW"      { Write-Host "  WOULD SYNC  $skill -> $repoName (new)" -ForegroundColor Cyan }
                    "UPDATED"  { Write-Host "  WOULD SYNC  $skill -> $repoName (repo is newer)" -ForegroundColor Yellow }
                    "DIVERGED" {
                        Write-Host "  WOULD SYNC  $skill -> $repoName (destination was modified)" -ForegroundColor Magenta
                        Show-Divergence $source $dest | Out-Null
                    }
                    "CURRENT"  { Write-Host "  UP TO DATE  $skill -> $repoName" -ForegroundColor DarkGray }
                }
                if ($status -ne "CURRENT") { $wsSynced++ } else { $wsUpToDate++ }
                continue
            }

            if ($status -eq "CURRENT") {
                Write-Host "  UP TO DATE  $skill -> $repoName" -ForegroundColor DarkGray
                $wsUpToDate++
                continue
            }

            # For diverged workspace skills, show diff and confirm
            if ($status -eq "DIVERGED" -and -not $Force) {
                if (-not (Confirm-Overwrite "$skill -> $repoName" $source $dest)) {
                    Write-Host "  SKIPPED  $skill -> $repoName (kept local changes)" -ForegroundColor DarkYellow
                    $wsSkipped++
                    continue
                }
            }

            if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
            Copy-Item $source $dest -Recurse
            switch ($status) {
                "NEW"      { Write-Host "  SYNCED  $skill -> $repoName (new)" -ForegroundColor Green }
                "UPDATED"  { Write-Host "  SYNCED  $skill -> $repoName (updated)" -ForegroundColor Yellow }
                "DIVERGED" { Write-Host "  SYNCED  $skill -> $repoName (reset to repo master)" -ForegroundColor Magenta }
            }
            $wsSynced++
        }
    }

    Write-Host ""
    Write-Host "Workspace: $wsSynced synced, $wsUpToDate up to date, $wsSkipped missing" -ForegroundColor White
}
