<#
    deploy.ps1 — one-command deploy pipeline for the PS Dashboard.

    Runs, in order:
      1. git add + commit + push   -> GitHub (origin/main)
      2. clasp push                -> syncs files to the Apps Script project
      3. clasp create-version      -> new immutable version
      4. clasp redeploy <id> -V N  -> repoints the existing deployment to that version

    Usage:
      ./deploy.ps1 "Your commit / version message"
      ./deploy.ps1                      # uses a default message
#>
param(
    [string]$Message = "Update PS dashboard"
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# The versioned deployment that should always point at the newest version.
$DEPLOYMENT_ID = "AKfycbw1E_aAKDWebOac5DMcuDyU41YQ-hEh-1zUfGEvk8agGEtoDVVZwPEyadrXgqtktM5XuA"

function Step($text) { Write-Host "`n=== $text ===" -ForegroundColor Cyan }

# 1. Commit + push to GitHub -------------------------------------------------
Step "1/4  Git commit + push"
git add -A
$pending = git status --porcelain
if ($pending) {
    git commit -m $Message
    git push origin main
} else {
    Write-Host "No file changes to commit; skipping commit. Pushing any unpushed commits."
    git push origin main
}

# 2. Push files to the Apps Script project -----------------------------------
Step "2/4  clasp push"
clasp push -f

# 3. Create a new immutable version ------------------------------------------
Step "3/4  clasp create-version"
$versionOutput = clasp create-version $Message 2>&1 | Out-String
Write-Host $versionOutput.Trim()
$match = [regex]::Match($versionOutput, "version\s+(\d+)", "IgnoreCase")
if (-not $match.Success) {
    throw "Could not parse the new version number from clasp output:`n$versionOutput"
}
$version = $match.Groups[1].Value
Write-Host "Parsed new version number: $version"

# 4. Repoint the existing deployment at the new version ----------------------
Step "4/4  clasp redeploy (deployment $DEPLOYMENT_ID -> v$version)"
clasp redeploy $DEPLOYMENT_ID -V $version -d $Message

Write-Host "`nDone. Deployment $DEPLOYMENT_ID now serves version $version." -ForegroundColor Green
