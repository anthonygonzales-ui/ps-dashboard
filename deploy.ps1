<#
    deploy.ps1 — push local changes to GitHub.

    Deployment to Apps Script is handled automatically by GitHub Actions
    (.github/workflows/deploy.yml) on every push to main: it runs clasp push,
    creates a new version, and repoints the deployment ID. So this script only
    needs to commit and push — CI does the rest.

    Usage:
      ./deploy.ps1 "Your commit message"
      ./deploy.ps1                      # uses a default message
#>
param(
    [string]$Message = "Update PS dashboard"
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

git add -A
$pending = git status --porcelain
if ($pending) {
    git commit -m $Message
}
git push origin main

Write-Host "`nPushed to GitHub. GitHub Actions will deploy to Apps Script automatically." -ForegroundColor Green
Write-Host "Watch progress: https://github.com/anthonygonzales-ui/ps-dashboard/actions" -ForegroundColor Cyan
