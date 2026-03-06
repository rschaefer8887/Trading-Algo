# Link this folder to your GitHub repo. Run in PowerShell from this folder.
# Usage: .\Link_GitHub.ps1 [GitHubRepoURL]
# Example: .\Link_GitHub.ps1 https://github.com/YourUsername/Trading-Algo.git

param(
    [Parameter(Position = 0)]
    [string]$GitHubUrl = ""
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not $GitHubUrl) {
    $GitHubUrl = Read-Host "Paste your GitHub repo URL (e.g. https://github.com/You/Trading-Algo.git)"
}
$GitHubUrl = $GitHubUrl.Trim()
if (-not $GitHubUrl) {
    Write-Error "GitHub URL is required."
}

# Require git on PATH (e.g. run from Git Bash or a terminal where Git is installed)
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "Git not found. Install Git and add it to PATH, or run the commands in this script manually from Git Bash."
}

if (-not (Test-Path ".git")) {
    git init
    Write-Host "Initialized Git repo."
} else {
    Write-Host "Git repo already initialized."
}

# Remove existing 'origin' only if it exists (ignore error on first run)
$prevErrorAction = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
$null = git remote get-url origin 2>$null
$originExists = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = $prevErrorAction
if ($originExists) { git remote remove origin }
git remote add origin $GitHubUrl
Write-Host "Remote 'origin' set to $GitHubUrl"

git add .
git commit -m "Initial commit"
git branch -M main
Write-Host "Created initial commit on branch main."

Write-Host "Pushing to GitHub..."
git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "Remote may have existing commits (e.g. README). Pulling and merging..."
    git pull origin main --allow-unrelated-histories
    git push -u origin main
}
Write-Host "Done. Your project is linked to GitHub."
