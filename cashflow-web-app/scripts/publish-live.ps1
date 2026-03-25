param(
  [string]$FromBranch = "main",
  [string]$LiveBranch = "live-share",
  [string]$Remote = "origin"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Error "Git is not installed or not on PATH. Install Git and retry."
}

Write-Host "Checking repository status..." -ForegroundColor Cyan
$insideRepo = git rev-parse --is-inside-work-tree 2>$null
if ($insideRepo -ne "true") {
  throw "Current folder is not a git repository."
}

$dirty = git status --porcelain
if ($dirty) {
  throw "Working tree has uncommitted changes. Commit or stash them first."
}

Write-Host "Fetching latest from $Remote..." -ForegroundColor Cyan
git fetch $Remote --prune

Write-Host "Updating source branch '$FromBranch'..." -ForegroundColor Cyan
git checkout $FromBranch
git pull $Remote $FromBranch

Write-Host "Updating live branch '$LiveBranch'..." -ForegroundColor Cyan
$branchExists = (git branch --list $LiveBranch) -ne ""
if (-not $branchExists) {
  git checkout -b $LiveBranch $FromBranch
} else {
  git checkout $LiveBranch
  git merge --ff-only $FromBranch
}

Write-Host "Pushing '$LiveBranch' to $Remote..." -ForegroundColor Cyan
git push -u $Remote $LiveBranch

Write-Host "Done. Share your GitHub Pages link after enabling Pages for branch '$LiveBranch'." -ForegroundColor Green
