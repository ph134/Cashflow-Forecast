param(
  [string]$FromBranch = "main",
  [string]$LiveBranch = "live-share",
  [string]$Remote = "origin",
  [ValidateSet("major","minor","patch")]
  [string]$Bump = "patch"
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

# --- Sync cashflow.html -> index.html ---
$repoRoot = git rev-parse --show-toplevel
$cashflowFile = Join-Path $repoRoot "cashflow.html"
$indexFile = Join-Path $repoRoot "index.html"
if (Test-Path $cashflowFile) {
  Copy-Item $cashflowFile $indexFile -Force
  Write-Host "Synced cashflow.html -> index.html" -ForegroundColor Green
}

# --- Auto-increment version in index.html ---
$content = Get-Content $indexFile -Raw
if ($content -match 'class="version-tag">v(\d+)\.(\d+)\.(\d+)<') {
  [int]$major = $Matches[1]
  [int]$minor = $Matches[2]
  [int]$patch = $Matches[3]
  $old = "v$major.$minor.$patch"
  switch ($Bump) {
    "major" { $major++; $minor = 0; $patch = 0 }
    "minor" { $minor++; $patch = 0 }
    "patch" { $patch++ }
  }
  $new = "v$major.$minor.$patch"
  $content = $content -replace [regex]::Escape("version-tag"">$old<"), "version-tag"">$new<"
  Set-Content $indexFile $content -NoNewline
  Write-Host "Version bumped: $old -> $new" -ForegroundColor Yellow
  git add $indexFile $cashflowFile
  git commit -m "$new"
} else {
  Write-Host "No version tag found in index.html, skipping bump." -ForegroundColor DarkYellow
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
