# move-to-dev-personal.ps1
# Moves a file or directory from the main branch to the dev-personal branch

param(
    [Parameter(Mandatory=$true)]
    [string]$PathToMove
)

# Get repository root
$RepoRoot = git rev-parse --show-toplevel
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Not in a git repository!" -ForegroundColor Red
    exit 1
}

# Convert to Windows path if needed
$RepoRoot = $RepoRoot -replace '/', '\'

# Handle absolute paths by making them relative
if ([System.IO.Path]::IsPathRooted($PathToMove)) {
    $PathToMove = $PathToMove.Substring($RepoRoot.Length + 1)
}

# Normalize path separators
$PathToMove = $PathToMove -replace '/', '\'

$SourcePath = Join-Path $RepoRoot $PathToMove
$DevPersonalRoot = Join-Path $RepoRoot ".dev-personal"

Write-Host "`n=== Moving to dev-personal ===" -ForegroundColor Cyan
Write-Host "Repository root: $RepoRoot" -ForegroundColor Gray

# Validate source exists
if (-not (Test-Path $SourcePath)) {
    Write-Host "`nError: Source path does not exist: $SourcePath" -ForegroundColor Red
    exit 1
}

# Validate source is not in .dev-personal already
if ($SourcePath -like "$DevPersonalRoot*") {
    Write-Host "`nError: Source is already in .dev-personal!" -ForegroundColor Red
    exit 1
}

# Determine if it's a file or directory
$IsDirectory = Test-Path $SourcePath -PathType Container
$TargetPath = Join-Path $DevPersonalRoot $PathToMove

Write-Host "`nMoving: $PathToMove" -ForegroundColor Yellow
Write-Host "  From: $SourcePath" -ForegroundColor Gray
Write-Host "  To:   $TargetPath" -ForegroundColor Gray
Write-Host "  Type: $(if ($IsDirectory) { 'Directory' } else { 'File' })" -ForegroundColor Gray

# Confirm
$confirm = Read-Host "`nProceed with move? (y/N)"
if ($confirm -ne "y") {
    Write-Host "Cancelled." -ForegroundColor Gray
    exit 0
}

# Create parent directory in .dev-personal if needed
$TargetParent = Split-Path $TargetPath -Parent
if (-not (Test-Path $TargetParent)) {
    Write-Host "`nCreating directory: $TargetParent" -ForegroundColor Gray
    New-Item -ItemType Directory -Path $TargetParent -Force | Out-Null
}

# Copy to dev-personal
Write-Host "`nCopying to .dev-personal..." -ForegroundColor Yellow
try {
    if ($IsDirectory) {
        Copy-Item -Recurse -Path $SourcePath -Destination $TargetPath -Force
    } else {
        Copy-Item -Path $SourcePath -Destination $TargetPath -Force
    }
    Write-Host "✓ Copied successfully" -ForegroundColor Green
} catch {
    Write-Host "Error copying: $_" -ForegroundColor Red
    exit 1
}

# Commit in dev-personal
Write-Host "`nCommitting to dev-personal branch..." -ForegroundColor Yellow
Push-Location $DevPersonalRoot
try {
    git add $PathToMove
    if ($LASTEXITCODE -ne 0) { throw "git add failed" }
    
    git commit -m "Add $PathToMove from main"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Committed to dev-personal branch" -ForegroundColor Green
        
        # Ask to push
        $push = Read-Host "Push to origin/dev-personal? (y/N)"
        if ($push -eq "y") {
            git push
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Pushed to origin/dev-personal" -ForegroundColor Green
            } else {
                Write-Host "Warning: Push failed!" -ForegroundColor Yellow
            }
        }
    } else {
        throw "git commit failed"
    }
} catch {
    Write-Host "Error in dev-personal commit: $_" -ForegroundColor Red
    Pop-Location
    exit 1
}
Pop-Location

# Remove from main
Write-Host "`nRemoving from main branch..." -ForegroundColor Yellow
Push-Location $RepoRoot
try {
    git rm -rf $PathToMove
    if ($LASTEXITCODE -ne 0) { throw "git rm failed" }
    
    git commit -m "Remove $PathToMove (moved to dev-personal)"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Removed from main branch" -ForegroundColor Green
        
        # Ask to push
        $push = Read-Host "Push to origin/main? (y/N)"
        if ($push -eq "y") {
            git push
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Pushed to origin/main" -ForegroundColor Green
            } else {
                Write-Host "Warning: Push failed!" -ForegroundColor Yellow
            }
        }
    } else {
        throw "git commit failed"
    }
} catch {
    Write-Host "Error in main commit: $_" -ForegroundColor Red
    Pop-Location
    exit 1
}
Pop-Location

Write-Host "`n✓ Move complete!" -ForegroundColor Green
Write-Host "The $(if ($IsDirectory) { 'directory' } else { 'file' }) is now in .dev-personal\$PathToMove" -ForegroundColor Cyan
Write-Host "`nYou can access it at: .\.dev-personal\$PathToMove" -ForegroundColor Gray
