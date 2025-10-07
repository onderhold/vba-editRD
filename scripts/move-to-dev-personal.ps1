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

$DevPersonalRoot = Join-Path $RepoRoot ".dev-personal"

Write-Host "`n=== Moving to dev-personal ===" -ForegroundColor Cyan
Write-Host "Repository root: $RepoRoot" -ForegroundColor Gray

# Resolve wildcards - get matching items
$SearchPath = Join-Path $RepoRoot $PathToMove
$MatchingItems = Get-ChildItem -Path $SearchPath -ErrorAction SilentlyContinue

if (-not $MatchingItems) {
    Write-Host "`nError: No files or directories match: $PathToMove" -ForegroundColor Red
    exit 1
}

# Filter out items already in .dev-personal
$ItemsToMove = $MatchingItems | Where-Object { $_.FullName -notlike "$DevPersonalRoot*" }

if (-not $ItemsToMove) {
    Write-Host "`nError: All matching items are already in .dev-personal!" -ForegroundColor Red
    exit 1
}

# Show what will be moved
Write-Host "`nFound $($ItemsToMove.Count) item(s) to move:" -ForegroundColor Yellow
foreach ($item in $ItemsToMove) {
    $relativePath = $item.FullName.Substring($RepoRoot.Length + 1)
    $itemType = if ($item.PSIsContainer) { "Directory" } else { "File" }
    Write-Host "  - $relativePath [$itemType]" -ForegroundColor Gray
}

# Flush any pending input (like venv activation commands)
if ($Host.UI.RawUI.KeyAvailable) {
    while ($Host.UI.RawUI.KeyAvailable) {
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    }
}

# Small delay to let any background processes finish
Start-Sleep -Milliseconds 500

# Confirm
$confirm = Read-Host "`nProceed with move? (y/N)"
if ($confirm -ne "y") {
    Write-Host "Cancelled." -ForegroundColor Gray
    exit 0
}

# Process each item
$MovedPaths = @()
$FailedItems = @()

Write-Host "`nCopying to .dev-personal..." -ForegroundColor Yellow
foreach ($item in $ItemsToMove) {
    $relativePath = $item.FullName.Substring($RepoRoot.Length + 1)
    $targetPath = Join-Path $DevPersonalRoot $relativePath
    
    try {
        # Create parent directory if needed
        $targetParent = Split-Path $targetPath -Parent
        if (-not (Test-Path $targetParent)) {
            New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
        }
        
        # Copy the item
        if ($item.PSIsContainer) {
            Copy-Item -Recurse -Path $item.FullName -Destination $targetPath -Force
        } else {
            Copy-Item -Path $item.FullName -Destination $targetPath -Force
        }
        
        Write-Host "  ✓ Copied: $relativePath" -ForegroundColor Green
        $MovedPaths += $relativePath
    } catch {
        Write-Host "  ✗ Failed: $relativePath - $_" -ForegroundColor Red
        $FailedItems += $relativePath
    }
}

if ($FailedItems.Count -gt 0) {
    Write-Host "`nWarning: $($FailedItems.Count) item(s) failed to copy" -ForegroundColor Yellow
    $continue = Read-Host "Continue with remaining items? (y/N)"
    if ($continue -ne "y") {
        Write-Host "Aborted." -ForegroundColor Gray
        exit 1
    }
}

if ($MovedPaths.Count -eq 0) {
    Write-Host "`nError: No items were successfully copied" -ForegroundColor Red
    exit 1
}

# Commit in dev-personal
Write-Host "`nCommitting to dev-personal branch..." -ForegroundColor Yellow
Push-Location $DevPersonalRoot
try {
    foreach ($path in $MovedPaths) {
        git add $path
        if ($LASTEXITCODE -ne 0) { throw "git add failed for $path" }
    }
    
    $commitMessage = if ($MovedPaths.Count -eq 1) {
        "Add $($MovedPaths[0]) from main"
    } else {
        "Add $($MovedPaths.Count) items from main: $($MovedPaths -join ', ')"
    }
    
    git commit -m $commitMessage
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
    foreach ($path in $MovedPaths) {
        git rm -rf $path
        if ($LASTEXITCODE -ne 0) { throw "git rm failed for $path" }
    }
    
    $commitMessage = if ($MovedPaths.Count -eq 1) {
        "Remove $($MovedPaths[0]) (moved to dev-personal)"
    } else {
        "Remove $($MovedPaths.Count) items (moved to dev-personal): $($MovedPaths -join ', ')"
    }
    
    git commit -m $commitMessage
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

Write-Host "`n=== Move Complete ===" -ForegroundColor Green
Write-Host "Successfully moved $($MovedPaths.Count) item(s) to .dev-personal:" -ForegroundColor Cyan
foreach ($path in $MovedPaths) {
    Write-Host "  - $path" -ForegroundColor Gray
}
Write-Host "`nAccess them at: .\.dev-personal\" -ForegroundColor Gray
Write-Host "`n✓ Task completed successfully. Press any key or close this terminal." -ForegroundColor Green
Write-Host "==========================================`n" -ForegroundColor Green
