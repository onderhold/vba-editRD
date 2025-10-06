# scripts/release.ps1
<# 
# Normal usage - will prompt if release exists
.\scripts\release.ps1

# Force recreate without prompting
.\scripts\release.ps1 -Force

# Use with custom notes file
.\scripts\release.ps1 -NotesFile "docs/releases/custom.md"

# Force recreate with custom notes
.\scripts\release.ps1 -Force -NotesFile "docs/releases/custom.md"
#>

param(
    [string]$NotesFile = "",
    [string]$Repo = "onderhold/vba-editRD",
    [string]$RequiredUser = "onderhold",  # Set your required GitHub username
    [switch]$Force  # Force recreate release if it exists
)

# Find the repository root by looking for .git directory
function Get-RepositoryRoot {
    $currentPath = $PSScriptRoot
    while ($currentPath) {
        if (Test-Path (Join-Path $currentPath ".git")) {
            return $currentPath
        }
        $parent = Split-Path $currentPath -Parent
        if ($parent -eq $currentPath) {
            # Reached root without finding .git
            break
        }
        $currentPath = $parent
    }
    Write-Host "Error: Could not find repository root (.git directory)!" -ForegroundColor Red
    exit 1
}

$RepoRoot = Get-RepositoryRoot
Write-Host "Repository root: $RepoRoot" -ForegroundColor Cyan

# Verify GitHub authentication and user
Write-Host "Checking GitHub authentication..." -ForegroundColor Yellow
try {
    $currentUser = gh api user --jq '.login' 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Not authenticated with GitHub CLI!" -ForegroundColor Red
        Write-Host "Please run: gh auth login" -ForegroundColor Yellow
        exit 1
    }
    
    Write-Host "Currently authenticated as: $currentUser" -ForegroundColor Cyan
    
    if ($currentUser -ne $RequiredUser) {
        Write-Host "Error: Wrong GitHub user!" -ForegroundColor Red
        Write-Host "  Expected: $RequiredUser" -ForegroundColor Yellow
        Write-Host "  Current:  $currentUser" -ForegroundColor Yellow
        Write-Host "" 
        Write-Host "Please authenticate with the correct account:" -ForegroundColor Yellow
        Write-Host "  gh auth logout" -ForegroundColor Gray
        Write-Host "  gh auth login" -ForegroundColor Gray
        exit 1
    }
    
    Write-Host "✓ Authenticated as correct user: $RequiredUser" -ForegroundColor Green
} catch {
    Write-Host "Error checking GitHub authentication: $_" -ForegroundColor Red
    exit 1
}

# Read version from pyproject.toml
Write-Host "Reading version from pyproject.toml..." -ForegroundColor Yellow
$pyprojectPath = Join-Path $RepoRoot "pyproject.toml"
$pyprojectContent = Get-Content $pyprojectPath -Raw

if ($pyprojectContent -match 'version\s*=\s*"([^"]+)"') {
    $VersionNumber = $Matches[1]
    $Version = "v$VersionNumber"
    Write-Host "Found version: $Version" -ForegroundColor Green
} else {
    Write-Host "Error: Could not read version from pyproject.toml!" -ForegroundColor Red
    exit 1
}

# Set default notes file if not provided
if ([string]::IsNullOrEmpty($NotesFile)) {
    $NotesFile = Join-Path $RepoRoot "docs\releases\$Version-release-notes.md"
}

# Check if release already exists
Write-Host "Checking if release $Version already exists..." -ForegroundColor Yellow
$releaseExists = $false
gh release view $Version --repo $Repo 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    $releaseExists = $true
    Write-Host "⚠️  Release $Version already exists!" -ForegroundColor Yellow
    
    if ($Force) {
        Write-Host "Force flag detected. Deleting existing release..." -ForegroundColor Yellow
        gh release delete $Version --repo $Repo --yes
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Existing release deleted successfully" -ForegroundColor Green
            $releaseExists = $false
        } else {
            Write-Host "Error: Failed to delete existing release!" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host ""
        Write-Host "Options:" -ForegroundColor Cyan
        Write-Host "  1. Delete and recreate the release" -ForegroundColor White
        Write-Host "  2. Upload files to existing release" -ForegroundColor White
        Write-Host "  3. Cancel" -ForegroundColor White
        Write-Host ""
        $choice = Read-Host "Enter your choice (1-3)"
        
        switch ($choice) {
            "1" {
                Write-Host "Deleting existing release..." -ForegroundColor Yellow
                gh release delete $Version --repo $Repo --yes
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✓ Existing release deleted successfully" -ForegroundColor Green
                    $releaseExists = $false
                } else {
                    Write-Host "Error: Failed to delete existing release!" -ForegroundColor Red
                    exit 1
                }
            }
            "2" {
                Write-Host "Will upload files to existing release..." -ForegroundColor Cyan
            }
            "3" {
                Write-Host "Operation cancelled by user." -ForegroundColor Gray
                exit 0
            }
            default {
                Write-Host "Invalid choice. Exiting." -ForegroundColor Red
                exit 1
            }
        }
    }
} else {
    Write-Host "✓ Release $Version does not exist yet" -ForegroundColor Green
}

# Clean old dev builds from dist folder
Write-Host "Cleaning old dev builds from dist folder..." -ForegroundColor Yellow
$distPath = Join-Path $RepoRoot "dist"
if (Test-Path $distPath) {
    $oldDevFiles = Get-ChildItem -Path $distPath -Filter "*.dev*" | Where-Object { 
        $_.Name -like "*.whl" -or $_.Name -like "*.tar.gz" 
    }
    
    foreach ($file in $oldDevFiles) {
        # Only delete if it doesn't match the current version
        if ($file.Name -notlike "*$VersionNumber*") {
            Write-Host "  Deleting: $($file.Name)" -ForegroundColor Gray
            Remove-Item $file.FullName -Force
        }
    }
    Write-Host "Cleanup complete!" -ForegroundColor Green
} else {
    Write-Host "dist directory not found, skipping cleanup." -ForegroundColor Gray
}

Write-Host "Creating release $Version for repository $Repo..." -ForegroundColor Green

# Create the release if it doesn't exist
if (-not $releaseExists) {
    Write-Host "Creating release $Version for repository $Repo..." -ForegroundColor Green
    
    gh release create $Version `
      --repo $Repo `
      --title "VBA Edit $Version" `
      --notes-file $NotesFile `
      --prerelease
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error creating release!" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Release created successfully!" -ForegroundColor Green
}

# Upload files - only for the specific version
Write-Host "Uploading files for version $VersionNumber..." -ForegroundColor Yellow
$distWildcard = Join-Path $RepoRoot "dist"
gh release upload $Version `
  --repo $Repo `
  --clobber `
  "$distWildcard\*-vba.exe" "$distWildcard\vba_edit-$VersionNumber*.whl" "$distWildcard\vba_edit-$VersionNumber*.tar.gz"

if ($LASTEXITCODE -eq 0) {
    Write-Host "All files uploaded successfully!" -ForegroundColor Green
    Write-Host "View your release at: https://github.com/$Repo/releases/tag/$Version" -ForegroundColor Cyan
} else {
    Write-Host "Error uploading files!" -ForegroundColor Red
    exit 1
}