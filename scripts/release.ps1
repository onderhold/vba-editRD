# scripts/release.ps1
param(
    [string]$NotesFile = "",
    [string]$Repo = "onderhold/vba-editRD"
)

# Read version from pyproject.toml
Write-Host "Reading version from pyproject.toml..." -ForegroundColor Yellow
$pyprojectPath = Join-Path $PSScriptRoot ".." "pyproject.toml"
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
    $NotesFile = "docs/releases/$Version-release-notes.md"
}

# Clean old dev builds from dist folder
Write-Host "Cleaning old dev builds from dist folder..." -ForegroundColor Yellow
$distPath = Join-Path $PSScriptRoot ".." "dist"
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

# Create the release
gh release create $Version `
  --repo $Repo `
  --title "VBA Edit $Version" `
  --notes-file $NotesFile `
  --prerelease

if ($LASTEXITCODE -eq 0) {
    Write-Host "Release created successfully!" -ForegroundColor Green
    
    # Upload files - only for the specific version
    Write-Host "Uploading files for version $VersionNumber..." -ForegroundColor Yellow
    gh release upload $Version `
      --repo $Repo `
      dist/*-vba.exe dist/vba_edit-$VersionNumber*.whl dist/vba_edit-$VersionNumber*.tar.gz
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "All files uploaded successfully!" -ForegroundColor Green
        Write-Host "View your release at: https://github.com/$Repo/releases/tag/$Version" -ForegroundColor Cyan
    } else {
        Write-Host "Error uploading files!" -ForegroundColor Red
    }
} else {
    Write-Host "Error creating release!" -ForegroundColor Red
}