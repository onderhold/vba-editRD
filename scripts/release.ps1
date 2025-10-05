# scripts/release.ps1
param(
    [string]$Version = "v0.4.1a1",
    [string]$NotesFile = "docs/releases/v0.4.1a1-release-notes.md",
    [string]$Repo = "onderhold/vba-editRD"
)

Write-Host "Creating release $Version for repository $Repo..." -ForegroundColor Green

# Create the release
gh release create $Version `
  --repo $Repo `
  --title "VBA Edit $Version" `
  --notes-file $NotesFile `
  --prerelease

if ($LASTEXITCODE -eq 0) {
    Write-Host "Release created successfully!" -ForegroundColor Green
    
    # Upload files
    Write-Host "Uploading files..." -ForegroundColor Yellow
    gh release upload $Version `
      --repo $Repo `
      dist/*.exe dist/*.whl dist/*.tar.gz
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "All files uploaded successfully!" -ForegroundColor Green
        Write-Host "View your release at: https://github.com/$Repo/releases/tag/$Version" -ForegroundColor Cyan
    } else {
        Write-Host "Error uploading files!" -ForegroundColor Red
    }
} else {
    Write-Host "Error creating release!" -ForegroundColor Red
}