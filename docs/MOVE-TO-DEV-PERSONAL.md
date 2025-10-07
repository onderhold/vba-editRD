# Move to dev-personal - Usage Guide

This is a VS Code automation that makes it easy to move files and directories from the main branch to your dev-personal branch.

## How to Use

### Method 1: VS Code Task (Recommended)

1. Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac)
2. Type "Tasks: Run Task"
3. Select "Move to dev-personal"
4. Enter the path to move (e.g., `docs/myfile.md` or `build/`)
5. Confirm the move
6. Answer whether to push to remote (y/N)

**Examples:**
- Move a single file: `docs/releases/v0.4.0a2-release-notes.md`
- Move a directory: `build/`
- Move multiple nested items: `htmlcov/`
- Use wildcards: `*.spec` or `docs/releases/*.md`
- Move all test files: `tests/__pycache__/*`

### Method 2: Run Script Directly

```powershell
# From repository root
.\.dev-personal\scripts\move-to-dev-personal.ps1 "path/to/file-or-folder"
```

## What It Does

1. **Validates** the source exists and is not already in dev-personal
2. **Copies** the file/directory to `.dev-personal/`
3. **Commits** to the dev-personal branch
4. **Removes** from the main branch
5. **Commits** the removal to main
6. **Optionally pushes** to remote repositories

## Features

- ✅ Works with files and directories
- ✅ Supports wildcards (e.g., `*.spec`, `docs/*.md`)
- ✅ Can move multiple items at once
- ✅ Preserves directory structure
- ✅ Confirms before proceeding
- ✅ Commits to both branches automatically
- ✅ Optional push to remote
- ✅ Safe error handling
- ✅ Clear progress messages
- ✅ Skips items already in dev-personal

## Tips

- Use forward slashes or backslashes in paths (both work)
- Paths are relative to the repository root
- Supports PowerShell wildcards: `*`, `?`, `[abc]`, etc.
- The script will show all matching items before moving
- The script will ask for confirmation before moving
- You can choose to push immediately or later
- Items already in dev-personal are automatically skipped

## Wildcard Examples

```powershell
# Move all .spec files
*.spec

# Move all release notes
docs/releases/*.md

# Move all Python cache directories
**/__pycache__

# Move all files starting with "test"
test*

# Move all .pyc files in tests
tests/**/*.pyc

# Move all build artifacts
build/* dist/*
```

## Troubleshooting

### "No files or directories match"
- Check your wildcard pattern is correct
- Make sure files exist in the main branch
- Verify you're using the correct path separator

### "Source path does not exist"
- Check your path is correct and relative to repository root
- For single files, make sure the filename is exact

### "All matching items are already in .dev-personal"
- The files/directories are already in dev-personal
- No need to move them again

### Git errors
- Make sure you don't have uncommitted changes
- Ensure you're on the correct branch
- Check your Git authentication if push fails

## Examples

```powershell
# Move single file
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: docs/releases/v0.4.0a2-release-notes.md

# Move directory
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: build/

# Move coverage HTML
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: htmlcov/

# Move all .spec files using wildcard
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: *.spec

# Move all release notes
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: docs/releases/*.md

# Move all __pycache__ directories
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: **/__pycache__
```

## After Moving

Access your moved files at:
```powershell
.\.dev-personal\your-file-or-folder
```

Run scripts from dev-personal:
```powershell
.\.dev-personal\scripts\release.ps1
```
