# Move to dev-personal - Usage Guide

This automation makes it easy to move files and directories from the main branch to your dev-personal branch.

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
- ✅ Preserves directory structure
- ✅ Confirms before proceeding
- ✅ Commits to both branches automatically
- ✅ Optional push to remote
- ✅ Safe error handling
- ✅ Clear progress messages

## Tips

- Use forward slashes or backslashes in paths (both work)
- Paths are relative to the repository root
- The script will ask for confirmation before moving
- You can choose to push immediately or later

## Troubleshooting

### "Source path does not exist"
- Check your path is correct and relative to repository root
- Make sure you're using the right path separator

### "Source is already in .dev-personal"
- The file/directory is already in dev-personal
- No need to move it again

### Git errors
- Make sure you don't have uncommitted changes
- Ensure you're on the correct branch
- Check your Git authentication if push fails

## Examples

```powershell
# Move release notes
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: docs/releases/v0.4.0a2-release-notes.md

# Move build directory
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: build/

# Move coverage HTML
Ctrl+Shift+P → Run Task → Move to dev-personal
Enter: htmlcov/
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
