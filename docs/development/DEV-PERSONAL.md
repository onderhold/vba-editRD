# Personal Development Branch

Using a detached branch (orphan branch) for personal development automation is a smart strategy. Here's how to set it up and use it effectively:

## Creating an Orphan Branch for Personal Development

```powershell
# Create a new orphan branch (no commit history from main)
git checkout --orphan dev-personal

# Remove all files from staging (clean slate)
git rm -rf .

# Create a README explaining the branch purpose
New-Item -ItemType File -Path "README.md" -Value @"
# Personal Development Branch

This branch contains development automation scripts and configurations that are specific to your local workflow and should never be merged into the main branch or upstream.

## Purpose
- Local development scripts
- Personal build configurations
- Testing utilities
- Experimental automation tools

## Note
This branch is intentionally disconnected from the main development branches and should remain so.
"@

git add README.md
git commit -m "Initialize dev-personal branch"

# Push to your fork
git push -u origin dev-personal
```

## Using the Personal Development Branch

### 1. **Store automation files**

```powershell
# Switch to dev-personal branch
git checkout dev-personal

# Add your automation files
mkdir scripts
mkdir configs
mkdir tools

# Add files like:
# - Personal build scripts
# - Local configuration overrides
# - Testing helpers
# - Documentation for your workflow

git add .
git commit -m "Add personal automation scripts"
git push
```

### 2. **Access files from main branch**

You can easily grab files from the dev-personal branch while on another branch:

```powershell
# While on main or any feature branch
# Copy a specific file from dev-personal branch
git show dev-personal:scripts/my-script.ps1 > scripts/my-script.ps1

# Or checkout specific files
git checkout dev-personal -- scripts/my-script.ps1
```

### 3. **Use git worktree for simultaneous access**

This is especially useful if you frequently need both branches. Create the worktree as a subdirectory:

```powershell
# 1. Create the worktree as a subdirectory
git worktree add .dev-personal dev-personal

# 2. Add it to .gitignore
Add-Content .gitignore "`n# Personal development worktree`n.dev-personal/"

# 3. Commit the .gitignore change
git add .gitignore
git commit -m "Add .dev-personal worktree to gitignore"

# Now you have easy access to both:
# - c:\_daten.lokal\_workarea\git.repos\onderhold\vba-editRD (main worktree)
# - c:\_daten.lokal\_workarea\git.repos\onderhold\vba-editRD\.dev-personal (dev-personal worktree)

# Access personal scripts easily:
.\.dev-personal\scripts\local-build.ps1
```

**Benefits:**
- ✅ Scripts are directly accessible with relative paths
- ✅ No need for symbolic links
- ✅ IDE/editor can see both directories in the same project
- ✅ Easier to reference files between branches
- ✅ `.gitignore` ensures it never gets committed to main

## Visual Overview

Here's what the directory structure looks like with the worktree setup:

```
vba-editRD/                          # Main worktree (main branch or feature branches)
├── .gitignore                       # Contains: .dev-personal/
├── .git/                           # Shared git repository
├── src/
├── docs/
├── scripts/
└── .dev-personal/                  # Worktree for dev-personal branch
    ├── .gitignore                  # Different .gitignore (optional)
    ├── README.md
    ├── scripts/
    │   └── my-script.ps1          # Your personal scripts
    └── configs/
```

**How it works:**

- When you're in the **main worktree** and run `git status`, Git ignores `.dev-personal/` entirely (because of `.gitignore`)
- When you **`cd .dev-personal`** and run `git status`, you're in a different worktree checking out the `dev-personal` branch, and Git tracks everything normally
- Both worktrees share the same `.git` directory, so they're part of the same repository
- Each worktree has its own checked-out branch and independent working directory
- You can commit to `dev-personal` from inside `.dev-personal/` without it appearing in the main worktree

## Suggested Structure for Personal Development Branch

```
dev-personal/
├── README.md
├── scripts/
│   ├── local-build.ps1          # Your custom build scripts
│   ├── test-runner.ps1          # Local testing scripts
│   ├── sync-upstream.ps1        # Sync with upstream scripts
│   └── release-helper.ps1       # Your release workflow
├── configs/
│   ├── local-settings.json      # Local configuration overrides
│   ├── test-configs.toml        # Test configurations
│   └── .env.example             # Environment variables template
├── tools/
│   ├── utils.py                 # Helper utilities
│   └── validators.ps1           # Validation scripts
├── docs/
│   ├── MY-WORKFLOW.md           # Your personal workflow documentation
│   ├── NOTES.md                 # Development notes
│   └── IDEAS.md                 # Ideas for enhancements
└── .gitignore                   # Ignore secrets/credentials
```

## Example: `.gitignore` for dev-personal branch

```gitignore
# Secrets and credentials
*.token
*.key
.env
secrets/

# Local test data
test-data/
*.test.xlsx
*.test.docx

# Temporary files
temp/
*.tmp
```

## Benefits of This Approach

✅ **Never conflicts with upstream** - No risk of accidentally merging into upstream repo
✅ **Backed up on GitHub** - Your automation is safe and versioned
✅ **Easy to access** - Can pull files from dev-personal branch when needed
✅ **Clean separation** - Main development stays clean
✅ **Shareable** (optional) - Can share with collaborators who fork from you
✅ **No pull request pollution** - Won't show up in PR diffs

## Example Workflow

```powershell
# Working on a feature
git checkout -b feature/my-feature main

# Use scripts from dev-personal directly
.\.dev-personal\scripts\local-build.ps1

# Or copy a config temporarily
Copy-Item .\.dev-personal\configs\local-settings.json .\temp-settings.json

# Use it, then delete (or add to .gitignore)
Remove-Item .\temp-settings.json
```

## Protecting the Branch

Add a note to your fork's settings or create a `CONTRIBUTING.md` in the dev-personal branch:

```markdown
# Important: Branch Protection

**This branch should NEVER be merged into main or any other branch.**

This is an orphan branch for personal automation and tooling only.
```

## Workflow Integration

Once you have the worktree set up, you can directly reference files:

```powershell
# Run scripts from dev-personal directly
.\.dev-personal\scripts\local-build.ps1

# Or create aliases in PowerShell profile
# Add to $PROFILE:
function Run-PersonalBuild { .\.dev-personal\scripts\local-build.ps1 $args }
Set-Alias lpb Run-PersonalBuild

# Then just use:
lpb
```

## Maintenance Tips

### Keeping the branch updated with your changes

```powershell
# Navigate to the worktree
cd .dev-personal

# Make changes
# ... edit files ...

# Commit and push
git add .
git commit -m "Update personal automation scripts"
git push

# Return to main worktree
cd ..
```

### Cleaning up worktrees

```powershell
# List all worktrees
git worktree list

# Remove the worktree when no longer needed
git worktree remove .dev-personal

# Prune deleted worktrees
git worktree prune
```

## Use Cases

### 1. Personal release checklist
Store your pre-release verification steps that are specific to your setup.

### 2. Environment-specific configurations
Keep configurations that work on your machine but wouldn't be appropriate for the main repo.

### 3. Testing scripts
Experimental or work-in-progress testing utilities that aren't ready for the main branch.

### 4. Documentation
Personal notes, workflow documentation, and reminders that are just for you.

### 5. Integration scripts
Scripts that integrate with your other local tools or workflows.

## Understanding .gitignore with Worktrees

**Important clarification**: The `.gitignore` in each worktree is independent:

### In the main worktree:
```powershell
# Your .gitignore contains:
.dev-personal/

# So git status won't show the .dev-personal directory at all
cd c:\_daten.lokal\_workarea\git.repos\onderhold\vba-editRD
git status
# Output: nothing about .dev-personal (it's ignored)

# You cannot accidentally commit it
git add .
# .dev-personal is ignored
```

### In the .dev-personal worktree:
```powershell
# Switch to the worktree directory
cd .dev-personal

# This worktree checks out the dev-personal branch
# It has its own .gitignore (or none at all)
git status
# Output: shows changes in THIS worktree normally

# You can commit normally - the main branch's .gitignore doesn't affect this!
git add scripts/new-script.ps1
git commit -m "Add new personal script"
git push origin dev-personal
# Everything works normally!
```

**Key points:**
- ✅ Each worktree is independent with separate working directories
- ✅ The main branch's `.gitignore` only affects the main worktree
- ✅ The `.dev-personal/` entry prevents the main worktree from seeing the subdirectory as changes
- ✅ Inside `.dev-personal` you can commit normally to the `dev-personal` branch
- ✅ The two worktrees share the same `.git` directory but have different checked-out branches

## Security Considerations

⚠️ **Important**: This branch is still on GitHub, so:
- Don't commit actual secrets (API keys, passwords, tokens)
- Use `.env` files and add them to `.gitignore` (in the dev-personal branch)
- Store only templates and examples
- For real secrets, use a local-only solution (not in git)

## Summary

The dev-personal branch strategy allows you to:
- Version control your personal development tools
- Keep them separate from the main codebase
- Never worry about accidentally creating PRs with personal scripts
- Share your setup across machines
- Collaborate with your future self

This is especially useful for forked repositories where you want to maintain your own tooling without polluting the upstream project or your main branch.
