# Release Checklist for vba-edit

## Pre-Release
- [ ] Update version in [project] section of pyproject.toml
- [ ] Run `python build_release.py --update-version --check-only` 
- [ ] Update CHANGELOG.md
- [ ] Commit version changes
- [ ] Run full test suite: `pytest tests/ -v --skip-vba-check --cov-report=html`

## Build Release
- [ ] Run `python build_release.py --create-tag` (builds AND creates tag)
- [ ] Test all executables: `excel-vba.exe --version`, etc.
- [ ] Test basic functionality of each executable

## Release (Automated)
- [ ] Tag created automatically: `python build_release.py --tag-only` (if not done above)
- [ ] Push tag: Prompted during tag creation
- [ ] Create GitHub release with executables
- [ ] Update documentation if needed

## Manual Git Commands (Alternative)
```powershell
# Create tag from pyproject.toml version
git tag "v$(python -c "import tomllib; print(tomllib.load(open('pyproject.toml', 'rb'))['project']['version'])")"

# Push the tag
git push origin "v$(python -c "import tomllib; print(tomllib.load(open('pyproject.toml', 'rb'))['project']['version'])")"
```

## Post-Release  
- [ ] Test download links
- [ ] Update any external documentation
- [ ] Plan next version features