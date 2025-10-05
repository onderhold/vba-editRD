"""
Automated build script for vba-edit executables.
Ensures all necessary steps are completed before building.

This script will:

1. ✅ Check version consistency between pyproject.toml and __init__.py
2. 🔄 Optionally update the fallback version automatically
3. 📦 Reinstall the package with current metadata
4. 🔨 Build executables using create_binaries.py
5. 🧪 Test the built executables
6. 🏷️ Optionally create git tags automatically from the version

## Usage

```python
# Check version consistency only
python build_release.py --check-only

# Update fallback version automatically and check
python build_release.py --update-version --check-only

# Full build process
python build_release.py

# Build specific apps
python build_release.py --apps excel word

# Full release process (build + tag)
python build_release.py --create-tag

# Only create/update git tag (no building)
python build_release.py --tag-only

# Skip package reinstall (if you just did it)
python build_release.py --skip-install
```

"""

import argparse
import subprocess
import sys
from pathlib import Path


def run_command(cmd, description, check=True):
    """Run a command and handle errors."""
    print(f"\n🔄 {description}")
    print(f"   Command: {' '.join(cmd) if isinstance(cmd, list) else cmd}")

    try:
        result = subprocess.run(cmd, shell=isinstance(cmd, str), check=check, capture_output=True, text=True)
        if result.stdout:
            print(f"   Output: {result.stdout.strip()}")
        return result
    except subprocess.CalledProcessError as e:
        print(f"   ❌ Failed: {e}")
        if e.stdout:
            print(f"   Stdout: {e.stdout}")
        if e.stderr:
            print(f"   Stderr: {e.stderr}")
        if check:
            sys.exit(1)
        return e


def get_version_from_pyproject():
    """Read version from pyproject.toml."""
    pyproject_path = Path("pyproject.toml")
    if not pyproject_path.exists():
        print("   ❌ pyproject.toml not found!")
        sys.exit(1)

    try:
        # Try Python 3.11+ stdlib tomllib first
        import tomllib

        with open(pyproject_path, "rb") as f:
            data = tomllib.load(f)
    except ImportError:
        # Fallback to tomli for older Python versions
        try:
            import tomli

            with open(pyproject_path, "rb") as f:
                data = tomli.load(f)
        except ImportError:
            # Final fallback: parse manually
            content = pyproject_path.read_text()
            import re

            match = re.search(r'version = "([^"]+)"', content)
            if not match:
                print("   ❌ Version not found in pyproject.toml!")
                sys.exit(1)
            return match.group(1)

    return data["project"]["version"]


def check_version_consistency():
    """Check that versions are consistent across files."""
    print("\n🔍 Checking version consistency...")

    # Read version from pyproject.toml
    pyproject_version = get_version_from_pyproject()
    print(f"   pyproject.toml version: {pyproject_version}")

    # Read fallback version from __init__.py
    init_path = Path("src/vba_edit/__init__.py")
    if not init_path.exists():
        print("   ❌ __init__.py not found!")
        sys.exit(1)

    init_content = init_path.read_text()
    import re

    init_match = re.search(r'__version__ = "([^"]+)".*# Keep this in sync', init_content)
    if not init_match:
        print("   ❌ Fallback version not found in __init__.py!")
        sys.exit(1)

    init_version = init_match.group(1)
    print(f"   __init__.py fallback: {init_version}")

    if pyproject_version != init_version:
        print(f"   ❌ Version mismatch! Update __init__.py fallback to {pyproject_version}")
        sys.exit(1)
    else:
        print(f"   ✅ Versions are consistent: {pyproject_version}")


def update_fallback_version():
    """Update the fallback version in __init__.py to match pyproject.toml."""
    print("\n🔄 Updating fallback version...")

    # Read version from pyproject.toml
    current_version = get_version_from_pyproject()

    # Update __init__.py
    init_path = Path("src/vba_edit/__init__.py")
    init_content = init_path.read_text()

    # Replace the fallback version
    import re

    updated_content = re.sub(
        r'(__version__ = ")[^"]+(.*# Keep this in sync.*)', f"\\g<1>{current_version}\\g<2>", init_content
    )

    if updated_content != init_content:
        init_path.write_text(updated_content)
        print(f"   ✅ Updated fallback version to {current_version}")
    else:
        print(f"   ✅ Fallback version already current: {current_version}")


def create_git_tag():
    """Create a git tag for the current version."""
    version = get_version_from_pyproject()
    tag_name = f"v{version}"

    print(f"\n🏷️  Creating git tag: {tag_name}")

    # Check if tag already exists
    result = run_command(["git", "tag", "-l", tag_name], "Checking if tag exists", check=False)
    if result.stdout.strip():
        print(f"   ⚠️  Tag {tag_name} already exists!")
        response = input("   Do you want to delete and recreate it? (y/N): ")
        if response.lower() == "y":
            run_command(["git", "tag", "-d", tag_name], f"Deleting existing tag {tag_name}")
        else:
            print("   Skipping tag creation.")
            return tag_name

    # Create the tag
    run_command(["git", "tag", tag_name], f"Creating tag {tag_name}")
    print(f"   ✅ Created tag: {tag_name}")

    # Ask if user wants to push the tag
    response = input(f"   Do you want to push tag {tag_name} to origin? (y/N): ")
    if response.lower() == "y":
        run_command(["git", "push", "origin", tag_name], f"Pushing tag {tag_name}")
        print(f"   ✅ Pushed tag: {tag_name}")

    return tag_name


def check_pyinstaller():
    """Check if PyInstaller is installed."""
    print("\n🔍 Checking PyInstaller availability...")
    try:
        import PyInstaller

        print(f"   ✅ PyInstaller found: {PyInstaller.__version__}")
    except ImportError:
        print("   ❌ PyInstaller not installed!")
        print("   Install with: pip install pyinstaller")
        print("   Or install all dev dependencies: pip install -e .[dev]")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Build vba-edit executables with all prerequisites")
    parser.add_argument("--apps", nargs="*", help="Specific apps to build (default: all)")
    parser.add_argument("--skip-install", action="store_true", help="Skip package reinstallation")
    parser.add_argument("--update-version", action="store_true", help="Update fallback version automatically")
    parser.add_argument("--check-only", action="store_true", help="Only check versions, don't build")
    parser.add_argument("--create-tag", action="store_true", help="Create git tag after successful build")
    parser.add_argument("--tag-only", action="store_true", help="Only create git tag, don't build")

    args = parser.parse_args()

    print("🚀 vba-edit Release Build Process")
    print("=" * 50)

    if args.tag_only:
        create_git_tag()
        return

    # Step 1: Check/Update version consistency
    if args.update_version:
        update_fallback_version()
    else:
        check_version_consistency()

    if args.check_only:
        print("\n✅ Version check complete!")
        return

    # Step 2: Reinstall package (unless skipped)
    if not args.skip_install:
        print("\n🔄 Reinstalling package...")
        run_command(["pip", "uninstall", "vba_edit", "vba-edit", "-y"], "Uninstalling old package", check=False)
        run_command(["pip", "install", "-e", ".[dev]"], "Installing package in development mode")

    # Step 3: Verify installation
    print("\n🔍 Verifying installation...")
    result = run_command(
        [sys.executable, "-c", "from vba_edit import __version__; print(__version__)"], "Checking installed version"
    )
    installed_version = result.stdout.strip()
    print(f"   ✅ Installed version: {installed_version}")

    # Step 4: Check PyInstaller availability
    check_pyinstaller()

    # Step 5: Build executables
    print("\n🔨 Building executables...")
    build_args = [sys.executable, "create_binaries.py"]
    if args.apps:
        build_args.extend(["--apps"] + args.apps)

    run_command(build_args, "Building executables")

    # Step 6: Test executables
    print("\n🧪 Testing executables...")
    dist_path = Path("dist")
    if dist_path.exists():
        for exe in dist_path.glob("*-vba.exe"):
            run_command([str(exe), "--version"], f"Testing {exe.name}", check=False)

    print("\n🎉 Build process completed successfully!")

    # Step 7: Optionally create tag
    if args.create_tag:
        create_git_tag()

    print("\nNext steps:")
    print("- Test the executables thoroughly")
    print("- Update CHANGELOG.md if needed")
    print("- Create a git tag: python build_release.py --tag-only")
    print("- Upload executables to GitHub releases")


if __name__ == "__main__":
    main()
