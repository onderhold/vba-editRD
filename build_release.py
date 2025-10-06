"""
Automated build script for vba-edit executables.
Ensures all necessary steps are completed before building.

This script will:

1. ✅ Check version consistency between pyproject.toml and __init__.py
2. 🔄 Optionally update the fallback version automatically
3. 📦 Reinstall the package with current metadata
4. 🔨 Build executables using create_binaries.py
5. 📦 Build Python package distributions (.whl and .tar.gz)
6. 🧪 Test the built executables
7. 🏷️ Optionally create git tags automatically from the version

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

# Skip building Python packages (.whl/.tar.gz)
python build_release.py --skip-package-build
```

"""

import argparse
import re
import subprocess
import sys
import tomllib
from pathlib import Path

# Require Python 3.11+ for tomllib support
if sys.version_info < (3, 11):
    print("❌ Error: This script requires Python 3.11 or later (for tomllib support)")
    print(f"   Current version: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    print("   Please upgrade Python or use an older version of this script")
    sys.exit(1)


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

    with open(pyproject_path, "rb") as f:
        data = tomllib.load(f)

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


def check_build_package():
    """Check if build package is installed."""
    print("\n🔍 Checking build package availability...")
    try:
        import build

        print("   ✅ build package found")
    except ImportError:
        print("   ❌ build package not installed!")
        print("   Install with: pip install build")
        print("   Or install all dev dependencies: pip install -e .[dev]")
        sys.exit(1)


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


def build_python_packages():
    """Build Python package distributions (.whl and .tar.gz)."""
    print("\n📦 Building Python package distributions...")
    run_command([sys.executable, "-m", "build"], "Building .whl and .tar.gz packages")

    # List what was created
    dist_path = Path("dist")
    if dist_path.exists():
        version = get_version_from_pyproject()
        whl_files = list(dist_path.glob(f"vba_edit-{version}*.whl"))
        tar_files = list(dist_path.glob(f"vba_edit-{version}*.tar.gz"))

        if whl_files or tar_files:
            print("\n   ✅ Created packages:")
            for file in whl_files + tar_files:
                print(f"      - {file.name}")
        else:
            print(f"\n   ⚠️  Warning: No package files found for version {version}")


def main():
    parser = argparse.ArgumentParser(description="Build vba-edit executables with all prerequisites")
    parser.add_argument("--apps", nargs="*", help="Specific apps to build (default: all)")
    parser.add_argument("--skip-install", action="store_true", help="Skip package reinstallation")
    parser.add_argument(
        "--skip-package-build", action="store_true", help="Skip building Python packages (.whl/.tar.gz)"
    )
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

    # Step 4: Check build dependencies
    if not args.skip_package_build:
        check_build_package()
    check_pyinstaller()

    # Step 5: Build Python packages (.whl and .tar.gz)
    if not args.skip_package_build:
        build_python_packages()

    # Step 6: Build executables
    print("\n🔨 Building executables...")
    build_args = [sys.executable, "create_binaries.py"]
    if args.apps:
        build_args.extend(["--apps"] + args.apps)

    run_command(build_args, "Building executables")

    # Step 7: Test executables
    print("\n🧪 Testing executables...")
    dist_path = Path("dist")
    if dist_path.exists():
        for exe in dist_path.glob("*-vba.exe"):
            run_command([str(exe), "--version"], f"Testing {exe.name}", check=False)

    print("\n🎉 Build process completed successfully!")

    # Step 8: Optionally create tag
    if args.create_tag:
        create_git_tag()

    print("\nNext steps:")
    print("- Test the executables thoroughly")
    print("- Update CHANGELOG.md if needed")
    print("- Create a git tag: python build_release.py --tag-only")
    print("- Upload to GitHub: .\\scripts\\release.ps1")


if __name__ == "__main__":
    main()
