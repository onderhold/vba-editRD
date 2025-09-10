## Development Setup

### Prerequisites
- Python 3.9 or higher
- Windows (required for COM interface with MS Office)
- Git

### Initial Setup
1. Clone the repository:
```bash
git clone <repository-url>
cd vba-editRD
```

2. Create and activate a virtual environment:
```shell
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # Linux/macOS
```

or in VS Code, 
- open command palette (CTRL-SHIFT-P)
- search for: "Python: Create Environment..."
- select: Venv
- choose Python interpreter (recommended: Python 3.9+)


3. Install the package in development mode:
```shell
pip install -e .[dev]
```

This installs the package with all development dependencies including:
- Testing tools (pytest, pytest-cov)
- Code formatting (black, ruff)
- Build tools (pyinstaller, wheel, twine)

### Git Configuration

The repository already includes a comprehensive [`.gitignore`](.gitignore) that covers:
- Virtual environments (`.venv/`, `venv/`)
- Build artifacts (`build/`, `dist/`, `*.egg-info/`)
- Python cache files (`__pycache__/`, `*.pyc`)
- Test artifacts (`.pytest_cache/`, `.coverage`, `htmlcov/`)
- IDE files (`.vscode/`, `.idea/`)
- Log files (`*.log` - includes `vba_edit.log`, `vba_trust_access.log`)
- PyInstaller specs (`*.spec`)

No additional `.gitignore` configuration is needed.

## Development Workflow
### Running Tests
```shell
# Run all tests
pytest

# Run with coverage
pytest --cov=src --cov-report=html

# Run only unit tests
pytest -m "not integration"

# Run specific test file
pytest tests/test_cli.py -v
```

### Code Formatting and Linting
```shell
# Format code with black
black src tests

# Lint with ruff
ruff check src tests

# Fix linting issues automatically
ruff check --fix src tests
```

### Testing CLI Commands
```shell
word-vba --help
excel-vba --help
access-vba --help

# Test check functionality
python -m vba_edit.utils
```

## Creating Binaries
```shell
# Create standalone executables
python create_binaries.py
```

Binaries will be in the dist/ directory

## Packaging
In Python packaging, a "build" command is commonly used to create distribution files for your project.
- `python -m build` is the standard way to build Python packages.
- It reads `pyproject.toml` and creates distributable files.

**Standard build command:**
```sh
python -m build
```
This uses the [`build`](https://pypi.org/project/build/) package to generate source (`.tar.gz`) and wheel (`.whl`) distributions in the dist directory.

**Steps:**
1. Install the build tool (if not already):
   ```sh
   pip install build
   ```
2. Run the build command:
   ```sh
   python -m build
   ```

## Deployment

You can use the official GitHub CLI tool [`gh`](https://cli.github.com/) to upload an `.exe` (or any file) from a local directory to a GitHub release

### Example workflow

1. **Install GitHub CLI:**  
   [Download and install](https://cli.github.com/) for your platform.

2. **Authenticate:**  
   ```sh
   gh auth login
   ```

3. **Create a release (if needed):**  
   ```sh
   gh release create v1.0.0 --title "Release v1.0.0" --notes "Release notes here"
   ```

4. **Upload your `.exe` to the release:**  
   ```sh
   gh release upload v1.0.0 path/to/your/file.exe
   ```

You can upload any file from any location on your disk—**it does not need to be tracked by git or in your repo folder**.

**Summary:**  
Use `gh release upload <tag> <file>` to add binaries to a GitHub release from anywhere on your system.