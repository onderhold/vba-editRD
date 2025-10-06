onderhold: __init__.py and pyproject.toml both contain a version definition. can this somehow be unified?

GitHub Copilot: Yes, you can unify the version definition by having a **single source of truth** for your project version.  
The recommended approach is:

- **Define the version only in pyproject.toml.**
- In __init__.py, read the version dynamically from pyproject.toml (or use a tool to inject it).

### Option 1: Use `importlib.metadata` (Python ≥3.8)

```python
from importlib.metadata import version, PackageNotFoundError

try:
    __version__ = version("vba-editRD")
except PackageNotFoundError:
    __version__ = "unknown"
```

This reads the version from the installed package metadata, which is set from pyproject.toml.

### Option 2: Use a tool like `bumpver` or `setuptools_scm`

You already use `bumpver` and have file patterns configured in pyproject.toml:
```toml
[tool.bumpver.file_patterns]
    "pyproject.toml" = [
        'current_version = "{version}"',
        'version = "{version}"',
    ]
    "src/vba_edit/__init__.py" = ["{version}"]
```
This means when you bump the version, `bumpver` will update both files automatically.

**Summary:**  
- If you want true single-source, use `importlib.metadata` in `__init__.py`.
- If you want explicit version strings in both files, keep using `bumpver` as configured.

**Best practice:**  
Prefer dynamic reading (`importlib.metadata`) for less manual maintenance and avoid version mismatches.

onderhold: how does bumpver work? do I have to trigger/execute it myself?

GitHub Copilot: `bumpver` is a CLI tool for managing and updating version numbers in your project files.  
You **trigger it manually** via the command line when you want to bump the version.

### How it works in your setup

- You define the current version and file patterns in pyproject.toml.
- When you run a `bumpver` command, it updates the version in all configured files (e.g., pyproject.toml, __init__.py, etc.).
- It can also commit the change to git, tag it, and push (depending on your config).

### Typical usage

To bump the patch version:
```sh
bumpver update --patch
```
To bump the minor version:
```sh
bumpver update --minor
```
To bump the major version:
```sh
bumpver update --major
```
Or interactively:
```sh
bumpver update
```

### Summary

- **You must execute `bumpver` yourself**—it does not run automatically.
- It updates all files listed in `[tool.bumpver.file_patterns]` in your pyproject.toml.
- It helps keep your version numbers in sync across files.

onderhold: how about the importlib.metadata, can I use this for package_name in \__init\__.py as well?

GitHub Copilot: Yes, you can use `importlib.metadata` to dynamically get both the version and the package name in __init__.py.  
However, the package name is typically static (e.g., `"vba-editRD"`), while the version is dynamic.

Here’s how you can do it:

````python
from importlib.metadata import version, PackageNotFoundError

package_name = "vba-editRD"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    __version__ = "unknown"

__author__ = "onderhold"
__description__ = "Enable seamless MS Office VBA code editing in preferred editor or IDE"
````

**Notes:**
- `package_name` should match the name in your pyproject.toml.
- This approach ensures `__version__` is always in sync with your installed package metadata.
- You still need to set `package_name` manually, as `importlib.metadata` does not provide a way to discover the package name from within the code.

**Summary:**  
- Use `importlib.metadata.version(package_name)` for the version.
- Keep `package_name` as a static string.

onderhold: but if I would set package_name in \__init\__.py the same as "name" in pyproject.toml, could this be done dynamically?

GitHub Copilot: No, you **cannot set `package_name` dynamically from pyproject.toml at runtime** using standard Python tools.  
Python code does not have direct access to the pyproject.toml metadata at runtime.

**What you can do:**
- Set `package_name` manually in __init__.py to match the name in pyproject.toml.
- Dynamically get the version using `importlib.metadata.version(package_name)`.

**Example:**
````python
from importlib.metadata import version, PackageNotFoundError

package_name = "vba-editRD"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    __version__ = "unknown"

__author__ = "onderhold"
__description__ = "Enable seamless MS Office VBA code editing in preferred editor or IDE"
````

**Summary:**  
- `package_name` must be set manually.
- The version can be read dynamically.
- There is no built-in way for Python to read the package name from pyproject.toml at runtime.

onderhold: so then, could I use bumpver to set package_name?

GitHub Copilot: Yes, you **can use bumpver to set `package_name`** in your __init__.py, but this is not typical or recommended.  
`bumpver` is designed to update version strings, not package names, since the package name rarely changes.

**How it would work:**
- You would add a file pattern for `package_name` in your `[tool.bumpver.file_patterns]` section in pyproject.toml.
- When you run `bumpver`, it would update both the version and the package name wherever you specify.

**Example file pattern:**
```toml
[tool.bumpver.file_patterns]
"src/vba_edit/__init__.py" = [
    '__version__ = "{version}"',
    'package_name = "{package_name}"'
]
```
But:  
- `bumpver` only updates the version by default.  
- You would need to manually trigger a package name change, which is rare and not automated.

**Summary:**  
- Technically possible, but not standard practice.
- The package name should be set manually and only changed if you rename your project.
- Use `bumpver` for version updates; keep `package_name` static unless you have a special use case.

onderhold: when I change package_name in \__init\__.py, where else will I have to change it?

GitHub Copilot: If you change `package_name` in __init__.py, you should also update it in the following places:

1. **pyproject.toml**  
   - The `[project]` section’s `name` field must match your new package name:
     ```toml
     name = "your-new-package-name"
     ```

2. **Anywhere you reference the package name in your code**  
   - For example, in other modules (like `excel_vba.py`) if you import or use `package_name` from __init__.py.

3. **Documentation and README files**  
   - Update installation instructions, usage examples, and references to the package name.

4. **Setup scripts or CI/CD configs**  
   - If you use scripts or workflows that reference the package name (e.g., for publishing to PyPI), update those as well.

5. **Import statements (if you rename the actual package directory)**  
   - If you rename the folder `vba_edit` itself, update all import statements accordingly.

**Summary:**  
Change the package name in __init__.py, pyproject.toml, and anywhere else it is referenced in code, docs, or scripts.  
The most critical is pyproject.toml—it must match for packaging and distribution.

onderhold: what are the requirements for package_version? can I use e.g. "0.3.1-rdf"?

GitHub Copilot: Yes, you can use a version string like `"0.3.1-rdf"` for `package_version`.  
The requirements for `package_version` are:

- It should be a valid [PEP 440](https://www.python.org/dev/peps/pep-0440/) version string for Python packages.
- `"0.3.1-rdf"` is valid as a [local version identifier](https://www.python.org/dev/peps/pep-0440/#local-version-identifiers).
- This string is typically set in your pyproject.toml under `version = "0.3.1-rdf"`.
- `package_version` in your code is just a variable holding the version string, so it can be any string, but for packaging and distribution, it should follow PEP 440.

**Summary:**  
You can use `"0.3.1-rdf"` for `package_version` as long as it matches the version in your pyproject.toml and follows PEP 440 conventions.

onderhold: do I have to define  \__author\__ and \__description\__ ?

GitHub Copilot: No, you do **not** have to define `__author__` and `__description__` in your __init__.py.  
They are **optional** and are mainly for informational purposes.

- `__author__` is sometimes used by tools or for documentation, but it is not required by Python packaging.
- `__description__` is not a standard Python variable; it is just a custom string for your own use.

**Required for packaging:**  
- The `name`, `version`, `description`, and `authors` fields in your pyproject.toml are what matter for packaging and distribution.

**Summary:**  
You can remove `__author__` and `__description__` from __init__.py if you do not need them. They are not required.

