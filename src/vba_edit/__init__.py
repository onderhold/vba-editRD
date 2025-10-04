from importlib.metadata import version, PackageNotFoundError

package_name = "vba-edit"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    __version__ = "0.4.1a1"  # Keep this in sync with pyproject.toml

package_version = __version__
