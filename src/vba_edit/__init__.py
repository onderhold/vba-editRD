from importlib.metadata import version, PackageNotFoundError

package_name = "vba-edit"

try:
    __version__ = version(package_name)
except PackageNotFoundError:
    __version__ = "0.4.0a3"  # Keep this in sync with pyproject.toml

package_version = __version__
