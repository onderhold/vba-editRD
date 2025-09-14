"""Common CLI argument definitions for all Office VBA handlers."""

import argparse
import os
from typing import Dict, Any
from vba_edit.utils import get_windows_ansi_codepage

try:
    import tomli as toml_lib  # Python 3.11+ includes tomllib in stdlib
except ImportError:
    try:
        import tomli as toml_lib  # tomli is the recommended TOML parser for Python <3.11
    except ImportError:
        import toml as toml_lib  # Fall back to toml package if tomli isn't available


def load_config_file(config_path: str) -> Dict[str, Any]:
    """Load configuration from a TOML file.
    
    Args:
        config_path: Path to the TOML configuration file
        
    Returns:
        Dictionary containing the configuration
        
    Raises:
        FileNotFoundError: If the configuration file doesn't exist
        ValueError: If the configuration file isn't valid TOML
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    
    with open(config_path, "rb") as f:
        return toml_lib.load(f)


def merge_config_with_args(args: argparse.Namespace, config: Dict[str, Any]) -> argparse.Namespace:
    """Merge configuration from a file with command-line arguments.
    
    Command-line arguments take precedence over configuration file values.
    
    Args:
        args: Command-line arguments
        config: Configuration from file
        
    Returns:
        Updated arguments with values from configuration
    """
    # Create a copy of the args namespace as a dictionary
    args_dict = vars(args).copy()
    
    # For each key in the config, if the corresponding arg is None or not present,
    # update the arg with the config value
    for key, value in config.items():
        # Convert dashes to underscores for argument names
        arg_key = key.replace("-", "_")
        
        # Only update if the arg wasn't explicitly set (is None or is the default)
        if arg_key in args_dict and args_dict[arg_key] is None:
            args_dict[arg_key] = value
    
    # Convert back to a Namespace
    return argparse.Namespace(**args_dict)


def process_config_file(args: argparse.Namespace) -> argparse.Namespace:
    """Load configuration file if specified and merge with command-line arguments.
    
    Args:
        args: Command-line arguments
        
    Returns:
        Updated arguments with values from configuration file
    """
    if not args.conf:
        return args
    
    try:
        config = load_config_file(args.conf)
        return merge_config_with_args(args, config)
    except Exception as e:
        print(f"Error loading configuration file: {e}")
        return args


def add_common_arguments(parser: argparse.ArgumentParser) -> None:
    """Add common arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument("--file", "-f", help="Path to Office document (optional, defaults to active document)")
    parser.add_argument(
        "--vba-directory", help="Directory to export VBA files to (optional, defaults to current directory)"
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging output")
    parser.add_argument(
        "--logfile",
        "-l",
        nargs="?",
        const="vba_edit.log",
        help="Enable logging to file. Optional path can be specified (default: vba_edit.log)",
    )
    parser.add_argument(
        "--rubberduck-folders",
        action="store_true",
        help="If a module contains a RubberduckVBA '@Folder annotation, organize folders in the file system accordingly",
    )
    parser.add_argument(
        "--conf",
        help="Path to configuration file (TOML format) with default arguments",
    )

def add_excel_specific_arguments(parser: argparse.ArgumentParser) -> None:
    """Add Excel-specific arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--xlwings", "-x", action="store_true", help="Use wrapper for xlwings vba edit|import|export commands"
    )


def add_encoding_arguments(parser: argparse.ArgumentParser, default_encoding: str = None) -> None:
    """Add encoding-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
        default_encoding: Default encoding to use
    """
    if default_encoding is None:
        default_encoding = get_windows_ansi_codepage() or "cp1252"

    encoding_group = parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when reading/writing VBA files (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding", "-d", action="store_true", help="Auto-detect input encoding for VBA files"
    )


def add_header_arguments(parser: argparse.ArgumentParser) -> None:
    """Add header-related arguments to a parser.

    Args:
        parser: The argument parser to add arguments to
    """
    parser.add_argument(
        "--save-headers",
        action="store_true",
        help="Save VBA component headers to separate .header files (default: False)",
    )
