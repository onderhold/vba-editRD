"""Common CLI argument definitions for all Office VBA handlers."""

import argparse
from vba_edit.utils import get_windows_ansi_codepage


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
        help="Enable Rubberduck @Folder annotations for organizing VBA modules in subdirectories",
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
