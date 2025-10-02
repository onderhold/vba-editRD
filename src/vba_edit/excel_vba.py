"""Excel VBA CLI module."""

import argparse
import logging
import os
import subprocess
import sys
from pathlib import Path

from vba_edit.office_cli import create_office_main

logger = logging.getLogger(__name__)


def handle_xlwings_command(args: argparse.Namespace) -> None:
    """Handle command execution using xlwings wrapper (Excel only)."""
    # Convert our args to xlwings command format
    cmd = ["xlwings", "vba", args.command]

    if args.file:
        cmd.extend(["-f", args.file])
    if args.verbose:
        cmd.extend(["-v"])

    # Store original directory
    original_dir = None
    try:
        # Change to target directory if specified
        if args.vba_directory:
            original_dir = os.getcwd()
            target_dir = Path(args.vba_directory)
            # Create directory if it doesn't exist
            target_dir.mkdir(parents=True, exist_ok=True)
            os.chdir(str(target_dir))

        # Execute xlwings command
        result = subprocess.run(cmd, check=True)
        sys.exit(result.returncode)

    except subprocess.CalledProcessError as e:
        sys.exit(e.returncode)
    finally:
        # Restore original directory if we changed it
        if original_dir:
            os.chdir(original_dir)


def excel_xlwings_handler(cli_instance, args: argparse.Namespace) -> bool:
    """Handle Excel xlwings functionality. Returns True if handled, False otherwise."""
    if not getattr(args, "xlwings", False):
        return False

    try:
        import xlwings

        cli_instance.logger.info(f"Using xlwings {xlwings.__version__}")

        # Use the xlwings command handler from the CLI instance
        handle_xlwings_command(args)
        return True
    except ImportError:
        sys.exit("xlwings is not installed. Please install it with: pip install xlwings")


# Create the main function for Excel
main = create_office_main("excel")

if __name__ == "__main__":
    main()
