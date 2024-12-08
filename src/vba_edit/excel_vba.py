import argparse
import logging
import os
import subprocess
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import setup_logging, get_document_path, get_windows_ansi_codepage
from vba_edit.office_vba import (
    ExcelVBAHandler,
    VBAError,
    VBAAccessError,
    DocumentClosedError,
    RPCError,
)

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "excel-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool suite for managing VBA content in MS Office documents.

EXCEL-VBA allows you to edit, import, and export VBA content from Excel workbooks.
If no file is specified, the tool will attempt to use the currently active Excel workbook.

Commands:
    edit    Edit VBA content in Excel workbook
    import  Import VBA content into Excel workbook
    export  Export VBA content from Excel workbook

Examples:
    excel-vba edit   <--- uses active Excel workbook and current directory for exported 
                         VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                         active Excel workbook on save

    excel-vba import -f "C:/path/to/workbook.xlsm" --vba-directory "path/to/vba/files"
    excel-vba export --file "C:/path/to/workbook.xlsm" --encoding cp850 --save-metadata

IMPORTANT: 
           [!] It's early days. Use with care and backup your important macro-enabled
               MS Office documents before using them with this tool!

               First tests have been very promising. Feedback appreciated via
               github issues. 

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in Excel.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Create parsers for each command with common arguments
    common_args = {
        "file": (["--file", "-f"], {"help": "Path to Excel workbook (optional, defaults to active workbook)"}),
        "vba_directory": (
            ["--vba-directory"],
            {"help": "Directory to export VBA files to (optional, defaults to current directory)"},
        ),
        "verbose": (["--verbose", "-v"], {"action": "store_true", "help": "Enable verbose logging output"}),
        "logfile": (
            ["--logfile", "-l"],
            {
                "nargs": "?",
                "const": "vba_edit.log",
                "help": "Enable logging to file. Optional path can be specified (default: vba_edit.log)",
            },
        ),
        "xlwings": (
            ["--xlwings", "-x"],
            {
                "action": "store_true",
                "help": "Use wrapper for xlwings vba edit|import|export commands",
            },
        ),
    }

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Excel workbook")
    encoding_group = edit_parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when reading VBA files from Excel workbook (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Excel workbook",
    )

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into Excel workbook")
    import_parser.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when writing VBA files back into Excel workbook (default: {default_encoding})",
        default=default_encoding,
    )

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Excel workbook")
    export_parser.add_argument(
        "--save-metadata",
        "-m",
        action="store_true",
        help="Save metadata file with character encoding information (default: False)",
    )
    encoding_group = export_parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when reading VBA files from Excel workbook (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Excel workbook",
    )

    # Add common arguments to all subparsers
    subparser_list = [edit_parser, import_parser, export_parser]
    for subparser in subparser_list:
        for arg_name, (arg_flags, arg_kwargs) in common_args.items():
            subparser.add_argument(*arg_flags, **arg_kwargs)

    return parser


def handle_excel_vba_command(args: argparse.Namespace) -> None:
    """Handle the excel-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        logger.debug(f"Starting excel-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Validate paths
        try:
            validate_paths(args)
        except FileNotFoundError as e:
            logger.error(str(e))
            sys.exit(1)

        # Get document path
        try:
            doc_path = get_document_path(file_path=args.file, app_type="excel")
            logger.info(f"Using workbook: {doc_path}")
        except Exception as e:
            logger.error(f"Failed to get workbook path: {str(e)}")
            sys.exit(1)

        # Determine encoding
        encoding = None if getattr(args, "detect_encoding", False) else args.encoding
        logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

        # Create handler instance
        try:
            handler = ExcelVBAHandler(
                doc_path=doc_path,
                vba_dir=args.vba_directory,
                encoding=encoding,
                verbose=getattr(args, "verbose", False),
            )
        except Exception as e:
            logger.error(f"Failed to initialize Excel VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                # For edit command, first export without overwriting
                handler.export_vba(overwrite=False)
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart Excel and the tool to continue editing.")
                    sys.exit(1)
            elif args.command == "import":
                handler.import_vba()
            elif args.command == "export":
                # For export command, always overwrite
                handler.export_vba(save_metadata=getattr(args, "save_metadata", False), overwrite=True)
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        except VBAAccessError as e:
            logger.error(str(e))
            logger.error("Please check Excel Trust Center Settings and try again.")
            sys.exit(1)
        except VBAError as e:
            logger.error(f"VBA operation failed: {str(e)}")
            sys.exit(1)
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")
            if getattr(args, "verbose", False):
                logger.exception("Detailed error information:")
            sys.exit(1)

    except KeyboardInterrupt:
        logger.info("\nOperation interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Critical error: {str(e)}")
        if getattr(args, "verbose", False):
            logger.exception("Detailed error information:")
        sys.exit(1)
    finally:
        logger.debug("Command execution completed")


def handle_xlwings_command(args):
    """Handle command execution using xlwings wrapper."""

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


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments."""
    if args.file and not Path(args.file).exists():
        raise FileNotFoundError(f"Workbook not found: {args.file}")

    if args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def main() -> None:
    """Main entry point for the excel-vba CLI."""
    try:
        parser = create_cli_parser()
        args = parser.parse_args()

        # Set up logging first
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))

        # If xlwings option is used, check dependency before proceeding
        if args.xlwings:
            try:
                import xlwings

                logger.info(f"Using xlwings {xlwings.__version__}")
                handle_xlwings_command(args)
            except ImportError:
                sys.exit("xlwings is not installed. Please install it with: pip install xlwings")
        else:
            handle_excel_vba_command(args)

    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)

    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
