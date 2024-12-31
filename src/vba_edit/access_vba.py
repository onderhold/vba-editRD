import argparse
import logging
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import setup_logging, get_document_path, get_windows_ansi_codepage
from vba_edit.office_vba import (
    AccessVBAHandler,
    VBAError,
    VBAAccessError,
    DocumentClosedError,
    RPCError,
)

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "access-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool suite for managing VBA content in MS Access databases.

ACCESS-VBA allows you to edit, import, and export VBA content from Access databases.
If no file is specified, the tool will attempt to use the currently active Access database.
Only standard modules (*.bas) and class modules (*.cls) are supported.

Commands:
    edit    Edit VBA content in Access database
    import  Import VBA content into Access database
    export  Export VBA content from Access database

Examples:
    access-vba edit   <--- uses active Access database and current directory for exported 
                          VBA files (*.bas/*.cls) & syncs changes back to the 
                          active Access database on save

    access-vba import -f "C:/path/to/database.accdb" --vba-directory "path/to/vba/files"
    access-vba export --file "C:/path/to/database.accdb" --encoding cp850 --save-metadata

IMPORTANT: 
           [!] It's early days. Use with care and backup your important macro-enabled
               MS Access databases before using them with this tool!

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in Access.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Create parsers for each command with common arguments
    common_args = {
        "file": (["--file", "-f"], {"help": "Path to Access database (optional, defaults to active database)"}),
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
    }

    # # Edit command
    # edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Access database")
    # encoding_group = edit_parser.add_mutually_exclusive_group()
    # encoding_group.add_argument(
    #     "--encoding",
    #     "-e",
    #     help=f"Encoding to be used when reading VBA files from Access database (default: {default_encoding})",
    #     default=default_encoding,
    # )
    # encoding_group.add_argument(
    #     "--detect-encoding",
    #     "-d",
    #     action="store_true",
    #     help="Auto-detect input encoding for VBA files exported from Access database",
    # )
    # edit_parser.add_argument(
    #     "--save-headers",
    #     action="store_true",
    #     help="Save VBA component headers to separate .header files (default: False)",
    # )

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into Access database")
    import_parser.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when writing VBA files back into Access database (default: {default_encoding})",
        default=default_encoding,
    )

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Access database")
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
        help=f"Encoding to be used when reading VBA files from Access database (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Access database",
    )
    export_parser.add_argument(
        "--save-headers",
        action="store_true",
        help="Save VBA component headers to separate .header files (default: False)",
    )

    # Add common arguments to all subparsers
    subparser_list = [
        # edit_parser,
        import_parser,
        export_parser,
    ]
    for subparser in subparser_list:
        for arg_name, (arg_flags, arg_kwargs) in common_args.items():
            subparser.add_argument(*arg_flags, **arg_kwargs)

    return parser


def handle_access_vba_command(args: argparse.Namespace) -> None:
    """Handle the access-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        logger.debug(f"Starting access-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Validate paths
        try:
            validate_paths(args)
        except FileNotFoundError as e:
            logger.error(str(e))
            sys.exit(1)

        # Get document path
        try:
            doc_path = get_document_path(file_path=args.file, app_type="access")
            logger.info(f"Using database: {doc_path}")
        except Exception as e:
            logger.error(f"Failed to get database path: {str(e)}")
            sys.exit(1)

        # Determine encoding
        encoding = None if getattr(args, "detect_encoding", False) else args.encoding
        logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

        # Create handler instance
        try:
            logger.debug(f"Creating AccessVBAHandler with save_headers={getattr(args, 'save_headers', False)}")
            handler = AccessVBAHandler(
                doc_path=doc_path,
                vba_dir=args.vba_directory,
                encoding=encoding,
                verbose=getattr(args, "verbose", False),
                save_headers=getattr(args, "save_headers", False),  # Explicitly log this
            )
            logger.debug(f"Handler created with save_headers={handler.save_headers}")
        except Exception as e:
            logger.error(f"Failed to initialize Access VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                print("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")
                handler.export_vba(overwrite=False)
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart Access and the tool to continue editing.")
                    sys.exit(1)
            elif args.command == "import":
                handler.import_vba()
            elif args.command == "export":
                handler.export_vba(save_metadata=getattr(args, "save_metadata", False), overwrite=True)
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        except VBAAccessError as e:
            logger.error(str(e))
            logger.error("Please check Access Trust Center Settings and try again.")
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


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments."""
    if args.file and not Path(args.file).exists():
        raise FileNotFoundError(f"Database not found: {args.file}")

    if args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def main() -> None:
    """Main entry point for the access-vba CLI."""
    try:
        parser = create_cli_parser()
        args = parser.parse_args()

        # Set up logging first
        setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
        handle_access_vba_command(args)

    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
