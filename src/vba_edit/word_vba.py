import argparse
import logging
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import setup_logging, get_document_path, get_windows_ansi_codepage
from vba_edit.office_vba import (
    WordVBAHandler,
    VBAError,
    VBAAccessError,
    DocumentClosedError,
    RPCError,
)

# Configure module logger
logger = logging.getLogger(__name__)


def create_cli_parser() -> argparse.ArgumentParser:
    """Create the command-line interface parser."""
    entry_point_name = "word-vba"
    package_name_formatted = package_name.replace("_", "-")

    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"

    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name_formatted} v{package_version} ({entry_point_name})

A command-line tool suite for managing VBA content in MS Office documents.

WORD-VBA allows you to edit, import, and export VBA content from Word documents.
If no file is specified, the tool will attempt to use the currently active Word document.

Commands:
    edit    Edit VBA content in Word document
    import  Import VBA content into Word document
    export  Export VBA content from Word document

Examples:
    word-vba edit   <--- uses active Word document and current directory for exported 
                         VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                         active Word document on save

    word-vba import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
    word-vba export --file "C:/path/to/document.docx" --encoding cp850 --save-metadata

IMPORTANT: 
           [!] It's early days. Use with care and backup your imortant macro-enabled
               MS Office documents before using them with this tool!

               First tests have been very promissing. Feedback appreciated via
               github issues. 

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in Word.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Create parsers for each command with common arguments
    common_args = {
        "file": (["--file", "-f"], {"help": "Path to Word document (optional, defaults to active workbook)"}),
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

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Word document")
    encoding_group = edit_parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when reading VBA files from Word document (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Word document",
    )

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into Word document")
    import_parser.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to be used when writing VBA files back into Word document (default: {default_encoding})",
        default=default_encoding,
    )

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Word document")
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
        help=f"Encoding to be used when reading VBA files from Word document (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Word document",
    )

    # Add common arguments to all subparsers
    subparser_list = [edit_parser, import_parser, export_parser]
    for subparser in subparser_list:
        for arg_name, (arg_flags, arg_kwargs) in common_args.items():
            subparser.add_argument(*arg_flags, **arg_kwargs)

    return parser


def validate_paths(args: argparse.Namespace) -> None:
    """Validate file and directory paths from command line arguments."""
    if args.file and not Path(args.file).exists():
        raise FileNotFoundError(f"Document not found: {args.file}")

    if args.vba_directory:
        vba_dir = Path(args.vba_directory)
        if not vba_dir.exists():
            logger.info(f"Creating VBA directory: {vba_dir}")
            vba_dir.mkdir(parents=True, exist_ok=True)


def handle_word_vba_command(args: argparse.Namespace) -> None:
    """Handle the word-vba command execution."""
    try:
        # Initialize logging
        setup_logging(verbose=getattr(args, "verbose", False))
        logger.debug(f"Starting word-vba command: {args.command}")
        logger.debug(f"Command arguments: {vars(args)}")

        # Validate paths
        try:
            validate_paths(args)
        except FileNotFoundError as e:
            logger.error(str(e))
            sys.exit(1)

        # Get document path
        try:
            doc_path = get_document_path(file_path=args.file, app_type="word")
            logger.info(f"Using document: {doc_path}")
        except Exception as e:
            logger.error(f"Failed to get document path: {str(e)}")
            sys.exit(1)

        # Determine encoding
        encoding = None if getattr(args, "detect_encoding", False) else args.encoding
        logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

        # Create handler instance
        try:
            handler = WordVBAHandler(
                doc_path=doc_path,
                vba_dir=args.vba_directory,
                encoding=encoding,
                verbose=getattr(args, "verbose", False),
            )
        except Exception as e:
            logger.error(f"Failed to initialize Word VBA handler: {str(e)}")
            sys.exit(1)

        # Execute requested command
        logger.info(f"Executing command: {args.command}")
        try:
            if args.command == "edit":
                handler.export_vba()
                try:
                    handler.watch_changes()
                except (DocumentClosedError, RPCError) as e:
                    logger.error(str(e))
                    logger.info("Edit session terminated. Please restart Word and the tool to continue editing.")
                    sys.exit(1)
            elif args.command == "import":
                handler.import_vba()
            elif args.command == "export":
                handler.export_vba(save_metadata=getattr(args, "save_metadata", False))
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        except VBAAccessError as e:
            logger.error(str(e))
            logger.error("Please check Word Trust Center Settings and try again.")
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


def main() -> None:
    """Main entry point for the word-vba CLI."""
    try:
        parser = create_cli_parser()
        args = parser.parse_args()
        handle_word_vba_command(args)
    except Exception as e:
        print(f"Critical error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
