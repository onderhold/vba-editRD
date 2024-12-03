import argparse
import sys

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import get_document_path, get_windows_ansi_codepage
from vba_edit.office_vba import WordVBAHandler


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

A command-line tool for managing VBA content in Word documents.
This tool allows you to edit, import, and export VBA content from Word documents.
If no file is specified, the tool will attempt to use the currently active Word document.

Commands:
    edit    Edit VBA content in Word document
    import  Import VBA content into Word document
    export  Export VBA content from Word document

Examples:
    {entry_point_name} edit
    {entry_point_name} import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
    {entry_point_name} export --file "C:/path/to/document.docx" --encoding cp850 --save-metadata

IMPORTANT: This tool requires "Trust access to the VBA project object model" enabled in Word.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Word document")
    edit_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")
    edit_parser.add_argument(
        "--vba-directory", help="Directory to export VBA files to (optional, defaults to current directory)"
    )
    encoding_group = edit_parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to use to read VBA files from Word document (default: {default_encoding})",
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
    import_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")
    import_parser.add_argument(
        "--vba-directory",
        help="Directory containing VBA files to be imported into Word document (optional, defaults to current directory)",
    )
    import_parser.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to use to write VBA files back into Word document (default: {default_encoding})",
        default=default_encoding,
    )

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Word document")
    export_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")
    export_parser.add_argument(
        "--vba-directory", help="Directory to export VBA files to (optional, defaults to current directory)"
    )
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
        help=f"Encoding to use to read VBA files from Word document (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding",
        "-d",
        action="store_true",
        help="Auto-detect input encoding for VBA files exported from Word document",
    )

    return parser


def handle_word_vba_command(args: argparse.Namespace) -> None:
    """Handle the word-vba command execution."""
    try:
        # Get document path
        doc_path = get_document_path(file_path=args.file, app_type="word")

        # Determine encoding
        encoding = None if getattr(args, "detect_encoding", False) else args.encoding

        # Create handler instance
        handler = WordVBAHandler(doc_path=doc_path, vba_dir=args.vba_directory, encoding=encoding)

        # Execute requested command
        if args.command == "edit":
            handler.export_vba()  # First export all content
            handler.watch_changes()  # Then watch for changes
        elif args.command == "import":
            handler.import_vba()
        elif args.command == "export":
            handler.export_vba(save_metadata=getattr(args, "save_metadata", False))

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def main() -> None:
    """Main entry point for the word-vba CLI."""
    parser = create_cli_parser()
    args = parser.parse_args()
    handle_word_vba_command(args)


if __name__ == "__main__":
    main()
