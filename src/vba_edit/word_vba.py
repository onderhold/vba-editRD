import argparse
import sys
from typing import Optional
import win32com.client
from pathlib import Path
from vba_edit import __version__ as package_version
from vba_edit import __name__ as package_name

package_name = package_name.replace("_", "-")


def get_active_word_document() -> str:
    """Get the path of the currently active Word document.

    Returns:
        str: Full path to the active document

    Raises:
        RuntimeError: If Word is not running or no document is active
    """
    try:
        word = win32com.client.GetObject(Class="Word.Application")

        if not word.Documents.Count:
            raise RuntimeError("No Word document is currently open")

        active_doc = word.ActiveDocument
        if not active_doc:
            raise RuntimeError("Could not get active Word document")

        return active_doc.FullName
    except Exception as e:
        raise RuntimeError(f"Could not connect to Word or get active document: {e}")


def get_document_path(file_path: Optional[str] = None) -> str:
    """Get the document path from either the provided file path or active document.

    Args:
        file_path: Optional path to the Word document

    Returns:
        str: Path to the document

    Raises:
        RuntimeError: If no valid document path can be determined
    """
    doc_path = file_path or get_active_word_document()

    if not Path(doc_path).exists():
        raise RuntimeError(f"Document not found: {doc_path}")

    return doc_path


def word_vba_edit(doc_path: str) -> None:
    """Edit Word VBA content.

    Args:
        doc_path: Path to the Word document
    """
    print(f"Editing VBA content in {doc_path}")
    # Implement VBA editing logic here


def word_vba_import(doc_path: str) -> None:
    """Import Word VBA content.

    Args:
        doc_path: Path to the Word document
    """
    print(f"Importing VBA content from {doc_path}")
    # Implement VBA import logic here


def word_vba_export(doc_path: str) -> None:
    """Export Word VBA content.

    Args:
        doc_path: Path to the Word document
    """
    print(f"Exporting VBA content from {doc_path}")
    # Implement VBA export logic here


def main() -> None:
    """Main entry point for the word-vba CLI."""
    entry_point_name = __name__.split(".")[-1].replace("_", "-")
    parser = argparse.ArgumentParser(
        prog=entry_point_name,
        description=f"""
{package_name} v{package_version} ({entry_point_name})

A command-line tool for managing VBA content in Word documents.
This tool allows you to edit, import, and export VBA content from Word documents.
If no file is specified, the tool will attempt to use the currently active Word document.

Commands:
    edit    Edit VBA content in Word document
    import  Import VBA content into Word document
    export  Export VBA content from Word document

Examples:
    {entry_point_name} edit
    {entry_point_name} import -f "C:/path/to/document.docx"
    {entry_point_name} export

IMPORTANT: This tool requires "Trust access to the VBA project object model" enabled in Word.

Inspired by xlwings' vba functionality: https://docs.xlwings.org/en/stable/command_line.html#command-line

For more information, visit: https://github.com/markuskiller/vba-edit
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Edit command
    edit_parser = subparsers.add_parser("edit", help="Edit VBA content in Word document")
    edit_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA content into Word document")
    import_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA content from Word document")
    export_parser.add_argument("--file", "-f", help="Path to Word document (optional, defaults to active document)")

    args = parser.parse_args()

    try:
        # Get document path once
        doc_path = get_document_path(args.file)

        # Pass the resolved path to functions
        if args.command == "edit":
            word_vba_edit(doc_path)
        elif args.command == "import":
            word_vba_import(doc_path)
        elif args.command == "export":
            word_vba_export(doc_path)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
