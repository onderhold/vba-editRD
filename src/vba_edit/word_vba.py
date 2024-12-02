import argparse
import datetime
import json
import sys
from pathlib import Path
from typing import Optional

import win32com.client

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import detect_vba_encoding, get_document_path, get_windows_ansi_codepage

package_name = package_name.replace("_", "-")


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


def word_vba_export(doc_path: str, encoding: Optional[str] = "cp1252") -> None:
    """Export Word VBA content.

    Args:
        doc_path: Path to the Word document
        encoding: Encoding to use for VBA files. If None, encoding will be auto-detected.
                 Defaults to cp1252.
    """
    print(f"Exporting VBA content from {doc_path}")
    if encoding is None:
        print("Auto-detecting encodings for VBA files...")
    else:
        print(f"Using encoding: {encoding}")

    word = None
    try:
        # Create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # Open document
        doc = word.Documents.Open(str(doc_path))

        # Check if VBA project is accessible
        try:
            vba_project = doc.VBProject
        except Exception as e:
            raise RuntimeError(
                "Cannot access VBA project. Please ensure 'Trust access to the VBA project object model' "
                "is enabled in Word Trust Center Settings."
            ) from e

        components = vba_project.VBComponents

        if not components.Count:
            print("No VBA components found in the document.")
            return

        print(f"\nFound {components.Count} VBA components:")
        component_names = [component.Name for component in components]
        print(", ".join(component_names))

        # Create export directory
        export_dir = Path(doc_path).parent / "VBA"
        export_dir.mkdir(exist_ok=True)

        detected_encodings = {}

        # Export each component
        for component in components:
            try:
                # Map component types to file extensions
                type_to_ext = {
                    1: ".bas",  # Standard Module
                    2: ".cls",  # Class Module
                    3: ".frm",  # MSForm
                    100: ".cls",  # Document Module
                }

                if component.Type not in type_to_ext:
                    print(f"Skipping {component.Name} (unsupported type {component.Type})")
                    continue

                # Generate file paths
                file_name = component.Name + type_to_ext[component.Type]
                temp_file_path = export_dir / f"{file_name}.temp"
                final_file_path = export_dir / file_name

                # Export to temporary file first
                component.Export(str(temp_file_path))

                used_encoding = None
                confidence = 1.0

                if encoding is None:
                    # Only detect encoding if explicitly requested
                    used_encoding, confidence = detect_vba_encoding(str(temp_file_path))
                else:
                    used_encoding = encoding

                try:
                    with open(temp_file_path, "r", encoding=used_encoding) as source:
                        content = source.read()
                except UnicodeDecodeError as e:
                    raise RuntimeError(
                        f"Failed to decode file with {'detected' if encoding is None else 'specified'} "
                        f"encoding {used_encoding}. Error: {e}"
                    )

                with open(final_file_path, "w", encoding=used_encoding) as target:
                    target.write(content)

                # Remove temporary file
                temp_file_path.unlink()

                detected_encodings[file_name] = {
                    "encoding": used_encoding,
                    "confidence": confidence if encoding is None else 1.0,
                }

                if encoding is None:
                    print(
                        f"Exported: {final_file_path} (Detected encoding: {used_encoding}, "
                        f"confidence: {confidence:.2%})"
                    )
                else:
                    print(f"Exported: {final_file_path} (Using encoding: {used_encoding})")

            except Exception as e:
                print(f"Warning: Failed to export {component.Name}: {e}")
                continue

        # Store metadata including encoding information
        metadata_path = export_dir / "vba_metadata.json"
        metadata = {
            "source_document": str(doc_path),
            "export_date": datetime.datetime.now().isoformat(),
            "encoding_mode": "auto-detect" if encoding is None else "fixed",
            "encodings": detected_encodings,
        }

        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2)

    except Exception as e:
        raise RuntimeError(f"Failed to export VBA content: {e}")

    finally:
        if word is not None:
            try:
                doc.Close(SaveChanges=False)
                word.Quit()
            except win32com.client.pywintypes.com_error as e:
                print(f"Warning: Failed to cleanly close Word: {e}", file=sys.stderr)


def main() -> None:
    """Main entry point for the word-vba CLI."""
    entry_point_name = __name__.split(".")[-1].replace("_", "-")
    # Get system default encoding
    default_encoding = get_windows_ansi_codepage() or "cp1252"
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
    {entry_point_name} export --file "C:/path/to/document.docx" --encoding cp850

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

    # Export command encoding group
    encoding_group = export_parser.add_mutually_exclusive_group()
    encoding_group.add_argument(
        "--encoding",
        "-e",
        help=f"Encoding to use for VBA files (default: {default_encoding})",
        default=default_encoding,
    )
    encoding_group.add_argument(
        "--detect-encoding", "-d", action="store_true", help="Auto-detect encoding for VBA files"
    )

    args = parser.parse_args()

    try:
        # Get document path once
        doc_path = get_document_path(file_path=args.file, app_type="word")

        # Pass the resolved path to functions
        if args.command == "edit":
            word_vba_edit(doc_path)
        elif args.command == "import":
            word_vba_import(doc_path)
        elif args.command == "export":
            word_vba_export(doc_path, encoding=None if args.detect_encoding else args.encoding)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
