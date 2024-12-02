import argparse
import asyncio
import datetime
import json
import sys
from pathlib import Path
from typing import Optional

import win32com.client
from watchgod import awatch

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.utils import VBAFileChangeHandler, detect_vba_encoding, get_document_path, get_windows_ansi_codepage

package_name = package_name.replace("_", "-")


async def watch_vba_files(doc_path: str, vba_dir: Path, encoding: Optional[str] = "cp1252") -> None:
    """Watch VBA files for changes using watchgod.

    Args:
        doc_path: Path to the Word document
        vba_dir: Directory containing VBA files to watch
        encoding: Encoding to use for writing VBA content back to Word document
    """
    handler = VBAFileChangeHandler(doc_path, str(vba_dir), encoding)

    async for changes in awatch(str(vba_dir)):
        handler.handle_changes(changes)


def word_vba_edit(doc_path: str, vba_dir: Optional[str] = None, encoding: Optional[str] = "cp1252") -> None:
    """Edit Word VBA content interactively.

    Exports VBA content to the specified directory and watches for changes,
    automatically reimporting modified files back into the Word document.

    Args:
        doc_path: Path to the Word document
        vba_dir: Directory to export/watch VBA files. Defaults to current working directory.
        encoding: Encoding to use for VBA files. If None, encoding will be auto-detected.
                 Defaults to cp1252.
    """
    print(f"Starting interactive VBA editing for {doc_path}")

    # Use current working directory if vba_dir is not specified
    vba_dir = Path(vba_dir) if vba_dir else Path.cwd()
    vba_dir = vba_dir.resolve()

    # First export all VBA content
    word_vba_export(doc_path, vba_dir=str(vba_dir), encoding=encoding)

    try:
        print("\nWaiting for changes... (Press Ctrl+C to stop)")
        # Run the async watch function in the event loop
        asyncio.run(watch_vba_files(doc_path, vba_dir, encoding))

    except KeyboardInterrupt:
        print("\nStopping VBA editor...")
    finally:
        print("VBA editor stopped. Your changes have been saved.")


def word_vba_import(doc_path: str, vba_dir: Optional[str] = None, encoding: Optional[str] = "cp1252") -> None:
    """Import Word VBA content from previously exported files.

    Args:
        doc_path: Path to the Word document
        vba_dir: Directory containing VBA files to import. Defaults to current working directory.
        encoding: Encoding to use for writing VBA content. If None, encoding will be read
                 from metadata file. Defaults to cp1252.
    """
    print(f"\n\nImporting VBA content into {doc_path}")

    # Use current working directory if vba_dir is not specified
    vba_dir = Path(vba_dir) if vba_dir else Path.cwd()
    vba_dir = vba_dir.resolve()
    print(f"\nLooking for VBA files in: {vba_dir}")

    word = None
    doc = None
    try:
        # Create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  # Make Word visible

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

        # Find VBA directory
        vba_dir = Path(doc_path).parent / "VBA"
        if not vba_dir.exists():
            raise RuntimeError(f"VBA directory not found: {vba_dir}")

        # Read metadata file if it exists
        metadata_path = vba_dir / "vba_metadata.json"
        metadata = None
        if metadata_path.exists():
            with open(metadata_path, "r", encoding="utf-8") as f:
                metadata = json.load(f)

        # Create mapping of extensions to component types
        ext_to_type = {
            ".bas": 1,  # Standard Module
            ".cls": 2,  # Class Module
            ".frm": 3,  # MSForm
        }

        # Get list of VBA files
        vba_files = [f for f in vba_dir.glob("*.*") if f.suffix in ext_to_type]
        if not vba_files:
            print("No VBA files found to import.")
            return

        print(f"\nFound {len(vba_files)} VBA files to import:")
        print(", ".join(f.name for f in vba_files))

        components = vba_project.VBComponents

        # Process each VBA file
        for vba_file in vba_files:
            try:
                # Read content with UTF-8 encoding (as exported)
                with open(vba_file, "r", encoding="utf-8") as f:
                    content = f.read()

                # Determine the output encoding
                output_encoding = encoding
                if output_encoding is None and metadata and "encodings" in metadata:
                    file_info = metadata["encodings"].get(vba_file.name, {})
                    output_encoding = file_info.get("encoding", "cp1252")

                # Remove extension to get component name
                component_name = vba_file.stem

                # Special handling for ThisDocument (Document Module)
                is_document_module = component_name == "ThisDocument"

                if is_document_module:
                    # For ThisDocument, find existing document module and update its code
                    try:
                        doc_component = components("ThisDocument")

                        # Skip the header section (first 9 lines) for ThisDocument
                        content_lines = content.splitlines()
                        if len(content_lines) > 9:
                            actual_code = "\n".join(content_lines[9:])
                        else:
                            actual_code = ""  # Empty if file is too short

                        # Convert content to specified encoding
                        content_bytes = actual_code.encode(output_encoding)

                        # Create temporary file with proper encoding
                        temp_file = vba_file.with_suffix(".temp")
                        with open(temp_file, "wb") as f:
                            f.write(content_bytes)

                        # Read back the code with proper encoding
                        with open(temp_file, "r", encoding=output_encoding) as f:
                            new_code = f.read()

                        # Update the code of the existing ThisDocument module
                        doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)
                        if new_code.strip():  # Only add if there's actual code
                            doc_component.CodeModule.AddFromString(new_code)

                        # Remove temporary file
                        temp_file.unlink()

                    except win32com.client.pywintypes.com_error as e:
                        print(f"Warning: Failed to update ThisDocument: {e}")
                else:
                    # For other components, convert content and create temp file
                    content_bytes = content.encode(output_encoding)
                    temp_file = vba_file.with_suffix(".temp")
                    with open(temp_file, "wb") as f:
                        f.write(content_bytes)

                    # Remove if exists and import new
                    try:
                        existing = components(component_name)
                        components.Remove(existing)
                    except Exception as e:
                        print(f"Warning: Failed to import {vba_file.name}: {e}")
                        if "temp_file" in locals():
                            try:
                                temp_file.unlink()
                            except (OSError, PermissionError) as e:
                                print(f"Warning: Failed to remove temporary file: {e}")
                        continue

                    # Import the component with appropriate type
                    components.Import(str(temp_file))

                    # Remove temporary file
                    temp_file.unlink()

                print(f"Imported: {vba_file.name} (Using encoding: {output_encoding})")

            except Exception as e:
                print(f"Warning: Failed to import {vba_file.name}: {e}")
                if "temp_file" in locals():
                    try:
                        temp_file.unlink()
                    except (OSError, PermissionError) as e:
                        print(f"Warning: Failed to remove temporary file: {e}")
                continue

    except Exception as e:
        raise RuntimeError(f"Failed to import VBA content: {e}")

    finally:
        if doc is not None:
            try:
                doc.Save()  # Save changes
                print("Document has been saved and left open for further editing.")
            except win32com.client.pywintypes.com_error as e:
                print(f"Warning: Failed to save document: {e}", file=sys.stderr)


def word_vba_export(doc_path: str, vba_dir: Optional[str] = None, encoding: Optional[str] = "cp1252") -> None:
    """Export Word VBA content.

    Args:
        doc_path: Path to the Word document
        vba_dir: Directory to export VBA files to. Defaults to current working directory.
        encoding: Encoding to use for VBA files. If None, encoding will be auto-detected.
                 Defaults to cp1252.
    """
    print(f"Exporting VBA content from {doc_path}")
    if encoding is None:
        print("Auto-detecting encodings for VBA files...")
    else:
        print(f"Using input encoding: {encoding}")

    # Use current working directory if vba_dir is not specified
    export_dir = Path(vba_dir) if vba_dir else Path.cwd()
    export_dir = export_dir.resolve()
    print(f"Exporting to directory: {export_dir}")

    word = None
    doc = None
    try:
        # Create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  # Make Word visible

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

        # If the export directory is not the current working directory,
        # create it if it doesn't exist
        if export_dir != Path.cwd():
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

                with open(final_file_path, "w", encoding="utf-8") as target:
                    target.write(content)

                # Remove temporary file
                temp_file_path.unlink()

                detected_encodings[file_name] = {
                    "encoding": used_encoding,
                    "confidence": confidence if encoding is None else 1.0,
                }

                if encoding is None:
                    print(
                        f"Exported: {final_file_path} (Using output encoding: utf-8)"
                        f"          (Detected input encoding: {used_encoding}, "
                        f"           confidence: {confidence:.2%})"
                    )
                else:
                    print(f"Exported: {final_file_path} (Using output encoding: utf-8)")

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
        if doc is not None:
            try:
                doc.Save()  # Save changes
                print("Document has been saved and left open for further editing.")
            except win32com.client.pywintypes.com_error as e:
                print(f"Warning: Failed to save document: {e}", file=sys.stderr)


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
    {entry_point_name} import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
    {entry_point_name} export --file "C:/path/to/document.docx" --encoding cp850

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
        help="Auto-detect input encoding for VBA files exported from Worddocument",
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

    args = parser.parse_args()

    try:
        # Get document path once
        doc_path = get_document_path(file_path=args.file, app_type="word")

        # Pass the resolved path to functions
        if args.command == "edit":
            word_vba_edit(
                doc_path, vba_dir=args.vba_directory, encoding=None if args.detect_encoding else args.encoding
            )
        elif args.command == "import":
            word_vba_import(doc_path, vba_dir=args.vba_directory, encoding=args.encoding)
        elif args.command == "export":
            word_vba_export(
                doc_path, vba_dir=args.vba_directory, encoding=None if args.detect_encoding else args.encoding
            )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
