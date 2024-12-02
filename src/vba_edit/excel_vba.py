import argparse
from vba_edit import __version__ as package_version
from vba_edit import __name__ as package_name

package_name = package_name.replace("_", "-")

def vba_edit(file: str = None) -> None:
    """Edit Excel VBA code.

    Args:
        file: Optional path to the Excel workbook
    """
    print(f"Editing VBA content in {file or 'active workbook'}")
    # Implement VBA editing logic here
    raise NotImplementedError("VBA editing without xlwings is not implemented yet")


def vba_import(file: str = None) -> None:
    """Import Excel VBA code.

    Args:
        file: Optional path to the Excel workbook
    """
    print(f"Importing VBA content from {file or 'active workbook'}")
    # Implement VBA import logic here
    raise NotImplementedError("VBA import without xlwings is not implemented yet")


def vba_export(file: str = None) -> None:
    """Export Excel VBA code.

    Args:
        file: Optional path to the Excel workbook
    """
    print(f"Exporting VBA content from {file or 'active workbook'}")
    # Implement VBA export logic here
    raise NotImplementedError("VBA export without xlwings is not implemented yet")


def main():
    vba_edit_version = f"{package_name} v{package_version}"
    entry_point_name = __name__.split(".")[-1].replace("_", "-")

    # Check if xlwings is installed
    try:
        from xlwings import __version__ as xw_version
        from xlwings.cli import vba_edit as xw_vba_edit
        from xlwings.cli import vba_import as xw_vba_import
        from xlwings.cli import vba_export as xw_vba_export

        USE_XLWINGS = True
    except ImportError:
        print("xlwings is not installed. Please install it to use this script.")
        USE_XLWINGS = False

    if USE_XLWINGS:
        excel_vba_processing_engine = f"xlwings {xw_version}"
        edit_function = xw_vba_edit
        import_function = xw_vba_import
        export_function = xw_vba_export
    else:
        excel_vba_processing_engine = f"{vba_edit_version} (fallback if 'xlwings' library is not installed)"
        edit_function = vba_edit
        import_function = vba_import
        export_function = vba_export

    """Main entry point for the excel-vba CLI."""
    
    # Implement the CLI logic here
    parser = argparse.ArgumentParser(
    prog=entry_point_name,
    description=f"""
{package_name} v{package_version} ({entry_point_name})

A command-line tool for managing VBA content in Excel documents.
This tool allows you to edit, import, and export VBA content from Excel documents.
If no file is specified, the tool will attempt to use the currently active Excel document.

Commands:
    edit    Edit VBA content in Excel document
    import  Import VBA content into Excel document
    export  Export VBA content from Excel document

Examples:
    {entry_point_name} edit
    {entry_point_name} import -f "C:/path/to/document.docx"
    {entry_point_name} export

IMPORTANT: This tool requires "Trust access to the VBA project object model" enabled in Excel.

Inspired by xlwings' vba functionality: https://docs.xlwings.org/en/stable/command_line.html#command-line

For more information, visit: https://github.com/markuskiller/vba-edit
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument("command", choices=["edit", "import", "export"], help="Command to execute")
    parser.add_argument(
        "-f",
        "--file",
        help="Optional parameter to select a specific workbook, otherwise it uses the " "active one.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Optional parameter to print messages whenever a module has been updated " "successfully.",
    )
    parser.add_argument("--version", action="version", version=vba_edit_version)

    args = parser.parse_args()

    if USE_XLWINGS:
        xw_command = f"\n\tCommand: xlwings VBA {(args.command).upper()} {args.file or ''}"
    else:
        xw_command = "\n"

    print(f"\nInteracting with Excel VBA code using {excel_vba_processing_engine}")

    if args.command == "edit":
        print(xw_command)
        edit_function(args)
    elif args.command == "import":
        print(xw_command)
        import_function(args)
    elif args.command == "export":
        print(xw_command)
        export_function(args)


if __name__ == "__main__":
    main()
