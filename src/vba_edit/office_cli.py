"""Generic Office VBA CLI factory for creating standardized CLI interfaces.

This module provides a centralized, generic CLI factory that eliminates code duplication
across Office VBA command-line tools (Excel, Word, Access, PowerPoint).

## Key Benefits

- **Massive Code Reduction**: Each Office module is now just 3-4 lines instead of 200+
- **Single Source of Truth**: All CLI logic consolidated in one maintainable location
- **Consistent Behavior**: All Office tools behave identically with unified error handling
- **Easy Maintenance**: Bug fixes and new features automatically apply to all Office apps
- **Extensibility**: Adding new Office applications requires minimal effort
- **Clean Special Handling**: Office-specific features (xlwings, Access multi-DB) are isolated

## Architecture

The `OfficeVBACLI` class provides a generic implementation that can be configured for any
Office application. Special handling for unique features is managed through:

- **Hook Functions**: Pre-command processing (e.g., Access multi-database checks)
- **Handler Functions**: Special command processing (e.g., Excel xlwings integration)
- **Configuration**: Office-specific settings and arguments

## Usage

```python
# Create simplified Office modules
from vba_edit.office_cli import create_office_main

# Excel VBA module becomes just:
main = create_office_main("excel")

# Word VBA module becomes just:
main = create_office_main("word")

# etc.
```

## Supported Office Applications

- **Excel**: Full support including xlwings integration
- **Word**: Standard VBA operations
- **Access**: Standard operations with multi-database handling
- **PowerPoint**: Standard VBA operations

## Special Handling

### Excel-Specific Features
- xlwings wrapper integration (`--xlwings` flag)
- Excel-specific argument handling

### Access-Specific Features
- Multi-database detection and handling
- Database-specific warning messages

### Standard Features (Word, PowerPoint)
- No special handling required
- Full standard VBA functionality
"""

import argparse
import logging
import sys
from pathlib import Path

from vba_edit import __name__ as package_name
from vba_edit import __version__ as package_version
from vba_edit.exceptions import (
    ApplicationError,
    DocumentClosedError,
    DocumentNotFoundError,
    PathError,
    RPCError,
    VBAAccessError,
    VBAError,
)
from vba_edit.office_vba import ExcelVBAHandler, WordVBAHandler, AccessVBAHandler, PowerPointVBAHandler
from vba_edit.path_utils import get_document_paths
from vba_edit.utils import get_active_office_document, get_windows_ansi_codepage, setup_logging
from vba_edit.cli_common import (
    add_common_arguments,
    process_config_file,
    add_encoding_arguments,
    add_header_arguments,
    add_metadata_arguments,
    add_excel_specific_arguments,
    validate_header_options,
    create_office_cli_description,
    get_help_string,
    get_office_config,
)

logger = logging.getLogger(__name__)

# Mapping of Office applications to their handler classes
OFFICE_HANDLERS = {
    "excel": ExcelVBAHandler,
    "word": WordVBAHandler,
    "access": AccessVBAHandler,
    "powerpoint": PowerPointVBAHandler,
}


def _get_office_function(office_app: str, function_name: str):
    """Dynamically import and return an office-specific function.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)
        function_name: Name of the function to import

    Returns:
        The requested function, or None if not found
    """
    try:
        if office_app == "excel":
            from vba_edit import excel_vba

            return getattr(excel_vba, function_name, None)
        elif office_app == "word":
            from vba_edit import word_vba

            return getattr(word_vba, function_name, None)
        elif office_app == "access":
            from vba_edit import access_vba

            return getattr(access_vba, function_name, None)
        elif office_app == "powerpoint":
            from vba_edit import powerpoint_vba

            return getattr(powerpoint_vba, function_name, None)
    except (ImportError, AttributeError):
        return None

    return None


# Office-specific configuration with dynamic function names
OFFICE_SPECIAL_HANDLING = {
    "access": {
        "pre_command_hook": "access_pre_command_hook",
        "extra_notes": ["NOTE: The database will remain open - close it manually when finished."],
    },
    "excel": {"xlwings_handler": "excel_xlwings_handler", "extra_arguments": add_excel_specific_arguments},
    "word": {},
    "powerpoint": {},
}


class OfficeVBACLI:
    """Generic Office VBA CLI that can be configured for any Office application."""

    def __init__(self, office_app: str):
        """Initialize CLI for specific Office application.

        Args:
            office_app: Office application name (excel, word, access, powerpoint)
        """
        self.office_app = office_app
        self.config = get_office_config(office_app)
        self.handler_class = OFFICE_HANDLERS[office_app]
        self.special_config = OFFICE_SPECIAL_HANDLING.get(office_app, {})
        self.logger = logging.getLogger(f"{__name__}.{office_app}")

    def _get_special_function(self, function_key: str):
        """Get a special function for this office app."""
        function_name = self.special_config.get(function_key)
        if isinstance(function_name, str):
            return _get_office_function(self.office_app, function_name)
        return function_name

    def create_cli_parser(self) -> argparse.ArgumentParser:
        """Create the command-line interface parser."""
        entry_point_name = self.config["entry_point"]
        package_name_formatted = package_name.replace("_", "-")

        # Get system default encoding
        default_encoding = get_windows_ansi_codepage() or "cp1252"

        # Generate description using centralized template
        description = create_office_cli_description(self.office_app, package_name_formatted, package_version)

        parser = argparse.ArgumentParser(
            prog=entry_point_name,
            description=description,
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )

        # Add --version argument to the main parser
        parser.add_argument(
            "--version", action="version", version=f"{package_name_formatted} v{package_version} ({entry_point_name})"
        )
        add_common_arguments(parser)

        subparsers = parser.add_subparsers(dest="command", required=True)

        # Edit command
        edit_parser = subparsers.add_parser("edit", help=get_help_string("edit", self.office_app))
        add_common_arguments(edit_parser)
        add_encoding_arguments(edit_parser, default_encoding)
        add_header_arguments(edit_parser)

        # Import command
        import_parser = subparsers.add_parser("import", help=get_help_string("import", self.office_app))
        add_common_arguments(import_parser)
        add_encoding_arguments(import_parser, default_encoding)
        add_header_arguments(import_parser)

        # Export command
        export_parser = subparsers.add_parser("export", help=get_help_string("export", self.office_app))
        add_common_arguments(export_parser)
        add_encoding_arguments(export_parser, default_encoding)
        add_header_arguments(export_parser)
        add_metadata_arguments(export_parser)

        # Check command
        check_parser = subparsers.add_parser("check", help=get_help_string("check", self.office_app))
        add_common_arguments(check_parser)

        # Add office-specific arguments
        extra_args_func = self._get_special_function("extra_arguments")
        if extra_args_func:
            extra_args_func(edit_parser)
            extra_args_func(import_parser)
            extra_args_func(export_parser)

        return parser

    def validate_paths(self, args: argparse.Namespace) -> None:
        """Validate file and directory paths from command line arguments."""
        if args.file and not Path(args.file).exists():
            doc_type = self.config["document_type"]
            raise FileNotFoundError(f"{doc_type.title()} not found: {args.file}")

        if args.vba_directory:
            vba_dir = Path(args.vba_directory)
            if not vba_dir.exists():
                self.logger.info(f"Creating VBA directory: {vba_dir}")
                vba_dir.mkdir(parents=True, exist_ok=True)

    def handle_office_vba_command(self, args: argparse.Namespace) -> None:
        """Handle the office-vba command execution."""
        try:
            # Initialize logging
            setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))
            self.logger.debug(f"Starting {self.office_app}-vba command: {args.command}")
            self.logger.debug(f"Command arguments: {vars(args)}")

            # Ensure paths exist early (creates vba_directory if provided)
            self.validate_paths(args)

            # Run office-specific pre-command hook
            pre_hook = self._get_special_function("pre_command_hook")
            if pre_hook:
                pre_hook(args)

            # Handle xlwings option if present (Excel only)
            xlwings_handler = self._get_special_function("xlwings_handler")
            if xlwings_handler and xlwings_handler(self, args):
                return  # xlwings handled the command

            # Get document path and active document path
            active_doc = None
            if not args.file:
                try:
                    active_doc = get_active_office_document(self.office_app)
                except ApplicationError:
                    pass

            try:
                doc_path, vba_dir = get_document_paths(args.file, active_doc, args.vba_directory)
                doc_type = self.config["document_type"]
                self.logger.info(f"Using {doc_type}: {doc_path}")
                self.logger.debug(f"Using VBA directory: {vba_dir}")
            except (DocumentNotFoundError, PathError) as e:
                self.logger.error(f"Failed to resolve paths: {str(e)}")
                sys.exit(1)

            # Determine encoding
            encoding = None if getattr(args, "detect_encoding", False) else args.encoding
            self.logger.debug(f"Using encoding: {encoding or 'auto-detect'}")

            # Validate header options
            validate_header_options(args)

            # Create handler instance
            try:
                handler = self.handler_class(
                    doc_path=str(doc_path),
                    vba_dir=str(vba_dir),
                    encoding=encoding,
                    verbose=getattr(args, "verbose", False),
                    save_headers=getattr(args, "save_headers", False),
                    use_rubberduck_folders=getattr(args, "rubberduck_folders", False),
                    open_folder=args.open_folder,
                    in_file_headers=getattr(args, "in_file_headers", True),
                )
            except VBAError as e:
                app_name = self.config["app_name"]
                self.logger.error(f"Failed to initialize {app_name} VBA handler: {str(e)}")
                sys.exit(1)

            # Execute requested command
            self.logger.info(f"Executing command: {args.command}")
            try:
                if args.command == "edit":
                    print("NOTE: Deleting a VBA module file will also delete it in the VBA editor!")

                    # Add office-specific notes
                    extra_notes = self.special_config.get("extra_notes", [])
                    for note in extra_notes:
                        print(note)

                    handler.export_vba(overwrite=False)
                    try:
                        handler.watch_changes()
                    except (DocumentClosedError, RPCError) as e:
                        self.logger.error(str(e))
                        app_name = self.config["app_name"]
                        self.logger.info(
                            f"Edit session terminated. Please restart {app_name} and the tool to continue editing."
                        )
                        sys.exit(1)
                elif args.command == "import":
                    handler.import_vba()
                elif args.command == "export":
                    handler.export_vba(save_metadata=getattr(args, "save_metadata", False), overwrite=True)
            except (DocumentClosedError, RPCError) as e:
                self.logger.error(str(e))
                sys.exit(1)
            except VBAAccessError as e:
                self.logger.error(str(e))
                app_name = self.config["app_name"]
                self.logger.error(f"Please check {app_name} Trust Center Settings and try again.")
                sys.exit(1)
            except VBAError as e:
                self.logger.error(f"VBA operation failed: {str(e)}")
                sys.exit(1)
            except Exception as e:
                self.logger.error(f"Unexpected error: {str(e)}")
                if getattr(args, "verbose", False):
                    self.logger.exception("Detailed error information:")
                sys.exit(1)

        except KeyboardInterrupt:
            self.logger.info("\nOperation interrupted by user")
            sys.exit(0)
        except Exception as e:
            self.logger.error(f"Critical error: {str(e)}")
            if getattr(args, "verbose", False):
                self.logger.exception("Detailed error information:")
            sys.exit(1)
        finally:
            self.logger.debug("Command execution completed")

    def main(self) -> None:
        """Main entry point for the Office VBA CLI."""
        try:
            parser = self.create_cli_parser()
            args = parser.parse_args()

            # Process configuration file BEFORE setting up logging
            args = process_config_file(args)

            # Set up logging first
            setup_logging(verbose=getattr(args, "verbose", False), logfile=getattr(args, "logfile", None))

            # Create target directories and validate inputs early
            self.validate_paths(args)

            # Run 'check' command (Check if VBA project model is accessible)
            if args.command == "check":
                from vba_edit.utils import check_vba_trust_access

                try:
                    if getattr(args, "subcommand", None) == "all":
                        check_vba_trust_access()  # Check all supported Office applications
                    else:
                        check_vba_trust_access(self.office_app)  # Check specific Office app only
                except Exception as e:
                    self.logger.error(f"Failed to check Trust Access to VBA project object model: {str(e)}")
                sys.exit(0)
            else:
                self.handle_office_vba_command(args)

        except Exception as e:
            print(f"Critical error: {str(e)}", file=sys.stderr)
            sys.exit(1)


def create_office_main(office_app: str):
    """Create a main function for a specific Office application.

    Args:
        office_app: Office application name (excel, word, access, powerpoint)

    Returns:
        Main function for the Office application
    """

    def main():
        cli = OfficeVBACLI(office_app)
        cli.main()

    return main
