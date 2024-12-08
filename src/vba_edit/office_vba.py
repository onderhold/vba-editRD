from abc import ABC, abstractmethod
import datetime
import json
import logging
import os
import sys
import time
from pathlib import Path
from typing import Dict, Optional, Any, List
import win32com.client
from pywintypes import com_error
from watchgod import Change, RegExpWatcher

# Configure logging
logger = logging.getLogger(__name__)


def check_rpc_error(error: Exception) -> bool:
    """Check if an error is related to RPC server unavailability.

    Args:
        error: The exception to check

    Returns:
        bool: True if the error is RPC-related
    """
    error_str = str(error).lower()
    rpc_indicators = [
        "rpc server",
        "rpc-server",
        "remote procedure call",
        "0x800706BA",  # RPC server unavailable error code
        "-2147023174",  # Same error in decimal
    ]
    return any(indicator in error_str for indicator in rpc_indicators)


class VBAError(Exception):
    """Base exception class for VBA-related errors."""

    pass


class VBAAccessError(VBAError):
    """Exception raised when VBA project access is denied."""

    pass


class VBAImportError(VBAError):
    """Exception raised during VBA import operations."""

    pass


class VBAExportError(VBAError):
    """Exception raised during VBA export operations."""

    pass


class DocumentClosedError(VBAError):
    """Exception raised when attempting to access a closed document."""

    def __init__(self, doc_type: str = "document"):
        super().__init__(
            f"\nThe Office {doc_type} has been closed. The edit session will be terminated.\n"
            f"IMPORTANT: Any changes made after closing the {doc_type} must be imported using\n"
            f"'*-vba import' before starting a new edit session, otherwise they will be lost."
        )


class RPCError(VBAError):
    """Exception raised when RPC server is unavailable."""

    def __init__(self, app_name: str = "Office application"):
        super().__init__(
            f"\nLost connection to {app_name}. The edit session will be terminated.\n"
            "IMPORTANT: Any changes made after losing connection must be imported using\n"
            "'*-vba import' before starting a new edit session, otherwise they will be lost."
        )


class OfficeVBAHandler(ABC):
    """Base class for handling VBA operations across different Office applications."""

    def __init__(self, doc_path: str, vba_dir: Optional[str] = None, encoding: str = "cp1252", verbose: bool = False):
        self.doc_path = doc_path
        self.vba_dir = Path(vba_dir) if vba_dir else Path.cwd()
        self.vba_dir = self.vba_dir.resolve()
        self.encoding = encoding
        self.verbose = verbose
        self.app = None
        self.doc = None

        # Configure logging based on verbosity
        log_level = logging.DEBUG if verbose else logging.INFO
        logger.setLevel(log_level)

        # Map component types to file extensions
        self.type_to_ext = {
            1: ".bas",  # Standard Module
            2: ".cls",  # Class Module
            3: ".frm",  # MSForm
            100: ".cls",  # Document Module
        }

        logger.debug(f"Initialized {self.__class__.__name__} with document: {doc_path}")
        logger.debug(f"VBA directory: {self.vba_dir}")
        logger.debug(f"Encoding: {encoding}")

    @property
    @abstractmethod
    def app_name(self) -> str:
        """Name of the Office application."""
        pass

    @property
    @abstractmethod
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        pass

    @abstractmethod
    def get_vba_project(self) -> Any:
        """Get the VBA project from the document."""
        pass

    @abstractmethod
    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        pass

    @abstractmethod
    def is_document_open(self) -> bool:
        """Check if the document is still open and accessible."""
        pass

    @abstractmethod
    def import_vba(self) -> None:
        """Import VBA content into the Office document."""
        pass

    @abstractmethod
    def import_single_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed."""
        pass

    @abstractmethod
    def watch_changes(self) -> None:
        """Watch for changes in VBA files and automatically reimport them."""
        pass

    @abstractmethod
    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA content from the Office document.

        Args:
            save_metadata: Whether to save encoding metadata
            overwrite: Whether to overwrite existing files. If False, only exports files
                       that don't exist yet and sheet modules that contain code.
        """
        pass

    def initialize_app(self) -> None:
        """Initialize the Office application."""
        try:
            if self.app is None:
                logger.debug(f"Initializing {self.app_name} application")
                self.app = win32com.client.Dispatch(self.app_progid)
                self.app.Visible = True
        except Exception as e:
            error_msg = f"Failed to initialize {self.app_name} application"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def open_document(self) -> None:
        """Open the Office document."""
        try:
            if self.doc is None:
                self.initialize_app()
                logger.debug(f"Opening document: {self.doc_path}")
                if self.app_name == "Word":
                    self.doc = self.app.Documents.Open(str(self.doc_path))
                elif self.app_name == "Excel":
                    self.doc = self.app.Workbooks.Open(str(self.doc_path))
        except Exception as e:
            error_msg = f"Failed to open document: {self.doc_path}"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def save_document(self) -> None:
        """Save the document if it's open."""
        if self.doc is not None:
            try:
                self.doc.Save()
                logger.info("Document has been saved and left open for further editing")
            except Exception as e:
                # Don't log the error here since it will be handled at a higher level
                raise VBAError("Failed to save document") from e

    def handle_document_module(self, component: Any, content: str, temp_file: Path) -> None:
        """Handle the special document module."""
        try:
            # Skip header section for document module
            content_lines = content.splitlines()
            if len(content_lines) > 9:
                actual_code = "\n".join(content_lines[9:])
            else:
                actual_code = ""

            logger.debug(f"Processing document module: {component.Name}")

            # Convert content to specified encoding
            content_bytes = actual_code.encode(self.encoding)

            with open(temp_file, "wb") as f:
                f.write(content_bytes)

            # Read back with proper encoding
            with open(temp_file, "r", encoding=self.encoding) as f:
                new_code = f.read()

            # Update existing document module
            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            if new_code.strip():
                component.CodeModule.AddFromString(new_code)

            logger.debug(f"Successfully updated document module: {component.Name}")

        except Exception as e:
            error_msg = f"Failed to handle document module: {component.Name}"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def get_component_list(self) -> List[Dict[str, Any]]:
        """Get list of VBA components with their details."""
        try:
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            component_list = []
            for component in components:
                component_info = {
                    "name": component.Name,
                    "type": component.Type,
                    "code_lines": component.CodeModule.CountOfLines if hasattr(component, "CodeModule") else 0,
                    "extension": self.type_to_ext.get(component.Type, "unknown"),
                }
                component_list.append(component_info)

            return component_list
        except Exception as e:
            error_msg = "Failed to get component list"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def _save_metadata(self, encodings: Dict[str, Dict[str, Any]]) -> None:
        """Save metadata including encoding information."""
        try:
            metadata = {
                "source_document": str(self.doc_path),
                "export_date": datetime.datetime.now().isoformat(),
                "encoding_mode": "fixed",
                "encodings": encodings,
            }

            metadata_path = self.vba_dir / "vba_metadata.json"
            with open(metadata_path, "w", encoding="utf-8") as f:
                json.dump(metadata, f, indent=2)

            logger.info(f"Metadata saved to {metadata_path}")

        except Exception as e:
            error_msg = "Failed to save metadata"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e


class WordVBAHandler(OfficeVBAHandler):
    """Word-specific VBA handler implementation."""

    @property
    def app_name(self) -> str:
        return "Word"

    @property
    def app_progid(self) -> str:
        return "Word.Application"

    def get_vba_project(self) -> Any:
        try:
            return self.doc.VBProject
        except Exception as e:
            error_msg = (
                "Cannot access VBA project. Please ensure 'Trust access to the VBA project object model' "
                "is enabled in Word Trust Center Settings."
            )
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAAccessError(error_msg) from e

    def get_document_module_name(self) -> str:
        return "ThisDocument"

    def is_document_open(self) -> bool:
        try:
            if self.doc is None:
                return False

            # Try to access document name and state
            _ = self.doc.Name

            # Check if document is still open in Word
            for doc in self.app.Documents:
                if doc.FullName == self.doc_path:
                    return True
            return False

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            return False

    def import_vba(self) -> None:
        """Import VBA content into the Word document."""
        try:
            # First check if document is accessible
            if self.doc is None:
                self.open_document()
            _ = self.doc.Name  # Check connection

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            vba_files = [f for f in self.vba_dir.glob("*.*") if f.suffix in self.type_to_ext.values()]
            if not vba_files:
                logger.info("No VBA files found to import.")
                return

            logger.info(f"\nFound {len(vba_files)} VBA files to import:")
            for vba_file in vba_files:
                logger.info(f"  - {vba_file.name}")

            for vba_file in vba_files:
                temp_file = None
                try:
                    logger.debug(f"Processing {vba_file.name}")
                    with open(vba_file, "r", encoding="utf-8") as f:
                        content = f.read()

                    component_name = vba_file.stem
                    temp_file = vba_file.with_suffix(".temp")

                    if component_name == self.get_document_module_name():
                        # Handle ThisDocument module
                        doc_component = components(self.get_document_module_name())
                        self.handle_document_module(doc_component, content, temp_file)
                    else:
                        # Handle regular components
                        content_bytes = content.encode(self.encoding)
                        with open(temp_file, "wb") as f:
                            f.write(content_bytes)

                        # Remove existing component if it exists
                        try:
                            existing = components(component_name)
                            components.Remove(existing)
                            logger.debug(f"Removed existing component: {component_name}")
                        except Exception:
                            logger.debug(f"No existing component to remove: {component_name}")

                        # Import the component
                        components.Import(str(temp_file))

                    temp_file.unlink()
                    logger.info(f"Imported: {vba_file.name}")

                except Exception:
                    if temp_file and temp_file.exists():
                        temp_file.unlink()
                    raise  # Re-raise to be handled by outer try/except

            # Only try to save if we successfully imported all files
            self.save_document()

        except Exception as e:
            if check_rpc_error(e):
                raise DocumentClosedError("document")
            raise VBAImportError(str(e))

    def import_single_file(self, file_path: Path) -> None:
        logger.info(f"Processing changes in {file_path.name}")
        temp_file = None

        try:
            # Check if document is still open
            if not self.is_document_open():
                raise DocumentClosedError("document")

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents
            component_name = file_path.stem

            # Read content with UTF-8 encoding (as exported)
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            if component_name == self.get_document_module_name():
                logger.debug("Processing ThisDocument module")
                doc_component = components(self.get_document_module_name())
                temp_file = file_path.with_suffix(".temp")
                self.handle_document_module(doc_component, content, temp_file)
            else:
                logger.debug(f"Processing regular component: {component_name}")
                content_bytes = content.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Remove existing component if it exists
                try:
                    existing = components(component_name)
                    logger.debug(f"Removing existing component: {component_name}")
                    components.Remove(existing)
                except Exception:
                    logger.debug(f"No existing component to remove: {component_name}")

                # Import the component
                logger.debug(f"Importing component: {component_name}")
                components.Import(str(temp_file))

            logger.info(f"Successfully imported: {file_path.name}")
            self.doc.Save()

        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except Exception as e:
                    logger.warning(f"Failed to remove temporary file {temp_file}: {e}")

    def watch_changes(self) -> None:
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 30  # Check connection every 30 seconds

            # Setup the file watcher
            watcher = RegExpWatcher(self.vba_dir, re_files=r"^.*(\.cls|\.frm|\.bas)$")

            while True:
                # Always check connection if interval has elapsed
                current_time = time.time()
                if current_time - last_check_time >= check_interval:
                    if not self.is_document_open():
                        raise DocumentClosedError("document")
                    last_check_time = current_time

                # Check for file changes
                changes = watcher.check()
                if changes:
                    for change_type, path in changes:
                        if change_type == Change.modified:
                            try:
                                logger.debug(f"Detected change in {path}")
                                self.import_single_file(Path(path))
                            except (DocumentClosedError, RPCError) as e:
                                raise e
                            except Exception as e:
                                logger.warning(f"Error handling changes (will retry): {str(e)}")
                                continue

                # Small sleep to prevent excessive CPU usage
                time.sleep(0.8)

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        finally:
            logger.info("VBA editor stopped.")

    def export_vba(self, save_metadata: bool = False) -> None:
        try:
            self.open_document()
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            if not components.Count:
                logger.info("No VBA components found in the document.")
                return

            component_list = self.get_component_list()
            logger.info(f"\nFound {len(component_list)} VBA components:")
            for comp in component_list:
                logger.info(f"  - {comp['name']} ({comp['extension']}, {comp['code_lines']} lines)")

            detected_encodings = {}

            for component in components:
                try:
                    if component.Type not in self.type_to_ext:
                        logger.warning(f"Skipping {component.Name} (unsupported type {component.Type})")
                        continue

                    file_name = f"{component.Name}{self.type_to_ext[component.Type]}"
                    temp_file = self.vba_dir / f"{file_name}.temp"
                    final_file = self.vba_dir / file_name

                    logger.debug(f"Exporting component {component.Name} to {final_file}")

                    # Export to temporary file
                    component.Export(str(temp_file))

                    # Read with specified encoding and write as UTF-8
                    with open(temp_file, "r", encoding=self.encoding) as source:
                        content = source.read()

                    with open(final_file, "w", encoding="utf-8") as target:
                        target.write(content)

                    temp_file.unlink()
                    logger.info(f"Exported: {final_file}")

                except Exception as e:
                    error_msg = f"Failed to export {component.Name}"
                    logger.error(f"{error_msg}: {str(e)}")
                    if temp_file.exists():
                        temp_file.unlink()
                    continue

            if save_metadata:
                self._save_metadata(detected_encodings)

            os.startfile(self.vba_dir)

        except Exception as e:
            error_msg = "Failed to export VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAExportError(error_msg) from e
        finally:
            self.save_document()


class ExcelVBAHandler(OfficeVBAHandler):
    """Excel-specific VBA handler implementation."""

    # VBA Component Type Constants
    VBEXT_CT_DOCUMENT = 100  # Document module type
    VBEXT_CT_MSFORM = 3  # UserForm type
    VBEXT_CT_STDMODULE = 1  # Standard module type
    VBEXT_CT_CLASSMODULE = 2  # Class module type

    # Excel Type Constants
    XL_WORKSHEET = -4167  # xlWorksheet type

    @property
    def app_name(self) -> str:
        return "Excel"

    @property
    def app_progid(self) -> str:
        return "Excel.Application"

    def get_vba_project(self) -> Any:
        try:
            return self.doc.VBProject
        except Exception as e:
            error_msg = (
                "Cannot access VBA project. Please ensure 'Trust access to the VBA project object model' "
                "is enabled in Excel Trust Center Settings."
            )
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAAccessError(error_msg) from e

    def get_document_module_name(self) -> str:
        return "ThisWorkbook"

    def is_document_open(self) -> bool:
        try:
            if self.doc is None:
                return False

            # Try to access workbook name and state
            _ = self.doc.Name

            # Check if workbook is still open in Excel
            for wb in self.app.Workbooks:
                if wb.FullName == self.doc_path:
                    return True
            return False

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            return False

    def _get_component_type(self, component_name: str) -> Optional[int]:
        """Get the VBA component type for an existing component.

        Args:
            component_name: Name of the component

        Returns:
            int: The VBA component type if found, None otherwise
        """
        try:
            component = self.doc.VBProject.VBComponents(component_name)
            return component.Type
        except (AttributeError, com_error) as e:
            logger.debug(f"Could not get component type for {component_name}: {str(e)}")
            return None

    def _is_worksheet_module(self, component: Any) -> bool:
        """Check if a component is a worksheet module using VBA type constants.

        Args:
            component: VBA component to check

        Returns:
            bool: True if component is a worksheet module
        """
        try:
            return (
                component.Type == self.VBEXT_CT_DOCUMENT  # Is document module
                and hasattr(component.Parent, "Type")  # Has a parent with Type
                and component.Parent.Type == self.XL_WORKSHEET
            )  # Parent is worksheet
        except (AttributeError, com_error) as e:
            logger.debug(f"Could not check worksheet module type: {str(e)}")
            return False

    def import_single_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed."""
        logger.info(f"Processing changes in {file_path.name}")
        temp_file = None

        try:
            # Check if workbook is still open
            if not self.is_document_open():
                raise DocumentClosedError("workbook")

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents
            component_name = file_path.stem

            # Read content with UTF-8 encoding (as exported)
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            if component_name == self.get_document_module_name():
                # Handle ThisWorkbook module
                logger.debug("Processing ThisWorkbook module")
                wb_component = components(self.get_document_module_name())
                temp_file = file_path.with_suffix(".temp")
                self.handle_document_module(wb_component, content, temp_file)

            elif component_name.startswith("Sheet") and component_name.replace("Sheet", "").isdigit():
                # This is a worksheet module - update existing module's code
                logger.debug(f"Processing worksheet module: {component_name}")
                try:
                    sheet_component = components(component_name)
                    if sheet_component.CodeModule.CountOfLines > 0:
                        sheet_component.CodeModule.DeleteLines(1, sheet_component.CodeModule.CountOfLines)
                    if content.strip():
                        sheet_component.CodeModule.AddFromString(content)
                    logger.info(f"Updated worksheet module: {file_path.name}")
                except Exception as e:
                    logger.error(f"Could not update worksheet module {component_name}: {str(e)}")
                    raise

            else:
                # Handle regular components
                logger.debug(f"Processing regular module: {component_name}")
                content_bytes = content.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Remove existing component if it exists
                try:
                    existing = components(component_name)
                    components.Remove(existing)
                    logger.debug(f"Removed existing component: {component_name}")
                except Exception:
                    logger.debug(f"No existing component to remove: {component_name}")

                # Import the component
                components.Import(str(temp_file))
                logger.info(f"Imported: {file_path.name}")

            self.doc.Save()

        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except Exception as e:
                    logger.warning(f"Failed to remove temporary file {temp_file}: {e}")

    def watch_changes(self) -> None:
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 30  # Check connection every 30 seconds

            # Setup the file watcher
            watcher = RegExpWatcher(self.vba_dir, re_files=r"^.*(\.cls|\.frm|\.bas)$")

            while True:
                # Always check connection if interval has elapsed
                current_time = time.time()
                if current_time - last_check_time >= check_interval:
                    if not self.is_document_open():
                        raise DocumentClosedError("workbook")
                    last_check_time = current_time

                # Check for file changes
                changes = watcher.check()
                if changes:
                    for change_type, path in changes:
                        if change_type == Change.modified:
                            try:
                                logger.debug(f"Detected change in {path}")
                                self.import_single_file(Path(path))
                            except (DocumentClosedError, RPCError) as e:
                                raise e
                            except Exception as e:
                                logger.warning(f"Error handling changes (will retry): {str(e)}")
                                continue

                # Small sleep to prevent excessive CPU usage
                time.sleep(0.8)

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        except (DocumentClosedError, RPCError) as e:
            logger.error(str(e))
            sys.exit(1)
        finally:
            logger.info("VBA editor stopped.")

    def import_vba(self) -> None:
        """Import VBA content into the Excel workbook."""
        try:
            # First check if document is accessible
            if self.doc is None:
                self.open_document()
            _ = self.doc.Name  # Check connection

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            vba_files = [f for f in self.vba_dir.glob("*.*") if f.suffix in self.type_to_ext.values()]
            if not vba_files:
                logger.info("No VBA files found to import.")
                return

            logger.info(f"\nFound {len(vba_files)} VBA files to import:")
            for vba_file in vba_files:
                logger.info(f"  - {vba_file.name}")

            for vba_file in vba_files:
                temp_file = None
                try:
                    component_name = vba_file.stem
                    logger.debug(f"Processing {vba_file.name}")

                    with open(vba_file, "r", encoding="utf-8") as f:
                        content = f.read()

                    if component_name == self.get_document_module_name():
                        # Handle ThisWorkbook module
                        logger.debug("Processing ThisWorkbook module")
                        wb_component = components(self.get_document_module_name())
                        temp_file = vba_file.with_suffix(".temp")
                        self.handle_document_module(wb_component, content, temp_file)

                    elif component_name.startswith("Sheet") and component_name.replace("Sheet", "").isdigit():
                        # This is a worksheet module - update existing module's code
                        logger.debug(f"Processing worksheet module: {component_name}")
                        try:
                            sheet_component = components(component_name)
                            if sheet_component.CodeModule.CountOfLines > 0:
                                sheet_component.CodeModule.DeleteLines(1, sheet_component.CodeModule.CountOfLines)
                            if content.strip():
                                sheet_component.CodeModule.AddFromString(content)
                            logger.info(f"Updated worksheet module: {vba_file.name}")
                        except Exception as e:
                            logger.error(f"Could not update worksheet module {component_name}: {str(e)}")
                            continue

                    else:
                        # Regular module or class
                        logger.debug(f"Processing regular module: {component_name}")
                        content_bytes = content.encode(self.encoding)
                        temp_file = vba_file.with_suffix(".temp")

                        with open(temp_file, "wb") as f:
                            f.write(content_bytes)

                        # Remove existing if it exists
                        try:
                            existing = components(component_name)
                            components.Remove(existing)
                            logger.debug(f"Removed existing component: {component_name}")
                        except Exception:
                            logger.debug(f"No existing component to remove: {component_name}")

                        # Import the component
                        components.Import(str(temp_file))
                        logger.info(f"Imported: {vba_file.name}")

                    if temp_file and temp_file.exists():
                        temp_file.unlink()

                except Exception as e:
                    logger.error(f"Failed to process {vba_file.name}: {str(e)}")
                    if temp_file and temp_file.exists():
                        temp_file.unlink()
                    continue

            # Only try to save if we successfully imported all files
            self.save_document()

        except Exception as e:
            if check_rpc_error(e):
                raise DocumentClosedError("workbook")
            raise VBAImportError(str(e))

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA content from the Excel workbook.

        Args:
            save_metadata: Whether to save metadata about the export
            overwrite: Whether to overwrite existing files. If False, only exports files
                      that don't exist yet and sheet modules that contain code.
        """
        try:
            self.open_document()
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            if not components.Count:
                logger.info("No VBA components found in the workbook.")
                return

            component_list = self.get_component_list()
            logger.info(f"\nFound {len(component_list)} VBA components:")
            for comp in component_list:
                logger.info(f"  - {comp['name']} ({comp['extension']}, {comp['code_lines']} lines)")

            detected_encodings = {}

            for component in components:
                try:
                    if component.Type not in self.type_to_ext:
                        logger.warning(f"Skipping {component.Name} (unsupported type {component.Type})")
                        continue

                    file_name = f"{component.Name}{self.type_to_ext[component.Type]}"
                    temp_file = self.vba_dir / f"{file_name}.temp"
                    final_file = self.vba_dir / file_name

                    # Skip if file exists and we're not overwriting
                    # For sheet modules (Type 100), only export if they contain code
                    should_export = (
                        overwrite
                        or not final_file.exists()
                        or (component.Type == self.VBEXT_CT_DOCUMENT and component.CodeModule.CountOfLines > 0)
                    )

                    if not should_export:
                        logger.debug(f"Skipping existing file: {final_file}")
                        continue

                    logger.debug(f"Exporting component {component.Name} to {final_file}")

                    # Export to temporary file
                    component.Export(str(temp_file))

                    # Read with specified encoding and write as UTF-8
                    # For worksheet modules, skip the header
                    with open(temp_file, "r", encoding=self.encoding) as source:
                        if component.Type == self.VBEXT_CT_DOCUMENT:
                            content = "".join(source.readlines()[9:])  # Skip header for document modules
                        else:
                            content = source.read()

                    with open(final_file, "w", encoding="utf-8") as target:
                        target.write(content)

                    temp_file.unlink()
                    logger.info(f"Exported: {final_file}")

                except Exception as e:
                    error_msg = f"Failed to export {component.Name}"
                    logger.error(f"{error_msg}: {str(e)}")
                    if temp_file.exists():
                        temp_file.unlink()
                    continue

            if save_metadata:
                self._save_metadata(detected_encodings)

            os.startfile(self.vba_dir)

        except Exception as e:
            error_msg = "Failed to export VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAExportError(error_msg) from e
        finally:
            self.save_document()
