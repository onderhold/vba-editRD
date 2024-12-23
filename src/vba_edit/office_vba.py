from abc import ABC, abstractmethod
import datetime
import json
import logging
import os
import re
import shutil
import sys
import time
from enum import Enum, auto
from pathlib import Path
from typing import Dict, Optional, Any, List
import win32com.client
from pywintypes import com_error
from watchgod import Change, RegExpWatcher

"""
The VBA import/export/edit functionality is based on the excellent work done by the xlwings project
(https://github.com/xlwings/xlwings) which is distributed under the BSD 3-Clause License:

Copyright (c) 2014-present, Zoomer Analytics GmbH.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

* Neither the name of the copyright holder nor the names of its
  contributors may be used to endorse or promote products derived from
  this software without specific prior written permission.

This module extends the original xlwings VBA interaction concept to provide a consistent 
interface for interacting with VBA code across different Microsoft Office applications.
"""

# Configure logging
logger = logging.getLogger(__name__)


class VBAModuleType(Enum):
    """VBA module types"""

    DOCUMENT = auto()  # ThisWorkbook or Worksheet modules
    CLASS = auto()  # Regular class modules
    STANDARD = auto()  # Standard modules (.bas)
    FORM = auto()  # UserForm modules


def determine_cls_type(header: str) -> VBAModuleType:
    """Determine if a .cls file is a document module or regular class module.

    Args:
        header: Content of the VBA component header

    Returns:
        VBAModuleType.DOCUMENT or VBAModuleType.CLASS
    """
    # Extract key attributes
    predeclared = re.search(r"Attribute VB_PredeclaredId = (\w+)", header)
    exposed = re.search(r"Attribute VB_Exposed = (\w+)", header)

    # Document modules have both attributes set to True
    if predeclared and exposed and predeclared.group(1).lower() == "true" and exposed.group(1).lower() == "true":
        return VBAModuleType.DOCUMENT

    return VBAModuleType.CLASS


def get_module_type(file_path: Path) -> VBAModuleType:
    """Determine VBA module type from file extension and content.

    Args:
        file_path: Path to the VBA module file

    Returns:
        Appropriate VBAModuleType
    """
    suffix = file_path.suffix.lower()

    if suffix == ".bas":
        return VBAModuleType.STANDARD
    elif suffix == ".frm":
        return VBAModuleType.FORM
    elif suffix == ".cls":
        # For .cls files, we need to check the header
        header_file = file_path.with_suffix(".header")
        if header_file.exists():
            with open(header_file, "r", encoding="utf-8") as f:
                return determine_cls_type(f.read())
        else:
            logger.warning(f"No header file found for {file_path}, treating as regular class module")
            return VBAModuleType.CLASS
    else:
        raise ValueError(f"Unknown file extension: {suffix}")


def split_vba_content(content: str) -> tuple[str, str]:
    """Split VBA content into header and code sections.

    Args:
        content: Complete VBA component content

    Returns:
        Tuple of (header, code)

    Note:
        Only module-level attributes (VB_Name, VB_GlobalNameSpace, VB_Creatable,
        VB_PredeclaredId, VB_Exposed) are considered part of the header.
        Procedure-level attributes are considered part of the code.
    """
    if not content.strip():
        return "", ""

    lines = content.splitlines()

    # Module-level attributes start with "Attribute VB_"
    # They appear at the start of the file
    last_attr_idx = -1

    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("Attribute VB_"):
            last_attr_idx = i
        elif last_attr_idx >= 0 and not stripped.startswith("Attribute VB_"):
            # We've found the end of the module-level attributes
            break

    if last_attr_idx == -1:
        # No module-level attributes found, treat everything as code
        return "", content

    # Split at the last module-level attribute
    header = "\n".join(lines[: last_attr_idx + 1])
    code = "\n".join(lines[last_attr_idx + 1 :])

    return header.strip(), code.strip()


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
            f"'*-vba import' or by saving the file again in the next edit session.\n"
            f"As of version 0.2.1, the '*-vba edit' command will no longer overwrite files\n"
            f"already present in the VBA directory."
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

    # In OfficeVBAHandler.handle_document_module:
    def handle_document_module(self, component: Any, content: str, temp_file: Path) -> None:
        """Handle the special document module."""
        try:
            # No header stripping during import as it was already stripped during export
            actual_code = content

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

            # Track existing files
            last_known_files = set(path.name for path in self.vba_dir.glob("[!~$]*.bas"))
            last_known_files.update(path.name for path in self.vba_dir.glob("[!~$]*.cls"))
            last_known_files.update(path.name for path in self.vba_dir.glob("[!~$]*.frm"))

            # Setup the file watcher
            watcher = RegExpWatcher(self.vba_dir, re_files=r"^.*(\.cls|\.frm|\.bas)$")

            while True:
                # Always check connection if interval has elapsed
                current_time = time.time()
                if current_time - last_check_time >= check_interval:
                    if not self.is_document_open():
                        raise DocumentClosedError("document")
                    last_check_time = current_time

                # Get current files
                current_files = set(path.name for path in self.vba_dir.glob("[!~$]*.bas"))
                current_files.update(path.name for path in self.vba_dir.glob("[!~$]*.cls"))
                current_files.update(path.name for path in self.vba_dir.glob("[!~$]*.frm"))

                # Check for deleted files
                deleted_files = last_known_files - current_files
                for deleted_file in deleted_files:
                    try:
                        module_name = os.path.splitext(deleted_file)[0]
                        vb_component = self.get_vba_project().VBComponents(module_name)
                        self.get_vba_project().VBComponents.Remove(vb_component)
                        logger.info(f"Deleted module: {module_name}")
                    except Exception as e:
                        logger.error(f"Failed to delete module {module_name}: {str(e)}")

                # Update last known files
                last_known_files = current_files

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

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA content from the Word document.

        Args:
            save_metadata: Whether to save metadata about the export
            overwrite: Whether to overwrite existing files. If False, only exports files
                    that don't exist yet.
        """
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

                    # Skip if file exists and we're not overwriting
                    if not overwrite and final_file.exists():
                        logger.debug(f"Skipping existing file: {final_file}")
                        continue

                    logger.debug(f"Exporting component {component.Name} to {final_file}")

                    # Export to temporary file
                    component.Export(str(temp_file))

                    # Read with specified encoding and write as UTF-8
                    with open(temp_file, "r", encoding=self.encoding) as source:
                        if component.Type == 100:  # Document module
                            content = "".join(source.readlines()[9:])  # Strip header during export
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

    def get_component_info(self, component: Any) -> dict:
        """Get detailed information about a VBA component.

        Args:
            component: A VBA component object

        Returns:
            Dictionary containing component metadata including correct type and extension
        """
        try:
            # Get code line count safely
            code_lines = component.CodeModule.CountOfLines if hasattr(component, "CodeModule") else 0

            # Map component types to metadata
            type_info = {
                1: {  # Standard Module
                    "type_name": "Standard Module",
                    "extension": ".bas",
                    "cls_header": False,
                },
                2: {  # Class Module
                    "type_name": "Class Module",
                    "extension": ".cls",
                    "cls_header": True,
                },
                3: {  # MSForm
                    "type_name": "UserForm",
                    "extension": ".frm",
                    "cls_header": True,
                },
                100: {  # Document Module
                    "type_name": "Document Module",
                    "extension": ".cls",
                    "cls_header": True,
                },
            }

            # Get type info or use defaults for unknown types
            type_data = type_info.get(
                component.Type, {"type_name": "Unknown", "extension": ".txt", "cls_header": False}
            )

            return {
                "name": component.Name,
                "type": component.Type,
                "type_name": type_data["type_name"],
                "extension": type_data["extension"],
                "code_lines": code_lines,
                "has_cls_header": type_data["cls_header"],
            }
        except Exception as e:
            logger.error(f"Failed to get component info for {component.Name}: {str(e)}")
            raise VBAError(f"Failed to analyze component {component.Name}") from e

    def export_component(self, component: Any, directory: Path) -> None:
        """Export a VBA component as separate header and code files.

        Args:
            component: VBA component to export
            directory: Target directory for export

        Raises:
            VBAError: If export fails
        """
        temp_file = None
        try:
            name = component.Name
            temp_file = directory / f"{name}.tmp"

            # Special handling for forms - export .frx first
            if component.Type == self.VBEXT_CT_MSFORM:
                frx_source = Path(self.doc.FullName).parent / f"{name}.frx"
                if frx_source.exists():
                    frx_target = directory / f"{name}.frx"
                    try:
                        shutil.copy2(frx_source, frx_target)
                        logger.debug(f"Exported form binary: {frx_target}")
                    except (OSError, shutil.Error) as e:
                        logger.error(f"Failed to copy form binary {name}.frx: {e}")
                        raise VBAError(f"Failed to export form binary {name}.frx") from e

            # Export component to temp file
            try:
                component.Export(str(temp_file))
            except Exception as e:
                raise VBAError(f"Failed to export component {name}") from e

            # Read content with specified encoding
            try:
                with open(temp_file, "r", encoding=self.encoding) as f:
                    content = f.read()
            except UnicodeError as e:
                logger.error(f"Encoding error reading {name}: {e}")
                raise VBAError(f"Failed to read {name} with encoding {self.encoding}") from e
            except OSError as e:
                logger.error(f"IO error reading {name}: {e}")
                raise VBAError(f"Failed to read exported component {name}") from e

            # Split content
            header, code = split_vba_content(content)

            # Validate split content
            if not header and component.Type != self.VBEXT_CT_STMODULE:
                logger.warning(f"No header found for non-standard module {name}")

            # Write header and code files
            header_file = directory / f"{name}.header"
            code_file = directory / f"{name}{self.type_to_ext[component.Type]}"

            try:
                # Write header if we have one
                if header:
                    with open(header_file, "w", encoding="utf-8") as f:
                        f.write(header + "\n")

                # Always write code file
                with open(code_file, "w", encoding="utf-8") as f:
                    f.write(code + "\n")

            except OSError as e:
                logger.error(f"Failed to write files for {name}: {e}")
                raise VBAError(f"Failed to write component files for {name}") from e

            logger.info(f"Exported: {code_file.name}")

        except VBAError:
            raise
        except Exception as e:
            logger.error(f"Unexpected error exporting {component.Name}: {e}")
            raise VBAError(f"Failed to export component {component.Name}") from e
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError:
                    pass

    def import_component(self, file_path: Path, components: Any) -> None:
        """Import a VBA component with proper header handling.

        Args:
            file_path: Path to the code file
            components: VBA components collection

        Raises:
            VBAError: If import fails
        """
        temp_file = None
        try:
            name = file_path.stem
            header_file = file_path.with_suffix(".header")

            # Handle form binaries first
            if file_path.suffix.lower() == ".frm":
                frx_source = file_path.with_suffix(".frx")
                if frx_source.exists():
                    frx_target = Path(self.doc.FullName).parent / f"{name}.frx"
                    try:
                        shutil.copy2(frx_source, frx_target)
                        logger.debug(f"Imported form binary: {frx_target}")
                    except (OSError, shutil.Error) as e:
                        logger.error(f"Failed to copy form binary {name}.frx: {e}")
                        raise VBAError(f"Failed to import form binary {name}.frx") from e

            # Read header if it exists
            header = ""
            if header_file.exists():
                try:
                    with open(header_file, "r", encoding="utf-8") as f:
                        header = f.read().strip()
                except OSError as e:
                    logger.error(f"Failed to read header for {name}: {e}")
                    raise VBAError(f"Failed to read header file for {name}") from e

            # Read code
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    code = f.read().strip()
            except OSError as e:
                logger.error(f"Failed to read code for {name}: {e}")
                raise VBAError(f"Failed to read code file for {name}") from e

            # For .cls files, determine if this is a document module
            if file_path.suffix.lower() == ".cls" and header:
                # Check if this is a document module by examining attributes
                if determine_cls_type(header) == VBAModuleType.DOCUMENT:
                    try:
                        # Get existing document module
                        doc_component = components(name)
                        # Clear existing code
                        if doc_component.CodeModule.CountOfLines > 0:
                            doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)
                        # Add new code
                        if code.strip():
                            doc_component.CodeModule.AddFromString(code)
                        logger.info(f"Updated document module: {file_path.name}")
                        return
                    except Exception as e:
                        logger.error(f"Failed to update document module {name}: {e}")
                        raise VBAError(f"Failed to update document module {name}") from e

            # For standard modules without header, create minimal header
            if not header and file_path.suffix.lower() == ".bas":
                header = f'Attribute VB_Name = "{name}"'

            # Combine content
            content = f"{header}\n\n{code}\n" if header else f"{code}\n"

            # Create temporary file for import
            temp_file = file_path.with_suffix(".tmp")
            try:
                with open(temp_file, "w", encoding=self.encoding) as f:
                    f.write(content)
            except OSError as e:
                logger.error(f"Failed to create temp file for {name}: {e}")
                raise VBAError(f"Failed to create temporary file for {name}") from e

            # Remove existing component if it exists (except document modules)
            try:
                existing = components(name)
                components.Remove(existing)
                logger.debug(f"Removed existing component: {name}")
            except Exception:
                logger.debug(f"No existing component to remove: {name}")

            try:
                # Import the component
                components.Import(str(temp_file))
                logger.info(f"Imported: {file_path.name}")
            except Exception as e:
                raise VBAError(f"Failed to import component {name}") from e

        except VBAError:
            raise
        except Exception as e:
            logger.error(f"Unexpected error importing {file_path}: {e}")
            raise VBAError(f"Failed to import {file_path.name}") from e
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError:
                    pass

    def get_document_module_name(self) -> str:
        return "ThisWorkbook"

    def is_document_open(self) -> bool:
        try:
            if self.doc is None:
                return False

            # Try to access workbook name
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
        """Import a single VBA file that has changed.

        Args:
            file_path: Path to the changed VBA file

        Raises:
            DocumentClosedError: If workbook is closed
            RPCError: If Excel is not responding
            VBAError: If import fails
        """
        logger.info(f"Processing changes in {file_path.name}")

        try:
            # Check if workbook is still open
            if not self.is_document_open():
                raise DocumentClosedError("workbook")

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            # Import the component
            self.import_component(file_path, components)

            # Save after successful import
            self.doc.Save()

        except (DocumentClosedError, RPCError):
            raise
        except Exception as e:
            logger.error(f"Failed to process {file_path.name}: {str(e)}")
            raise VBAError(f"Failed to import {file_path.name}") from e

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and update the workbook."""
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 30  # Check connection every 30 seconds

            # Track existing files
            last_known_files = set()
            for ext in [".cls", ".bas", ".frm"]:
                last_known_files.update(path.name for path in self.vba_dir.glob(f"[!~$]*{ext}"))

            # Setup file watcher
            watcher = RegExpWatcher(self.vba_dir, re_files=r"^.*\.(cls|frm|bas)$")

            while True:
                # Check connection periodically
                current_time = time.time()
                if current_time - last_check_time >= check_interval:
                    if not self.is_document_open():
                        raise DocumentClosedError("workbook")
                    last_check_time = current_time

                # Get current files
                current_files = set()
                for ext in [".cls", ".bas", ".frm"]:
                    current_files.update(path.name for path in self.vba_dir.glob(f"[!~$]*{ext}"))

                # Handle deleted files
                deleted_files = last_known_files - current_files
                for deleted_file in deleted_files:
                    try:
                        module_name = Path(deleted_file).stem
                        vb_component = self.get_vba_project().VBComponents(module_name)
                        self.get_vba_project().VBComponents.Remove(vb_component)
                        logger.info(f"Deleted module: {module_name}")
                    except Exception as e:
                        logger.error(f"Failed to delete module {module_name}: {str(e)}")

                # Update tracked files
                last_known_files = current_files

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

                # Prevent excessive CPU usage
                time.sleep(0.8)

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        except (DocumentClosedError, RPCError) as e:
            raise e
        finally:
            logger.info("VBA editor stopped.")

    def import_vba(self) -> None:
        """Import VBA content into the Excel workbook with type preservation."""
        try:
            # First check if document is accessible
            if self.doc is None:
                self.open_document()
            _ = self.doc.Name  # Check connection

            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            # Find all VBA files
            vba_files = []
            for ext in [".cls", ".bas", ".frm"]:
                vba_files.extend(self.vba_dir.glob(f"*{ext}"))

            if not vba_files:
                logger.info("No VBA files found to import.")
                return

            logger.info(f"\nFound {len(vba_files)} VBA files to import:")
            for vba_file in vba_files:
                logger.info(f"  - {vba_file.name}")

            # Import components
            for vba_file in vba_files:
                try:
                    self.import_component(vba_file, components)
                except Exception as e:
                    logger.error(f"Failed to import {vba_file.name}: {str(e)}")
                    continue

            # Only try to save if we successfully imported files
            self.save_document()

        except Exception as e:
            if check_rpc_error(e):
                raise DocumentClosedError("workbook")
            raise VBAError(str(e))

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA content from the Excel workbook with type preservation.

        Args:
            save_metadata: Whether to save metadata about the export
            overwrite: Whether to overwrite existing files
        """
        try:
            self.open_document()
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            if not components.Count:
                logger.info("No VBA components found in the workbook.")
                return

            # Get and log component information
            component_list = []
            for component in components:
                info = self.get_component_info(component)
                component_list.append(info)

            logger.info(f"\nFound {len(component_list)} VBA components:")
            for comp in component_list:
                logger.info(f"  - {comp['name']} ({comp['type_name']}, {comp['code_lines']} lines)")

            encoding_data = {}

            # Export components
            for component in components:
                try:
                    # Skip if file exists and we're not overwriting
                    info = self.get_component_info(component)
                    final_file = self.vba_dir / f"{info['name']}{info['extension']}"

                    if not overwrite and final_file.exists():
                        if (
                            info["type"] != 100  # Not a document module
                            or (info["type"] == 100 and info["code_lines"] == 0)
                        ):  # Empty document module
                            logger.debug(f"Skipping existing file: {final_file}")
                            continue

                    self.export_component(component, self.vba_dir)
                    encoding_data[info["name"]] = {"encoding": self.encoding, "type": info["type_name"]}

                except Exception as e:
                    logger.error(f"Failed to export component {component.Name}: {str(e)}")
                    continue

            if save_metadata:
                self._save_metadata(encoding_data)

            os.startfile(self.vba_dir)

        except Exception as e:
            error_msg = "Failed to export VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e
        finally:
            self.save_document()
