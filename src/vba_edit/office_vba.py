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
from typing import Dict, Optional, Any, Tuple

# Third-party imports
import win32com.client
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

# Configure module logger
logger = logging.getLogger(__name__)


# Office app configuration
OFFICE_MACRO_EXTENSIONS: Dict[str, str] = {
    "word": ".docm",
    "excel": ".xlsm",
    "access": ".accdb",
    "powerpoint": ".pptm",
    # Potential future support
    # "outlook": ".otm",
    # "project": ".mpp",
    # "visio": ".vsdm",
}

# Command-line entry points for different Office applications
OFFICE_CLI_NAMES = {app: f"{app}-vba" for app in OFFICE_MACRO_EXTENSIONS.keys()}

# Currently supported apps in vba-edit
# "access" is only partially supported and will be included
# in list as soon as tests are adapted to handle it
SUPPORTED_APPS = ["word", "excel"]


class VBADocumentNames:
    """Document module names across different languages."""

    # Excel document module names
    EXCEL_WORKBOOK_NAMES = {
        "ThisWorkbook",  # English
        "DieseArbeitsmappe",  # German
        "CeClasseur",  # French
        "EstaLista",  # Spanish
        "QuestoFoglio",  # Italian
        "EstaLista",  # Portuguese
        "このブック",  # Japanese
        "本工作簿",  # Chinese Simplified
        "本活頁簿",  # Chinese Traditional
        "이통합문서",  # Korean
        "ЭтаКнига",  # Russian
    }

    # Excel worksheet module prefixes
    EXCEL_SHEET_PREFIXES = {
        "Sheet",  # English
        "Tabelle",  # German
        "Feuil",  # French
        "Hoja",  # Spanish
        "Foglio",  # Italian
        "Planilha",  # Portuguese
        "シート",  # Japanese
        "工作表",  # Chinese Simplified/Traditional
        "시트",  # Korean
        "Лист",  # Russian
    }

    # Word document module names
    WORD_DOCUMENT_NAMES = {
        "ThisDocument",  # English
        "DiesesDokument",  # German
        "CeDocument",  # French
        "EsteDocumento",  # Spanish/Portuguese
        "QuestoDocumento",  # Italian
        "この文書",  # Japanese
        "本文檔",  # Chinese Traditional
        "本文档",  # Chinese Simplified
        "이문서",  # Korean
        "ЭтотДокумент",  # Russian
    }

    @classmethod
    def is_document_module(cls, name: str) -> bool:
        """Check if a name matches any known document module name.

        Args:
            name: Name to check

        Returns:
            bool: True if name matches any known document module name
        """
        # Direct match for workbook/document
        if name in cls.EXCEL_WORKBOOK_NAMES or name in cls.WORD_DOCUMENT_NAMES:
            return True

        # Check for sheet names with numbers
        return any(name.startswith(prefix) and name[len(prefix) :].isdigit() for prefix in cls.EXCEL_SHEET_PREFIXES)


# VBA type definitions and constants
class VBAModuleType(Enum):
    """VBA module types"""

    DOCUMENT = auto()  # ThisWorkbook/ThisDocument modules
    CLASS = auto()  # Regular class modules
    STANDARD = auto()  # Standard modules (.bas)
    FORM = auto()  # UserForm modules


class VBATypes:
    """Constants for VBA component types"""

    VBEXT_CT_DOCUMENT = 100  # Document module type
    VBEXT_CT_MSFORM = 3  # UserForm type
    VBEXT_CT_STDMODULE = 1  # Standard module type
    VBEXT_CT_CLASSMODULE = 2  # Class module type

    # Application specific constants
    XL_WORKSHEET = -4167  # xlWorksheet type for Excel

    # Map module types to file extensions and metadata
    TYPE_TO_EXT = {
        VBEXT_CT_STDMODULE: ".bas",  # Standard Module
        VBEXT_CT_CLASSMODULE: ".cls",  # Class Module
        VBEXT_CT_MSFORM: ".frm",  # MSForm
        VBEXT_CT_DOCUMENT: ".cls",  # Document Module
    }

    TYPE_INFO = {
        VBEXT_CT_STDMODULE: {
            "type_name": "Standard Module",
            "extension": ".bas",
            "cls_header": False,
        },
        VBEXT_CT_CLASSMODULE: {
            "type_name": "Class Module",
            "extension": ".cls",
            "cls_header": True,
        },
        VBEXT_CT_MSFORM: {
            "type_name": "UserForm",
            "extension": ".frm",
            "cls_header": True,
        },
        VBEXT_CT_DOCUMENT: {
            "type_name": "Document Module",
            "extension": ".cls",
            "cls_header": True,
        },
    }


# Exception classes
class VBAError(Exception):
    """Base exception class for all VBA-related errors.

    This exception serves as the parent class for all specific VBA error types
    in the module. It provides a common base for error handling and allows
    catching all VBA-related errors with a single except clause.
    """

    # Forms are not supported in Access
    pass


class VBAAccessError(VBAError):
    """Exception raised when access to the VBA project is denied.

    This typically occurs when "Trust access to the VBA project object model"
    is not enabled in the Office application's Trust Center settings.
    """

    pass


class VBAImportError(VBAError):
    """Exception raised when importing VBA components fails.

    This can occur due to various reasons such as invalid file format,
    encoding issues, or problems with the VBA project structure.
    """

    pass


class VBAExportError(VBAError):
    """Exception raised when exporting VBA components fails.

    This can occur due to file system permissions, encoding issues,
    or problems accessing the VBA components.
    """

    pass


class DocumentClosedError(VBAError):
    """Exception raised when attempting to access a closed Office document.

    This exception includes a custom error message that provides guidance
    on how to handle changes made after document closure.

    Args:
        doc_type (str): Type of document (e.g., "workbook", "document")
    """

    def __init__(self, doc_type: str = "document"):
        super().__init__(
            f"\nThe Office {doc_type} has been closed. The edit session will be terminated.\n"
            f"IMPORTANT: Any changes made after closing the {doc_type} must be imported using\n"
            f"'*-vba import' or by saving the file again in the next edit session.\n"
            f"As of version 0.2.1, the '*-vba edit' command will no longer overwrite files\n"
            f"already present in the VBA directory."
        )


class RPCError(VBAError):
    """Exception raised when the RPC server becomes unavailable.

    This typically occurs when the Office application crashes or is forcefully closed.

    Args:
        app_name (str): Name of the Office application
    """

    def __init__(self, app_name: str = "Office application"):
        super().__init__(
            f"\nLost connection to {app_name}. The edit session will be terminated.\n"
            f"IMPORTANT: Any changes made after closing {app_name} must be imported using\n"
            f"'*-vba import' or by saving the file again in the next edit session.\n"
            f"As of version 0.2.1, the '*-vba edit' command will no longer overwrite files\n"
            f"already present in the VBA directory."
        )


def check_rpc_error(error: Exception) -> bool:
    """Check if an exception is related to RPC server unavailability.

    This function examines the error message for common indicators of RPC
    server connection issues.

    Args:
        error: The exception to check

    Returns:
        bool: True if the error appears to be RPC-related, False otherwise
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


class VBAComponentHandler:
    """Handles VBA component operations independent of Office application type.

    This class provides core functionality for managing VBA components, including
    analyzing module types, handling headers, and preparing content for import/export
    operations. It serves as a utility class for the main Office-specific handlers.
    """

    def get_component_info(self, component: Any) -> Dict[str, Any]:
        """Get detailed information about a VBA component.

        Analyzes a VBA component and returns metadata including its type,
        line count, and appropriate file extension.

        Args:
            component: A VBA component object from any Office application

        Returns:
            Dict containing component metadata with the following keys:
                - name: Component name
                - type: VBA type code
                - type_name: Human-readable type name
                - extension: Appropriate file extension
                - code_lines: Number of lines of code
                - has_cls_header: Whether component requires a class header

        Raises:
            VBAError: If component information cannot be retrieved
        """
        try:
            # Get code line count safely
            code_lines = component.CodeModule.CountOfLines if hasattr(component, "CodeModule") else 0

            # Get type info or use defaults for unknown types
            type_data = VBATypes.TYPE_INFO.get(
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

    def determine_cls_type(self, header: str) -> VBAModuleType:
        """Determine if a .cls file is a document module or regular class module.

        Analyzes the VBA component header to determine its exact type based on
        the presence and values of specific attributes.

        Args:
            header: Content of the VBA component header

        Returns:
            VBAModuleType.DOCUMENT or VBAModuleType.CLASS based on header analysis
        """
        # Extract key attributes
        predeclared = re.search(r"Attribute VB_PredeclaredId = (\w+)", header)
        exposed = re.search(r"Attribute VB_Exposed = (\w+)", header)

        # Document modules have both attributes set to True
        if predeclared and exposed and predeclared.group(1).lower() == "true" and exposed.group(1).lower() == "true":
            return VBAModuleType.DOCUMENT

        return VBAModuleType.CLASS

    def get_module_type(self, file_path: Path) -> VBAModuleType:
        """Determine VBA module type from file extension and content.

        Args:
            file_path: Path to the VBA module file

        Returns:
            Appropriate VBAModuleType

        Raises:
            ValueError: If file extension is unknown
        """
        suffix = file_path.suffix.lower()
        name = file_path.stem

        # Check if it's a known document module name in any language
        if VBADocumentNames.is_document_module(name):
            return VBAModuleType.DOCUMENT

        if suffix == ".bas":
            return VBAModuleType.STANDARD
        elif suffix == ".frm":
            return VBAModuleType.FORM
        elif suffix == ".cls":
            # For .cls files, check the header if available
            header_file = file_path.with_suffix(".header")
            if header_file.exists():
                with open(header_file, "r", encoding="utf-8") as f:
                    return self.determine_cls_type(f.read())

            logger.debug(f"No header file found for {file_path}, treating as regular class module")
            return VBAModuleType.CLASS

        raise ValueError(f"Unknown file extension: {suffix}")

    def split_vba_content(self, content: str) -> Tuple[str, str]:
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
        last_attr_idx = -1

        for i, line in enumerate(lines):
            stripped = line.strip()
            if stripped.startswith("Attribute VB_"):
                last_attr_idx = i
            elif last_attr_idx >= 0 and not stripped.startswith("Attribute VB_"):
                break

        if last_attr_idx == -1:
            return "", content

        header = "\n".join(lines[: last_attr_idx + 1])
        code = "\n".join(lines[last_attr_idx + 1 :])

        return header.strip(), code.strip()

    def create_minimal_header(self, name: str, module_type: VBAModuleType) -> str:
        """Create a minimal header for a VBA component.

        Args:
            name: Name of the VBA component
            module_type: Type of the VBA module

        Returns:
            Minimal valid header for the component type
        """
        if module_type == VBAModuleType.CLASS:
            # Class modules need the class declaration and standard attributes
            header = [
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                f'Attribute VB_Name = "{name}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = False",
                "Attribute VB_Exposed = False",
            ]
        elif module_type == VBAModuleType.FORM:
            # UserForm requires specific form structure and GUID
            # {C62A69F0-16DC-11CE-9E98-00AA00574A4F} is the standard UserForm GUID
            header = [
                "VERSION 5.00",
                "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + name,
                f'   Caption         =   "{name}"',
                "   ClientHeight    =   3000",
                "   ClientLeft      =   100",
                "   ClientTop       =   400",
                "   ClientWidth     =   4000",
                '   OleObjectBlob   =   "' + name + '.frx":0000',
                "   StartUpPosition =   1  'CenterOwner",
                "End",
                f'Attribute VB_Name = "{name}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = False",
            ]
            logger.info(
                f"Created minimal header for UserForm: {name} \n"
                "Consider using the command-line option --save-headers "
                "in order not to lose previously specified form structure and GUID."
            )
        else:
            # Standard modules only need the name
            header = [f'Attribute VB_Name = "{name}"']

        return "\n".join(header)

    def prepare_import_content(self, name: str, module_type: VBAModuleType, header: str, code: str) -> str:
        """Prepare content for VBA component import.

        Args:
            name: Name of the VBA component
            module_type: Type of the VBA module
            header: Header content (may be empty)
            code: Code content

        Returns:
            Properly formatted content for import
        """
        if not header and module_type == VBAModuleType.STANDARD:
            header = self.create_minimal_header(name, module_type)

        return f"{header}\n{code}\n" if header else f"{code}\n"

    def validate_component_header(self, header: str, expected_type: VBAModuleType) -> bool:
        """Validate that a component's header matches its expected type.

        Args:
            header: Header content to validate
            expected_type: Expected module type

        Returns:
            True if header is valid for the expected type
        """
        if not header:
            return expected_type == VBAModuleType.STANDARD

        actual_type = self.determine_cls_type(header)

        if expected_type == VBAModuleType.DOCUMENT:
            return actual_type == VBAModuleType.DOCUMENT

        return True  # Other types are less strict about headers


class OfficeVBAHandler(ABC):
    """Abstract base class for handling VBA operations across Office applications.

    This class provides the foundation for application-specific VBA handlers,
    implementing common functionality while requiring specific implementations
    for application-dependent operations.

    Args:
        doc_path (str): Path to the Office document
        vba_dir (Optional[str]): Directory for VBA files (defaults to current directory)
        encoding (str): Character encoding for VBA files (default: cp1252)
        verbose (bool): Enable verbose logging
        save_headers (bool): Whether to save VBA component headers to separate files

    Attributes:
        doc_path (Path): Resolved path to the Office document
        vba_dir (Path): Resolved path to VBA directory
        encoding (str): Character encoding for file operations
        verbose (bool): Verbose logging flag
        save_headers (bool): Header saving flag
        app: Office application COM object
        doc: Office document COM object
        component_handler (VBAComponentHandler): Utility handler for VBA components
    """

    def __init__(
        self,
        doc_path: str,
        vba_dir: Optional[str] = None,
        encoding: str = "cp1252",
        verbose: bool = False,
        save_headers: bool = False,
    ):
        """Initialize the VBA handler.

        Args:
            doc_path (str): Path to the Office document
            vba_dir (Optional[str]): Directory for VBA files (defaults to current directory)
            encoding (str): Character encoding for VBA files (default: cp1252)
            verbose (bool): Enable verbose logging
            save_headers (bool): Whether to save VBA component headers to separate files
        """
        self.doc_path = Path(doc_path).resolve()
        self.vba_dir = Path(vba_dir).resolve() if vba_dir else Path.cwd()
        self.encoding = encoding
        self.verbose = verbose
        self.save_headers = save_headers
        self.app = None
        self.doc = None
        self.component_handler = VBAComponentHandler()

        # Configure logging
        log_level = logging.DEBUG if verbose else logging.INFO
        logger.setLevel(log_level)

        logger.debug(f"Initialized {self.__class__.__name__} with document: {doc_path}")
        logger.debug(f"VBA directory: {self.vba_dir}")
        logger.debug(f"Using encoding: {encoding}")
        logger.debug(f"Save headers: {save_headers}")

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

    @property
    def document_type(self) -> str:
        """Get the document type string for error messages."""
        return "workbook" if self.app_name == "Excel" else "document"

    def get_vba_project(self) -> Any:
        """Get the VBA project from the document.

        Returns:
            The VBA project object

        Raises:
            VBAAccessError: If VBA project access is denied
            VBAError: For other VBA-related errors
        """
        try:
            project = self.doc.VBProject
            # Verify access by attempting to get components
            _ = project.VBComponents
            return project
        except Exception as e:
            error_msg = (
                "Cannot access VBA project. Please ensure 'Trust access to the VBA "
                "project object model' is enabled in Trust Center Settings."
            )
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAAccessError(error_msg) from e

    @abstractmethod
    def get_document_module_name(self) -> str:
        """Get the name of the document module (e.g., ThisDocument, ThisWorkbook)."""
        pass

    def is_document_open(self) -> bool:
        """Check if the document is still open and accessible."""
        try:
            if self.doc is None:
                return False

            # Try to access document name
            name = self.doc.Name
            if callable(name):  # Handle Mock case in tests
                name = name()

            # Check if document is still active
            return self.doc.FullName == str(self.doc_path)

        except Exception as e:
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            raise DocumentClosedError(self.document_type)

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

    def _check_form_safety(self, vba_dir: Path) -> None:
        """Check if there are .frm files when headers are disabled.

        Args:
            vba_dir: Directory to check for .frm files

        Raises:
            VBAError: If .frm files are found and save_headers is False
        """
        if not self.save_headers:
            form_files = list(vba_dir.glob("*.frm"))
            if form_files:
                form_names = ", ".join(f.stem for f in form_files)
                error_msg = (
                    f"\nERROR: Found UserForm files ({form_names}) but --save-headers is not enabled!\n"
                    f"UserForms require their full header information to maintain form specifications.\n"
                    f"Please re-run the command with the --save-headers flag to preserve form settings."
                )
                logger.error(error_msg)
                sys.exit(1)

    def open_document(self) -> None:
        """Open the Office document."""
        try:
            if self.doc is None:
                self.initialize_app()
                logger.debug(f"Opening document: {self.doc_path}")
                self.doc = self._open_document_impl()
        except Exception as e:
            error_msg = f"Failed to open document: {self.doc_path}"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    @abstractmethod
    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        pass

    def save_document(self) -> None:
        """Save the document if it's open."""
        if self.doc is not None:
            try:
                self.doc.Save()
                logger.info("Document has been saved and left open for further editing")
            except Exception as e:
                raise VBAError("Failed to save document") from e

    def _save_metadata(self, encodings: Dict[str, Dict[str, Any]]) -> None:
        """Save metadata including encoding information.

        Args:
            encodings: Dictionary mapping module names to their encoding information

        Raises:
            VBAError: If metadata cannot be saved
        """
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

    def export_component(self, component: Any, directory: Path) -> None:
        """Export a VBA component as separate header and code files.

        Args:
            component: VBA component to export
            directory: Target directory for export
        """
        temp_file = None
        try:
            info = self.component_handler.get_component_info(component)
            name = info["name"]
            temp_file = directory / f"{name}.tmp"

            # Handle form binary if needed
            if info["type"] == VBATypes.VBEXT_CT_MSFORM:
                self._handle_form_binary_export(name)

            # Export to temp file
            component.Export(str(temp_file))

            # Read and process content
            with open(temp_file, "r", encoding=self.encoding) as f:
                content = f.read()

            # Split content and handle document module special case
            header, code = self.component_handler.split_vba_content(content)
            #            if info["type"] == VBATypes.VBEXT_CT_DOCUMENT:
            #               header = ""  # Strip header for document modules during export

            # Write files
            self._write_component_files(name, header, code, info, directory)

            logger.info(f"Exported: {name}")

        except Exception as e:
            logger.error(f"Failed to export component {component.Name}: {str(e)}")
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
        """
        temp_file = None
        try:
            name = file_path.stem
            module_type = self.component_handler.get_module_type(file_path)

            # Handle form binaries for UserForms
            if module_type == VBAModuleType.FORM:
                self._handle_form_binary_import(name)

            # Read content
            header = self._read_header_file(file_path)
            code = self._read_code_file(file_path)

            # Special handling for document modules
            if module_type == VBAModuleType.DOCUMENT:
                self._update_document_module(name, code, components)
                return

            # Get minimal header information for class modules and forms if no header provided
            if not header and module_type in [VBAModuleType.CLASS, VBAModuleType.FORM]:
                header = self.component_handler.create_minimal_header(name, module_type)
                logger.debug(f"Created minimal header for {name}")

            # Prepare content for import
            content = self.component_handler.prepare_import_content(name, module_type, header, code)

            # Create temp file for import
            temp_file = file_path.with_suffix(".tmp")
            with open(temp_file, "w", encoding=self.encoding) as f:
                f.write(content)

            # Remove existing component if present
            try:
                existing = components(name)
                components.Remove(existing)
                logger.debug(f"Removed existing component: {name}")
            except Exception:
                logger.debug(f"No existing component to remove: {name}")

            # Import component
            components.Import(str(temp_file))
            logger.info(f"Imported: {file_path.name}")

        except Exception as e:
            logger.error(f"Failed to import {file_path.name}: {str(e)}")
            raise VBAError(f"Failed to import {file_path.name}") from e
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError:
                    pass

    @abstractmethod
    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary (.frx) export."""
        pass

    @abstractmethod
    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary (.frx) import."""
        pass

    @abstractmethod
    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module."""
        pass

    def _read_header_file(self, code_file: Path) -> str:
        """Read the header file if it exists."""
        header_file = code_file.with_suffix(".header")
        if header_file.exists():
            with open(header_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        return ""

    def _read_code_file(self, code_file: Path) -> str:
        """Read the code file."""
        with open(code_file, "r", encoding="utf-8") as f:
            return f.read().strip()

    def _write_component_files(self, name: str, header: str, code: str, info: Dict[str, Any], directory: Path) -> None:
        """Write component files with proper encoding.

        Args:
            name: Name of the VBA component
            header: Header content (may be empty)
            code: Code content
            info: Component information dictionary
            directory: Target directory
        """
        # Save header if enabled and header content exists
        if self.save_headers and header:
            header_file = directory / f"{name}.header"
            with open(header_file, "w", encoding="utf-8") as f:
                f.write(header + "\n")
            logger.debug(f"Saved header file: {header_file}")

        # Always save code file
        code_file = directory / f"{name}{info['extension']}"
        with open(code_file, "w", encoding="utf-8") as f:
            f.write(code + "\n")
        logger.debug(f"Saved code file: {code_file}")

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and update the document."""
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")
            last_check_time = time.time()
            check_interval = 30  # Check connection every 30 seconds

            # Setup file watcher
            watcher = RegExpWatcher(
                self.vba_dir,
                re_files=r"^.*\.(cls|frm|bas)$",
            )

            while True:
                try:
                    # Check connection periodically
                    current_time = time.time()
                    if current_time - last_check_time >= check_interval:
                        if not self.is_document_open():
                            raise DocumentClosedError(self.document_type)
                        last_check_time = current_time
                        logger.debug("Connection check passed")

                    # Check for changes using watchgod
                    changes = watcher.check()
                    if changes:
                        logger.debug(f"Watchgod detected changes: {changes}")

                    for change_type, path in changes:
                        try:
                            path = Path(path)
                            if change_type == Change.deleted:
                                # Handle deleted files
                                logger.info(f"Detected deletion of {path.name}")
                                if not self.is_document_open():
                                    raise DocumentClosedError(self.document_type)

                                vba_project = self.get_vba_project()
                                components = vba_project.VBComponents
                                try:
                                    component = components(path.stem)
                                    components.Remove(component)
                                    logger.info(f"Removed component: {path.stem}")
                                    self.doc.Save()
                                except Exception:
                                    logger.debug(f"Component {path.stem} already removed or not found")

                            elif change_type in (Change.added, Change.modified):
                                # Handle both added and modified files the same way
                                action = "addition" if change_type == Change.added else "modification"
                                logger.debug(f"Processing {action} in {path}")
                                self.import_single_file(path)

                        except (DocumentClosedError, RPCError) as e:
                            raise e
                        except Exception as e:
                            logger.warning(f"Error handling changes (will retry): {str(e)}")
                            continue

                except (DocumentClosedError, RPCError) as error:
                    raise error
                except Exception as error:
                    logger.warning(f"Error in watch loop (will continue): {str(error)}")

                # Prevent excessive CPU usage but stay responsive
                time.sleep(0.5)

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        except (DocumentClosedError, RPCError) as error:
            raise error
        finally:
            logger.info("VBA editor stopped.")

    def import_vba(self) -> None:
        """Import VBA content into the Office document."""
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

            # Save if we successfully imported files
            self.save_document()

        except Exception as e:
            if check_rpc_error(e):
                raise DocumentClosedError(self.document_type)
            raise VBAError(str(e))

    def import_single_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed.

        Args:
            file_path: Path to the changed VBA file
        """
        logger.info(f"Processing changes in {file_path.name}")

        try:
            # Check if document is still open
            if not self.is_document_open():
                raise DocumentClosedError(self.document_type)

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

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Export VBA content from the Office document.

        Args:
            save_metadata: Whether to save metadata about the export
            overwrite: Whether to overwrite existing files

        Raises:
            VBAError: If operation fails or if UserForms found without headers enabled
        """
        try:
            self.open_document()
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            if not components.Count:
                logger.info(f"No VBA components found in the {self.document_type}.")
                return

            # Get and log component information
            component_list = []
            for component in components:
                info = self.component_handler.get_component_info(component)
                component_list.append(info)

            logger.info(f"\nFound {len(component_list)} VBA components:")
            for comp in component_list:
                logger.info(f"  - {comp['name']} ({comp['type_name']}, {comp['code_lines']} lines)")

            encoding_data = {}

            # Export components
            for component in components:
                try:
                    # Skip if file exists and we're not overwriting
                    info = self.component_handler.get_component_info(component)
                    final_file = self.vba_dir / f"{info['name']}{info['extension']}"

                    if not overwrite and final_file.exists():
                        if (
                            info["type"] != VBATypes.VBEXT_CT_DOCUMENT  # Not doc module
                            or (
                                info["type"] == VBATypes.VBEXT_CT_DOCUMENT and info["code_lines"] == 0
                            )  # Empty doc module
                        ):
                            logger.debug(f"Skipping existing file: {final_file}")
                            continue

                    self.export_component(component, self.vba_dir)
                    encoding_data[info["name"]] = {"encoding": self.encoding, "type": info["type_name"]}

                except Exception as e:
                    logger.error(f"Failed to export component {component.Name}: {str(e)}")
                    continue

            if save_metadata:
                self._save_metadata(encoding_data)

            # Small delay to ensure all files are written
            time.sleep(0.5)

            # Safety check for forms after export
            if not self.app_name == "Access":
                self._check_form_safety(self.vba_dir)
                logger.debug(f"Running form safety check with save_headers={self.save_headers}")

            # Open directory in explorer
            os.startfile(self.vba_dir)

        except Exception as e:
            error_msg = "Failed to export VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e
        finally:
            self.save_document()


class WordVBAHandler(OfficeVBAHandler):
    """Microsoft Word specific implementation of VBA operations.

    Provides Word-specific implementations of abstract methods from OfficeVBAHandler
    and any additional functionality specific to Word VBA projects.

    The handler manages operations like:
    - Importing/exporting VBA modules
    - Handling UserForm binaries (.frx files)
    - Managing ThisDocument module
    - Monitoring file changes
    """

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Word"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Word.Application"

    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        return "ThisDocument"

    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        return self.app.Documents.Open(str(self.doc_path))

    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary (.frx) export for Word."""
        frx_source = Path(self.doc.FullName).parent / f"{name}.frx"
        if frx_source.exists():
            frx_target = self.vba_dir / f"{name}.frx"
            try:
                shutil.copy2(frx_source, frx_target)
                logger.debug(f"Exported form binary: {frx_target}")
            except (OSError, shutil.Error) as e:
                logger.error(f"Failed to copy form binary {name}.frx: {e}")
                raise VBAError(f"Failed to export form binary {name}.frx") from e

    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary (.frx) import for Word."""
        frx_source = self.vba_dir / f"{name}.frx"
        if frx_source.exists():
            frx_target = Path(self.doc.FullName).parent / f"{name}.frx"
            try:
                shutil.copy2(frx_source, frx_target)
                logger.debug(f"Imported form binary: {frx_target}")
            except (OSError, shutil.Error) as e:
                logger.error(f"Failed to copy form binary {name}.frx: {e}")
                raise VBAError(f"Failed to import form binary {name}.frx") from e

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module for Word."""
        try:
            doc_component = components(name)

            # Clear existing code
            if doc_component.CodeModule.CountOfLines > 0:
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)

            # Add new code
            if code.strip():
                doc_component.CodeModule.AddFromString(code)

            logger.info(f"Updated document module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update document module {name}") from e


class ExcelVBAHandler(OfficeVBAHandler):
    """Microsoft Excel specific implementation of VBA operations.

    Provides Excel-specific implementations of abstract methods from OfficeVBAHandler
    and any additional functionality specific to Excel VBA projects.

    The handler manages operations like:
    - Importing/exporting VBA modules
    - Handling UserForm binaries (.frx files)
    - Managing ThisWorkbook and Sheet modules
    - Monitoring file changes
    """

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Excel"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Excel.Application"

    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        return "ThisWorkbook"

    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening logic."""
        return self.app.Workbooks.Open(str(self.doc_path))

    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary (.frx) export for Excel."""
        frx_source = Path(self.doc.FullName).parent / f"{name}.frx"
        if frx_source.exists():
            frx_target = self.vba_dir / f"{name}.frx"
            try:
                shutil.copy2(frx_source, frx_target)
                logger.debug(f"Exported form binary: {frx_target}")
            except (OSError, shutil.Error) as e:
                logger.error(f"Failed to copy form binary {name}.frx: {e}")
                raise VBAError(f"Failed to export form binary {name}.frx") from e

    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary (.frx) import for Excel."""
        frx_source = self.vba_dir / f"{name}.frx"
        if frx_source.exists():
            frx_target = Path(self.doc.FullName).parent / f"{name}.frx"
            try:
                shutil.copy2(frx_source, frx_target)
                logger.debug(f"Imported form binary: {frx_target}")
            except (OSError, shutil.Error) as e:
                logger.error(f"Failed to copy form binary {name}.frx: {e}")
                raise VBAError(f"Failed to import form binary {name}.frx") from e

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update an existing document module for Excel."""
        try:
            # Handle ThisWorkbook and Sheet modules
            doc_component = components(name)

            # Clear existing code
            if doc_component.CodeModule.CountOfLines > 0:
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)

            # Add new code
            if code.strip():
                doc_component.CodeModule.AddFromString(code)

            logger.info(f"Updated document module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update document module {name}") from e


class AccessVBAHandler(OfficeVBAHandler):
    """Microsoft Access specific implementation of VBA operations.

    Provides Access-specific implementations of abstract methods from OfficeVBAHandler.
    Currently only supports Standard (.bas) and Class (.cls) modules.

    The handler manages operations like:
    - Importing/exporting VBA modules
    - Managing module types specific to Access
    - Monitoring file changes
    """

    def __init__(
        self,
        doc_path: Path,
        vba_dir: Optional[Path] = None,
        encoding: str = "cp1252",
        save_headers: bool = True,
        verbose: bool = False,
    ) -> None:
        """Initialize Access VBA handler.

        Args:
            doc_path: Path to Access database
            vba_dir: Directory for VBA files
            encoding: Character encoding for VBA files
            save_headers: Whether to save VBA component headers
            verbose: Enable verbose logging
        """
        super().__init__(doc_path, vba_dir, encoding, save_headers, verbose)
        self._database_open = False

    @property
    def app_name(self) -> str:
        """Name of the Office application."""
        return "Access"

    @property
    def app_progid(self) -> str:
        """ProgID for COM automation."""
        return "Access.Application"

    @property
    def document_type(self) -> str:
        return "database"

    def _open_document_impl(self) -> Any:
        """Implementation-specific document opening for Access."""
        try:
            logger.debug(f"Opening Access database: {self.doc_path}")
            self.app.OpenCurrentDatabase(str(self.doc_path))
            self._database_open = True
            return self.app.CurrentDb()
        except Exception as e:
            self._database_open = False
            error_msg = f"Failed to open database: {self.doc_path}"
            logger.error(f"{error_msg}: {str(e)}")
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            raise VBAError(error_msg) from e

    def save_document(self) -> None:
        """Save changes while keeping database open."""
        if self.doc and self.app and self._database_open:
            if not self._ensure_connection():
                raise VBAError("Cannot verify database state - connection lost")
            logger.info("Database connection verified")
            try:
                logger.debug("Saving Access database")
                # Access doesn't need explicit save - changes are auto-committed
                # Just verify connection is still active
                _ = self.app.CurrentDb()
                logger.info("Database connection verified")
            except Exception as e:
                error_msg = "Failed to save database"
                logger.error(f"{error_msg}: {str(e)}")
                if check_rpc_error(e):
                    raise RPCError(self.app_name)
                raise VBAError(error_msg) from e

    def close_document(self) -> None:
        """Close database explicitly."""
        if self.app and self._database_open:
            try:
                logger.debug("Closing Access database")
                self.app.CloseCurrentDatabase()
                self._database_open = False
            except Exception as e:
                logger.error(f"Error closing database: {str(e)}")

    def __del__(self):
        """Cleanup handler."""
        try:
            self.close_document()
        except Exception as e:
            logger.error(f"Error during cleanup: {str(e)}")

    def get_vba_project(self) -> Any:
        """Get VBA project for Access database."""
        try:
            vba_project = self.app.VBE.ActiveVBProject
            if vba_project is None:
                raise VBAAccessError(
                    "Cannot access VBA project. Please ensure 'Trust access to the VBA "
                    "project object model' is enabled in Trust Center Settings."
                )
            return vba_project
        except Exception as e:
            logger.error(f"Failed to access VBA project: {str(e)}")
            if check_rpc_error(e):
                raise RPCError(self.app_name)
            raise VBAAccessError(
                "Cannot access VBA project. Please ensure 'Trust access to the VBA "
                "project object model' is enabled in Trust Center Settings."
            ) from e

    def get_document_module_name(self) -> str:
        """Get the name of the document module."""
        """
        Access does not have an equivalent document module.
        """
        return ""

    def initialize_app(self) -> None:
        """Initialize the Access application."""
        try:
            if self.app is None:
                logger.debug("Initializing Access application")
                self.app = win32com.client.Dispatch(self.app_progid)
                # Remove Visible=True setting
        except Exception as e:
            error_msg = "Failed to initialize Access application"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def _handle_form_binary_export(self, name: str) -> None:
        """Handle form binary export for Access."""
        # Not needed as forms are not supported
        pass

    def _handle_form_binary_import(self, name: str) -> None:
        """Handle form binary import for Access."""
        # Forms are not supported in Access, so this method is not needed.
        pass

    def _update_document_module(self, name: str, code: str, components: Any) -> None:
        """Update a module in Access."""
        try:
            # Get existing module
            module = components(name)

            # Clear existing code
            if module.CodeModule.CountOfLines > 0:
                module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines)

            # Add new code
            if code.strip():
                module.CodeModule.AddFromString(code)

            logger.info(f"Updated module: {name}")

        except Exception as e:
            raise VBAError(f"Failed to update module {name}") from e

    def export_vba(self, save_metadata: bool = False, overwrite: bool = True) -> None:
        """Override to maintain database connection after export."""
        super().export_vba(save_metadata, overwrite)
        # Don't close database after export
        logger.info("Database remains open for editing")

    def _ensure_connection(self) -> bool:
        """Ensure database connection is active."""
        try:
            if not self._database_open or self.app is None:
                self._open_document_impl()
            self._vba_project = self.app.VBE.ActiveVBProject
            return True
        except Exception as e:
            logger.error(f"Lost database connection: {str(e)}")
            return False

    def update_vba_component(self, name: str, content: str) -> None:
        """Update VBA component with connection recovery."""
        try:
            if not self._ensure_connection():
                raise VBAError("Cannot update VBA - database connection lost")

            component = self._vba_project.VBComponents(name)
            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(content)
            logger.info(f"Updated VBA component: {name}")

        except Exception as e:
            logger.error(f"Failed to update VBA component {name}: {str(e)}")
            raise VBAError(f"Failed to update VBA component {name}") from e
