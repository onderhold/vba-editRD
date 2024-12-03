from abc import ABC, abstractmethod
import datetime
import json
import logging
import os
from pathlib import Path
from typing import Dict, Optional, Any, List
import win32com.client
from watchgod import Change, RegExpWatcher, watch

# Configure logging
logger = logging.getLogger(__name__)

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
        """Get the name of the document module (e.g., 'ThisDocument' for Word)."""
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
                self.doc = self.app.Documents.Open(str(self.doc_path))
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
                error_msg = "Failed to save document"
                logger.error(f"{error_msg}: {str(e)}")
                raise VBAError(error_msg) from e

    def handle_document_module(self, component: Any, content: str, temp_file: Path) -> None:
        """Handle the special document module (ThisDocument, ThisWorkbook, etc.)."""
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
                    "extension": self.type_to_ext.get(component.Type, "unknown")
                }
                component_list.append(component_info)
            
            return component_list
        except Exception as e:
            error_msg = "Failed to get component list"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAError(error_msg) from e

    def export_vba(self, save_metadata: bool = False) -> None:
        """Export VBA content from the Office document."""
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

    def import_vba(self) -> None:
        """Import VBA content into the Office document."""
        try:
            self.open_document()
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
                try:
                    logger.debug(f"Processing {vba_file.name}")
                    with open(vba_file, "r", encoding="utf-8") as f:
                        content = f.read()

                    component_name = vba_file.stem
                    temp_file = vba_file.with_suffix(".temp")

                    if component_name == self.get_document_module_name():
                        # Handle document module
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

                except Exception as e:
                    error_msg = f"Failed to import {vba_file.name}"
                    logger.error(f"{error_msg}: {str(e)}")
                    if temp_file.exists():
                        temp_file.unlink()
                    continue

        except Exception as e:
            error_msg = "Failed to import VBA content"
            logger.error(f"{error_msg}: {str(e)}")
            raise VBAImportError(error_msg) from e
        finally:
            self.save_document()

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and automatically reimport them."""
        try:
            logger.info(f"Watching for changes in {self.vba_dir}...")

            for changes in watch(
                self.vba_dir,
                watcher_cls=RegExpWatcher,
                watcher_kwargs=dict(re_files=r"^.*(\.cls|\.frm|\.bas)$"),
                normal_sleep=400,
            ):
                for change_type, path in changes:
                    if change_type == Change.modified:
                        try:
                            logger.debug(f"Detected change in {path}")
                            self.import_vba()
                        except Exception as e:
                            logger.error(f"Error handling changes: {str(e)}")

        except KeyboardInterrupt:
            logger.info("\nStopping VBA editor...")
        finally:
            logger.info("VBA editor stopped. Your changes have been saved.")

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