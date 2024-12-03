from abc import ABC, abstractmethod
import datetime
import json
import os
from pathlib import Path
from typing import Dict, Optional, Any
import win32com.client
from watchgod import Change, RegExpWatcher, watch


class OfficeVBAHandler(ABC):
    """Base class for handling VBA operations across different Office applications."""

    def __init__(self, doc_path: str, vba_dir: Optional[str] = None, encoding: str = "cp1252"):
        self.doc_path = doc_path
        self.vba_dir = Path(vba_dir) if vba_dir else Path.cwd()
        self.vba_dir = self.vba_dir.resolve()
        self.encoding = encoding
        self.app = None
        self.doc = None

        # Map component types to file extensions
        self.type_to_ext = {
            1: ".bas",  # Standard Module
            2: ".cls",  # Class Module
            3: ".frm",  # MSForm
            100: ".cls",  # Document Module
        }

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
        if self.app is None:
            self.app = win32com.client.Dispatch(self.app_progid)
            self.app.Visible = True

    def open_document(self) -> None:
        """Open the Office document."""
        if self.doc is None:
            self.initialize_app()
            self.doc = self.app.Documents.Open(str(self.doc_path))

    def save_document(self) -> None:
        """Save the document if it's open."""
        if self.doc is not None:
            try:
                self.doc.Save()
                print("Document has been saved and left open for further editing.")
            except win32com.client.pywintypes.com_error as e:
                print(f"Warning: Failed to save document: {e}")

    def handle_document_module(self, component: Any, content: str, temp_file: Path) -> None:
        """Handle the special document module (ThisDocument, ThisWorkbook, etc.)."""
        # Skip header section for document module
        content_lines = content.splitlines()
        if len(content_lines) > 9:
            actual_code = "\n".join(content_lines[9:])
        else:
            actual_code = ""

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

    def export_vba(self, save_metadata: bool = False) -> None:
        """Export VBA content from the Office document."""
        try:
            self.open_document()
            vba_project = self.get_vba_project()
            components = vba_project.VBComponents

            if not components.Count:
                print("No VBA components found in the document.")
                return

            print(f"\nFound {components.Count} VBA components:")
            component_names = [component.Name for component in components]
            print(", ".join(component_names))

            detected_encodings = {}

            for component in components:
                try:
                    if component.Type not in self.type_to_ext:
                        print(f"Skipping {component.Name} (unsupported type {component.Type})")
                        continue

                    file_name = f"{component.Name}{self.type_to_ext[component.Type]}"
                    temp_file = self.vba_dir / f"{file_name}.temp"
                    final_file = self.vba_dir / file_name

                    # Export to temporary file
                    component.Export(str(temp_file))

                    # Read with specified encoding and write as UTF-8
                    with open(temp_file, "r", encoding=self.encoding) as source:
                        content = source.read()

                    with open(final_file, "w", encoding="utf-8") as target:
                        target.write(content)

                    temp_file.unlink()
                    print(f"Exported: {final_file}")

                except Exception as e:
                    print(f"Warning: Failed to export {component.Name}: {e}")
                    if temp_file.exists():
                        temp_file.unlink()
                    continue

            if save_metadata:
                self._save_metadata(detected_encodings)

            os.startfile(self.vba_dir)

        except Exception as e:
            raise RuntimeError(f"Failed to export VBA content: {e}")
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
                print("No VBA files found to import.")
                return

            print(f"\nFound {len(vba_files)} VBA files to import:")
            print(", ".join(f.name for f in vba_files))

            for vba_file in vba_files:
                try:
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
                        except Exception:
                            pass

                        # Import the component
                        components.Import(str(temp_file))

                    temp_file.unlink()
                    print(f"Imported: {vba_file.name}")

                except Exception as e:
                    print(f"Warning: Failed to import {vba_file.name}: {e}")
                    if temp_file.exists():
                        temp_file.unlink()
                    continue

        except Exception as e:
            raise RuntimeError(f"Failed to import VBA content: {e}")
        finally:
            self.save_document()

    def watch_changes(self) -> None:
        """Watch for changes in VBA files and automatically reimport them."""
        try:
            print(f"Watching for changes in {self.vba_dir}...")

            for changes in watch(
                self.vba_dir,
                watcher_cls=RegExpWatcher,
                watcher_kwargs=dict(re_files=r"^.*(\.cls|\.frm|\.bas)$"),
                normal_sleep=400,
            ):
                for change_type, path in changes:
                    if change_type == Change.modified:
                        try:
                            self.import_vba()
                        except Exception as e:
                            print(f"Error handling changes: {e}")

        except KeyboardInterrupt:
            print("\nStopping VBA editor...")
        finally:
            print("VBA editor stopped. Your changes have been saved.")

    def _save_metadata(self, encodings: Dict[str, Dict[str, Any]]) -> None:
        """Save metadata including encoding information."""
        metadata = {
            "source_document": str(self.doc_path),
            "export_date": datetime.datetime.now().isoformat(),
            "encoding_mode": "fixed",
            "encodings": encodings,
        }

        metadata_path = self.vba_dir / "vba_metadata.json"
        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2)

        print(f"Metadata saved to {metadata_path}")


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
            raise RuntimeError(
                "Cannot access VBA project. Please ensure 'Trust access to the VBA project object model' "
                "is enabled in Word Trust Center Settings."
            ) from e

    def get_document_module_name(self) -> str:
        return "ThisDocument"
