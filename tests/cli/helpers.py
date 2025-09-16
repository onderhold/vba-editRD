"""Test helpers for CLI interface testing."""

import subprocess
from typing import List, Optional
import pytest
import win32com.client
import pythoncom
from pathlib import Path
import os

from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS, SUPPORTED_APPS
from vba_edit.exceptions import VBAError
from vba_edit.cli_common import (
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_VBA_PROJECT,
    CONFIG_SECTION_GENERAL,
    CONFIG_KEY_FILE,
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_KEY_VERBOSE,
    CONFIG_KEY_RUBBERDUCK_FOLDERS,
)


class ReferenceDocuments:
    """Context manager for handling Office reference documents for testing purposes."""

    def __init__(self, path: Path, app_type: str):
        self.path = path
        self.app_type = app_type.lower()
        self.app = None
        self.doc = None

    def __enter__(self):
        """Open the document and create a basic VBA project."""
        app_configs = {
            "word": {
                "app_name": "Word.Application",
                "doc_method": lambda app: app.Documents.Add(),
                "save_format": 13,  # wdFormatDocumentMacroEnabled
            },
            "excel": {
                "app_name": "Excel.Application",
                "doc_method": lambda app: app.Workbooks.Add(),
                "save_format": 52,  # xlOpenXMLWorkbookMacroEnabled
            },
        }

        try:
            if self.app_type not in app_configs:
                raise ValueError(f"Unsupported application type: {self.app_type}")

            config = app_configs[self.app_type]
            # Initialize COM here
            pythoncom.CoInitialize()

            self.app = win32com.client.Dispatch(config["app_name"])
            self.app.DisplayAlerts = False
            self.app.Visible = False
            self.doc = config["doc_method"](self.app)

            try:
                vba_project = self.doc.VBProject

                # Add standard module with simple test code
                module = vba_project.VBComponents.Add(1)  # 1 = standard module
                module.Name = "TestModule"
                code = 'Sub Test()\n    Debug.Print "Test"\nEnd Sub'
                module.CodeModule.AddFromString(code)

                # Add class module with Rubberduck folder annotation
                class_module = vba_project.VBComponents.Add(2)  # 2 = class module
                class_module.Name = "TestClass"
                class_code = (
                    '\'@Folder("Business.Domain")\n'
                    "Option Explicit\n\n"
                    "Private m_name As String\n\n"
                    "Public Property Get Name() As String\n"
                    "    Name = m_name\n"
                    "End Property\n\n"
                    "Public Property Let Name(ByVal value As String)\n"
                    "    m_name = value\n"
                    "End Property\n\n"
                    "Public Sub Initialize()\n"
                    '    Debug.Print "TestClass initialized"\n'
                    "End Sub"
                )
                class_module.CodeModule.AddFromString(class_code)

            except Exception as ve:
                raise VBAError(
                    "Cannot access VBA project. Please ensure 'Trust access to the "
                    "VBA project object model' is enabled in Trust Center Settings."
                ) from ve

            self.doc.SaveAs(str(self.path), config["save_format"])
            return self.path

        except Exception as e:
            self._cleanup()
            raise VBAError(f"Failed to create test document: {e}") from e

    def _cleanup(self):
        """Clean up resources."""
        if hasattr(self, "doc") and self.doc:
            try:
                self.doc.Close(False)
            except Exception:
                pass
            self.doc = None

        if hasattr(self, "app") and self.app:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up by closing document and quitting application."""
        self._cleanup()


@pytest.fixture
def temp_office_doc(tmp_path, vba_app):
    """Fixture providing a temporary Office document for testing."""
    extension = OFFICE_MACRO_EXTENSIONS[vba_app]
    doc_path = tmp_path / f"test_doc{extension}"

    with ReferenceDocuments(doc_path, vba_app) as path:
        yield path


class CLITester:
    """Helper class for testing CLI interfaces."""

    def __init__(self, command: str):
        """Initialize with base command.

        Args:
            command: CLI command (e.g., 'word-vba', 'excel-vba')
        """
        self.command = command
        self.app_name = command.replace("-vba", "")

    def run(self, args: List[str], input_text: Optional[str] = None) -> subprocess.CompletedProcess:
        """Run CLI command with given arguments.

        Args:
            args: List of command arguments
            input_text: Optional input to provide to command

        Returns:
            CompletedProcess instance with command results
        """
        cmd = [self.command] + args
        return subprocess.run(cmd, input=input_text.encode() if input_text else None, capture_output=True, text=True)

    def assert_success(self, args: List[str], expected_output: Optional[str] = None) -> None:
        """Assert command succeeds and optionally check output.

        Args:
            args: Command arguments
            expected_output: Optional string to check in output
        """
        result = self.run(args)
        full_output = result.stdout + result.stderr
        # Consider success if either return code is 0 or it's an empty VBA project
        success = result.returncode == 0 or "No VBA components found" in full_output
        assert success, f"Command failed with output: {full_output}"
        if expected_output:
            assert expected_output in full_output, f"Expected '{expected_output}' in output"

    def assert_error(self, args: List[str], expected_error: str) -> None:
        """Assert command fails with expected error message.

        Args:
            args: Command arguments
            expected_error: Error message to check for
        """
        result = self.run(args)
        assert result.returncode != 0, "Command should have failed"
        full_output = result.stdout + result.stderr
        assert expected_error in full_output, f"Expected error '{expected_error}' not found in output"


def get_installed_apps(selected_apps=None) -> List[str]:
    """Get list of supported apps that are installed."""
    if selected_apps is None:
        selected_apps = ["excel", "word", "access"]

    return [app for app in selected_apps if app in SUPPORTED_APPS and _check_app_available(app)]


def _check_app_available(app_name: str) -> bool:
    """Check if an Office app is available without using COM.

    Args:
        app_name: Name of Office application to check

    Returns:
        True if app is available, False otherwise
    """
    try:
        cmd = [f"{app_name}-vba", "--help"]
        result = subprocess.run(cmd, capture_output=True, text=True)
        return result.returncode == 0
    except Exception:
        return False


def pytest_generate_tests(metafunc):
    """Dynamically parametrize vba_app based on command line options."""
    if "vba_app" in metafunc.fixturenames:
        # Get selected apps from command line
        apps_option = metafunc.config.getoption("--apps")
        if apps_option.lower() == "all":
            selected_apps = ["excel", "word", "access"]
        else:
            selected_apps = [app.strip().lower() for app in apps_option.split(",")]
            valid_apps = ["excel", "word", "access"]
            invalid_apps = [app for app in selected_apps if app not in valid_apps]
            if invalid_apps:
                raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")

        apps = get_installed_apps(selected_apps=selected_apps)
        metafunc.parametrize("vba_app", apps, ids=lambda x: f"{x}-vba")