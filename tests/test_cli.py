"""Test helpers for CLI interface testing."""

import subprocess
from typing import List, Optional
import pytest
import win32com.client
import pythoncom
from pathlib import Path

from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS, SUPPORTED_APPS


class ReferenceDocuments:
    """Context manager for handling Office reference documents for testing purposes."""

    def __init__(self, path: Path, app_type: str):
        self.path = path
        self.app_type = app_type
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
                vba_project.VBComponents.Add(1)  # Add standard module
            except Exception as ve:
                raise Exception(
                    "Cannot access VBA project. Please ensure 'Trust access to the "
                    "VBA project object model' is enabled in Trust Center Settings."
                ) from ve

            self.doc.SaveAs(str(self.path), config["save_format"])
            return self.path

        except Exception as e:
            self._cleanup()
            raise Exception(f"Failed to create test document: {e}")

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


def get_installed_apps():
    """Get list of supported apps that are installed."""
    try:
        return [app for app in SUPPORTED_APPS if _check_app_available(app)]
    except Exception:
        return []


def _check_app_available(app_name: str) -> bool:
    """Check if an Office app is available without using COM."""
    try:
        # Run CLI help command as a quick check
        cmd = [f"{app_name}-vba", "--help"]
        result = subprocess.run(cmd, capture_output=True, text=True)
        return result.returncode == 0
    except Exception:
        return False


class CLITester:
    """Helper class for testing CLI interfaces."""

    def __init__(self, command: str):
        """Initialize with base command."""
        self.command = command
        self.app_name = command.replace("-vba", "")

    def run(self, args: List[str], input_text: Optional[str] = None) -> subprocess.CompletedProcess:
        """Run CLI command with given arguments."""
        cmd = [self.command] + args
        return subprocess.run(cmd, input=input_text.encode() if input_text else None, capture_output=True, text=True)

    def assert_success(self, args: List[str], expected_output: Optional[str] = None) -> None:
        """Assert command succeeds and optionally check output."""
        result = self.run(args)
        full_output = result.stdout + result.stderr
        # Consider success if either return code is 0 or it's an empty VBA project
        success = result.returncode == 0 or "No VBA components found" in full_output
        assert success, f"Command failed with output: {full_output}"
        if expected_output:
            assert expected_output in full_output

    def assert_error(self, args: List[str], expected_error: str) -> None:
        """Assert command fails with expected error message."""
        result = self.run(args)
        assert result.returncode != 0
        # Check both stdout and stderr for the error message
        full_output = result.stdout + result.stderr
        assert expected_error in full_output


@pytest.fixture(params=get_installed_apps(), ids=lambda x: f"{x}-vba")
def vba_app(request):
    """Fixture providing installed Office app name."""
    return request.param


class TestCLICommon:
    """Common tests for all Office VBA CLIs."""

    def test_help(self, vba_app):
        """Test help text display."""
        cli = CLITester(f"{vba_app}-vba")
        cli.assert_success(["-h"])
        cli.assert_success(["--help"])

    def test_commands_help(self, vba_app):
        """Test help text for each command."""
        cli = CLITester(f"{vba_app}-vba")
        for cmd in ["edit", "import", "export"]:
            cli.assert_success([cmd, "-h"])

    def test_invalid_command(self, vba_app):
        """Test invalid command handling."""
        cli = CLITester(f"{vba_app}-vba")
        cli.assert_error(["invalid"], "invalid choice")

    def test_missing_file(self, vba_app):
        """Test handling of missing file."""
        cli = CLITester(f"{vba_app}-vba")
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        cli.assert_error(
            ["import", "-f", f"nonexistent{extension}"],
            "not found",  # Generic error message that works for all Office apps
        )
    @pytest.mark.integration
    @pytest.mark.com  # Mark tests that require COM initialization
    def test_directory_creation(self, vba_app, tmp_path):
        """Test VBA directory creation with test document."""
        cli = CLITester(f"{vba_app}-vba")
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"test_doc{extension}"

        with ReferenceDocuments(test_file, vba_app) as _:
            vba_dir = tmp_path / "vba_files"
            cli.assert_success(["export", "--vba-directory", str(vba_dir), "-f", str(test_file)])
            assert vba_dir.exists()
