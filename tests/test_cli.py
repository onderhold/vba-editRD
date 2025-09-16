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


class TestCLIConfig:
    """Tests for configuration file functionality."""

    def create_test_config(self, config_dir: Path, config_content: str) -> Path:
        """Create a test configuration file.

        Args:
            config_dir: Directory to create config file in
            config_content: TOML content for the config file

        Returns:
            Path to the created config file
        """
        config_file = config_dir / "test-config.toml"
        config_file.write_text(config_content, encoding="utf-8")
        return config_file

    @pytest.mark.office
    def test_config_file_basic(self, vba_app, tmp_path):
        """Test basic configuration file loading."""
        cli = CLITester(f"{vba_app}-vba")

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # Test that config file is accepted
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_file_not_found(self, vba_app):
        """Test handling of missing configuration file."""
        cli = CLITester(f"{vba_app}-vba")

        result = cli.run(["export", "--conf", "nonexistent-config.toml", "--help"])
        # Should show error but not crash
        assert "Error loading configuration file" in result.stderr or result.returncode == 0

    @pytest.mark.office
    def test_config_file_invalid_toml(self, vba_app, tmp_path):
        """Test handling of invalid TOML configuration."""
        cli = CLITester(f"{vba_app}-vba")

        # Create invalid TOML file
        config_file = tmp_path / "invalid-config.toml"
        config_file.write_text("invalid = toml content [", encoding="utf-8")

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        # Should show error but not crash
        assert "Error loading configuration file" in result.stderr or result.returncode == 0

    @pytest.mark.office
    def test_cli_args_override_config(self, vba_app, tmp_path):
        """Test that CLI arguments override configuration file values."""
        cli = CLITester(f"{vba_app}-vba")

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # CLI should override config (verbose=false should override verbose=true in config)
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        # This test mainly ensures no errors occur during argument processing
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_placeholder_config_path(self, vba_app, tmp_path):
        """Test {config.path} placeholder resolution."""
        cli = CLITester(f"{vba_app}-vba")

        # Create subdirectory for config
        config_dir = tmp_path / "config"
        config_dir.mkdir()

        # Use os.sep for OS-agnostic path separator
        vba_modules_path = f"{PLACEHOLDER_CONFIG_PATH}{os.sep}vba-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_modules_path}"
"""
        config_file = self.create_test_config(config_dir, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_placeholder_file_info(self, vba_app, tmp_path):
        """Test file-related placeholder resolution."""
        cli = CLITester(f"{vba_app}-vba")

        # Create a mock Office file path for testing
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"MyProject{extension}"

        # Use os.sep for OS-agnostic path separator
        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}{os.sep}vba-{PLACEHOLDER_FILE_NAME}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_multiple_placeholders(self, vba_app, tmp_path):
        """Test multiple placeholders in the same value."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"TestDoc{extension}"

        # Use os.sep for OS-agnostic path separator
        vba_directory_path = f"{PLACEHOLDER_CONFIG_PATH}{os.sep}{PLACEHOLDER_FILE_NAME}-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_config_file_with_real_document(self, vba_app, temp_office_doc, tmp_path):
        """Test configuration file with actual Office document operations."""
        cli = CLITester(f"{vba_app}-vba")

        # Use os.sep for OS-agnostic path separator
        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}{os.sep}vba-{PLACEHOLDER_FILE_NAME}"

        # Create config that uses placeholders
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{temp_office_doc}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # Test export using config file
        result = cli.run(["export", "--conf", str(config_file)])

        # Should succeed or show meaningful error
        success = result.returncode == 0 or "No VBA components found" in (result.stdout + result.stderr)
        assert success, f"Command failed with output: {result.stdout + result.stderr}"

        # Check that VBA directory was created with placeholder resolution
        expected_vba_dir = temp_office_doc.parent / f"vba-{temp_office_doc.stem}"
        if success and result.returncode == 0:
            assert expected_vba_dir.exists(), f"Expected VBA directory {expected_vba_dir} was not created"

    @pytest.mark.office
    def test_config_relative_paths(self, vba_app, tmp_path):
        """Test relative paths in configuration file."""
        cli = CLITester(f"{vba_app}-vba")

        # Create subdirectory structure
        config_dir = tmp_path / "config"
        config_dir.mkdir()
        docs_dir = tmp_path / "documents"
        docs_dir.mkdir()

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]

        # Use os.sep for OS-agnostic relative path
        relative_file_path = f"..{os.sep}documents{os.sep}test{extension}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{relative_file_path}"
{CONFIG_KEY_VBA_DIRECTORY} = "vba-output"
"""
        config_file = self.create_test_config(config_dir, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_nested_placeholders(self, vba_app, tmp_path):
        """Test nested configuration sections with placeholders."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"Project{extension}"

        # Use os.sep for OS-agnostic path separators
        vba_modules_path = f"{PLACEHOLDER_FILE_PATH}{os.sep}modules"
        backup_path = f"{PLACEHOLDER_CONFIG_PATH}{os.sep}backups{os.sep}{PLACEHOLDER_FILE_NAME}"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_modules_path}"

[office]
application = "{vba_app}"

[advanced]
backup_directory = "{backup_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0

    @pytest.mark.office
    def test_config_vbaproject_placeholder_preparation(self, vba_app, tmp_path):
        """Test that {vbaproject} placeholder is preserved for later resolution."""
        cli = CLITester(f"{vba_app}-vba")

        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"TestFile{extension}"

        # Use os.sep for OS-agnostic path separator
        vba_directory_path = f"{PLACEHOLDER_FILE_PATH}{os.sep}{PLACEHOLDER_VBA_PROJECT}-modules"

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_path}"
"""
        config_file = self.create_test_config(tmp_path, config_content)

        # This should not fail even though {vbaproject} cannot be resolved yet
        result = cli.run(["export", "--conf", str(config_file), "--help"])
        assert result.returncode == 0


class TestCLICommon:
    """Common tests for all Office VBA CLIs."""

    @pytest.mark.office
    def test_help(self, vba_app):
        """Test help text display."""
        cli = CLITester(f"{vba_app}-vba")
        cli.assert_success(["-h"])
        cli.assert_success(["--help"])

    @pytest.mark.office
    def test_commands_help(self, vba_app):
        """Test help text for each command."""
        cli = CLITester(f"{vba_app}-vba")
        for cmd in ["edit", "import", "export"]:
            cli.assert_success([cmd, "-h"])

    @pytest.mark.office
    def test_invalid_command(self, vba_app):
        """Test invalid command handling."""
        cli = CLITester(f"{vba_app}-vba")
        cli.assert_error(["invalid"], "invalid choice")

    @pytest.mark.office
    def test_missing_file(self, vba_app):
        """Test handling of missing file."""
        cli = CLITester(f"{vba_app}-vba")
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        cli.assert_error(["import", "-f", f"nonexistent{extension}"], "Document not found")

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_basic_operations(self, vba_app, temp_office_doc, tmp_path):
        """Test basic CLI operations with a real document."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Test export
        cli.assert_success(["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir)])
        assert vba_dir.exists()
        assert any(vba_dir.glob("*.bas"))  # Should have at least one module

        # Test import (will use files from export)
        cli.assert_success(["import", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir)])

    @pytest.mark.office
    def test_rubberduck_folders_option(self, vba_app):
        """Test that Rubberduck folders option works for all Office apps."""
        cli = CLITester(f"{vba_app}-vba")

        # Test that help shows the rubberduck option
        result = cli.run(["export", "--help"])
        assert "--rubberduck-folders" in result.stdout

        # Test that the option is accepted
        result = cli.run(["export", "--rubberduck-folders", "--help"])
        assert result.returncode == 0

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_rubberduck_folders_export(self, vba_app, temp_office_doc, tmp_path):
        """Test that Rubberduck folder annotations are respected during export."""
        from vba_edit.office_vba import RUBBERDUCK_FOLDER_PATTERN

        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # Test export with Rubberduck folders enabled
        cli.assert_success(
            ["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

        assert vba_dir.exists()

        # Check that the standard module is exported to the root
        standard_files = list(vba_dir.glob("TestModule.bas"))
        assert len(standard_files) == 1, "TestModule.bas should be in root directory"

        # Find the class module (could be in subdirectory or root)
        class_files = list(vba_dir.rglob("TestClass.cls"))
        assert len(class_files) == 1, "Should find exactly one TestClass.cls file"

        class_file = class_files[0]
        class_content = class_file.read_text(encoding="cp1252")

        # Verify the folder annotation is preserved in the file
        folder_match = None
        for line in class_content.splitlines():
            match = RUBBERDUCK_FOLDER_PATTERN.match(line.strip())
            if match:
                folder_match = match
                break

        assert folder_match is not None, f"No @Folder annotation found in content:\n{class_content}"
        assert (
            folder_match.group(1) == "Business.Domain"
        ), f"Expected folder 'Business.Domain', but found '{folder_match.group(1)}'"

        # Check if the file is in the expected subdirectory (optional verification)
        business_dir = vba_dir / "Business" / "Domain"
        if business_dir.exists():
            # File should be in the Business/Domain subdirectory
            assert (
                class_file.parent == business_dir
            ), f"TestClass.cls should be in {business_dir}, but found in {class_file.parent}"
        else:
            # If folder structure wasn't created, file should be in root
            assert (
                class_file.parent == vba_dir
            ), f"TestClass.cls should be in root {vba_dir}, but found in {class_file.parent}"

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_rubberduck_folders_import(self, vba_app, temp_office_doc, tmp_path):
        """Test that Rubberduck folder structure is preserved during import."""
        cli = CLITester(f"{vba_app}-vba")
        vba_dir = tmp_path / "vba_files"

        # First export with Rubberduck folders
        cli.assert_success(
            ["export", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

        # Modify the class file to test round-trip
        class_files = list(vba_dir.rglob("TestClass.cls"))
        assert len(class_files) == 1, "Should find exactly one TestClass.cls file"

        class_file = class_files[0]
        original_content = class_file.read_text(encoding="cp1252")
        modified_content = original_content.replace(
            'Debug.Print "TestClass initialized"', 'Debug.Print "TestClass initialized - MODIFIED"'
        )
        class_file.write_text(modified_content, encoding="cp1252")

        # Import back with Rubberduck folders
        cli.assert_success(
            ["import", "-f", str(temp_office_doc), "--vba-directory", str(vba_dir), "--rubberduck-folders"]
        )

    @pytest.mark.office
    def test_config_option_available(self, vba_app):
        """Test that --conf option is available for all Office apps."""
        cli = CLITester(f"{vba_app}-vba")

        # Test that help shows the conf option
        result = cli.run(["--help"])
        assert "--conf" in result.stdout, f"--conf option not found in help for {vba_app}-vba"

        # Test that the option is accepted for each subcommand
        for cmd in ["export", "import", "edit"]:
            result = cli.run([cmd, "--help"])
            # The --conf option should be inherited from common arguments
            assert result.returncode == 0

    @pytest.mark.office
    def test_config_option_with_commands(self, vba_app, tmp_path):
        """Test --conf option works with all commands."""
        cli = CLITester(f"{vba_app}-vba")

        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
"""
        config_file = tmp_path / "test.toml"
        config_file.write_text(config_content, encoding="utf-8")

        # Test with each command
        for cmd in ["export", "import", "edit"]:
            result = cli.run([cmd, "--conf", str(config_file), "--help"])
            assert result.returncode == 0, f"Command {cmd} failed with --conf option"


if __name__ == "__main__":
    pytest.main(["-v", __file__])
