"""Basic CLI functionality tests."""

import pytest
from .helpers import CLITester
from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS
from vba_edit.cli_common import (
    CONFIG_SECTION_GENERAL,
    CONFIG_KEY_VERBOSE,
    get_office_config,  # Import the centralized config function
)


class TestCLIBasic:
    """Basic tests for all Office VBA CLIs."""

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

        # Use centralized config to get the correct document type
        config = get_office_config(vba_app)
        expected_error = f"{config['document_type'].title()} not found"

        cli.assert_error(["import", "-f", f"nonexistent{extension}"], expected_error)

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
