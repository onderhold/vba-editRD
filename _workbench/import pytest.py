import pytest
import tempfile
import sys
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from argparse import ArgumentParser
from ..excel_vba import create_cli_parser, handle_excel_vba_command

"""Tests for Excel VBA CLI functionality."""




class TestCreateCliParser:
    """Test suite for create_cli_parser function."""

    def test_parser_creation(self):
        """Test that parser is created successfully."""
        parser = create_cli_parser()
        assert isinstance(parser, ArgumentParser)
        assert parser.prog == "excel-vba"

    def test_parser_description_contains_version(self):
        """Test that parser description includes version info."""
        parser = create_cli_parser()
        assert "vba-edit" in parser.description
        assert "excel-vba" in parser.description
        assert "Commands:" in parser.description

    def test_subcommands_exist(self):
        """Test that all required subcommands are present."""
        parser = create_cli_parser()
        
        # Test that command is required
        with pytest.raises(SystemExit):
            parser.parse_args([])
        
        # Test valid subcommands
        valid_commands = ["edit", "import", "export", "check"]
        for cmd in valid_commands:
            args = parser.parse_args([cmd])
            assert args.command == cmd

    def test_edit_command_arguments(self):
        """Test edit command specific arguments."""
        parser = create_cli_parser()
        
        # Test basic edit command
        args = parser.parse_args(["edit"])
        assert args.command == "edit"
        assert hasattr(args, "encoding")
        assert hasattr(args, "detect_encoding")
        assert hasattr(args, "save_headers")
        
        # Test encoding argument
        args = parser.parse_args(["edit", "--encoding", "utf-8"])
        assert args.encoding == "utf-8"
        
        # Test detect-encoding argument
        args = parser.parse_args(["edit", "--detect-encoding"])
        assert args.detect_encoding is True
        
        # Test save-headers argument
        args = parser.parse_args(["edit", "--save-headers"])
        assert args.save_headers is True
        
        # Test mutual exclusivity of encoding options
        with pytest.raises(SystemExit):
            parser.parse_args(["edit", "--encoding", "utf-8", "--detect-encoding"])

    def test_import_command_arguments(self):
        """Test import command specific arguments."""
        parser = create_cli_parser()
        
        # Test basic import command
        args = parser.parse_args(["import"])
        assert args.command == "import"
        assert hasattr(args, "encoding")
        
        # Test encoding argument
        args = parser.parse_args(["import", "--encoding", "cp850"])
        assert args.encoding == "cp850"

    def test_export_command_arguments(self):
        """Test export command specific arguments."""
        parser = create_cli_parser()
        
        # Test basic export command
        args = parser.parse_args(["export"])
        assert args.command == "export"
        assert hasattr(args, "save_metadata")
        assert hasattr(args, "encoding")
        assert hasattr(args, "detect_encoding")
        assert hasattr(args, "save_headers")
        
        # Test save-metadata argument
        args = parser.parse_args(["export", "--save-metadata"])
        assert args.save_metadata is True
        
        # Test save-headers argument
        args = parser.parse_args(["export", "--save-headers"])
        assert args.save_headers is True
        
        # Test mutual exclusivity of encoding options
        with pytest.raises(SystemExit):
            parser.parse_args(["export", "--encoding", "utf-8", "--detect-encoding"])

    def test_check_command_arguments(self):
        """Test check command specific arguments."""
        parser = create_cli_parser()
        
        # Test basic check command
        args = parser.parse_args(["check"])
        assert args.command == "check"
        assert hasattr(args, "verbose")
        assert hasattr(args, "logfile")
        
        # Check command should not have common args like file, vba_directory, xlwings
        assert not hasattr(args, "file")
        assert not hasattr(args, "vba_directory")
        assert not hasattr(args, "xlwings")

    def test_common_arguments(self):
        """Test common arguments across edit, import, export commands."""
        parser = create_cli_parser()
        commands_with_common_args = ["edit", "import", "export"]
        
        for cmd in commands_with_common_args:
            # Test file argument
            args = parser.parse_args([cmd, "--file", "test.xlsm"])
            assert args.file == "test.xlsm"
            
            # Test short form
            args = parser.parse_args([cmd, "-f", "test.xlsm"])
            assert args.file == "test.xlsm"
            
            # Test vba-directory argument
            args = parser.parse_args([cmd, "--vba-directory", "/path/to/vba"])
            assert args.vba_directory == "/path/to/vba"
            
            # Test verbose argument
            args = parser.parse_args([cmd, "--verbose"])
            assert args.verbose is True
            
            # Test short form
            args = parser.parse_args([cmd, "-v"])
            assert args.verbose is True
            
            # Test xlwings argument
            args = parser.parse_args([cmd, "--xlwings"])
            assert args.xlwings is True
            
            # Test short form
            args = parser.parse_args([cmd, "-x"])
            assert args.xlwings is True

    def test_logfile_argument(self):
        """Test logfile argument behavior."""
        parser = create_cli_parser()
        
        # Test without logfile
        args = parser.parse_args(["edit"])
        assert args.logfile is None
        
        # Test with logfile (no path specified)
        args = parser.parse_args(["edit", "--logfile"])
        assert args.logfile == "vba_edit.log"
        
        # Test with custom logfile path
        args = parser.parse_args(["edit", "--logfile", "custom.log"])
        assert args.logfile == "custom.log"
        
        # Test short form
        args = parser.parse_args(["edit", "-l"])
        assert args.logfile == "vba_edit.log"

    def test_default_encoding(self):
        """Test default encoding is set correctly."""
        parser = create_cli_parser()
        
        # Mock get_windows_ansi_codepage to return a known value
        with patch('vba_edit.excel_vba.get_windows_ansi_codepage', return_value="cp1252"):
            parser = create_cli_parser()
            args = parser.parse_args(["edit"])
            assert args.encoding == "cp1252"
        
        # Test fallback when get_windows_ansi_codepage returns None
        with patch('vba_edit.excel_vba.get_windows_ansi_codepage', return_value=None):
            parser = create_cli_parser()
            args = parser.parse_args(["import"])
            assert args.encoding == "cp1252"

    def test_argument_combinations(self):
        """Test various argument combinations."""
        parser = create_cli_parser()
        
        # Test complex edit command
        args = parser.parse_args([
            "edit",
            "--file", "workbook.xlsm",
            "--vba-directory", "/vba/files",
            "--verbose",
            "--logfile", "test.log",
            "--save-headers",
            "--xlwings"
        ])
        assert args.command == "edit"
        assert args.file == "workbook.xlsm"
        assert args.vba_directory == "/vba/files"
        assert args.verbose is True
        assert args.logfile == "test.log"
        assert args.save_headers is True
        assert args.xlwings is True
        
        # Test complex export command
        args = parser.parse_args([
            "export",
            "-f", "test.xlsm",
            "-v",
            "-l",
            "--save-metadata",
            "--detect-encoding"
        ])
        assert args.command == "export"
        assert args.file == "test.xlsm"
        assert args.verbose is True
        assert args.logfile == "vba_edit.log"
        assert args.save_metadata is True
        assert args.detect_encoding is True

    def test_boolean_argument_defaults(self):
        """Test that boolean arguments have correct defaults."""
        parser = create_cli_parser()
        
        args = parser.parse_args(["edit"])
        assert args.verbose is False
        assert args.save_headers is False
        assert args.xlwings is False
        assert args.detect_encoding is False
        
        args = parser.parse_args(["export"])
        assert args.save_metadata is False
        assert args.save_headers is False

    def test_invalid_commands(self):
        """Test that invalid commands raise SystemExit."""
        parser = create_cli_parser()
        
        with pytest.raises(SystemExit):
            parser.parse_args(["invalid_command"])

    def test_help_messages(self):
        """Test that help can be displayed without errors."""
        parser = create_cli_parser()
        
        # Test main help
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["--help"])
        assert exc_info.value.code == 0
        
        # Test subcommand help
        for cmd in ["edit", "import", "export", "check"]:
            with pytest.raises(SystemExit) as exc_info:
                parser.parse_args([cmd, "--help"])
            assert exc_info.value.code == 0

    @patch('vba_edit.excel_vba.get_windows_ansi_codepage')
    def test_encoding_edge_cases(self, mock_encoding):
        """Test encoding argument edge cases."""
        # Test with empty string encoding
        mock_encoding.return_value = ""
        parser = create_cli_parser()
        args = parser.parse_args(["edit"])
        assert args.encoding == "cp1252"  # Should fallback to default
        
        # Test with None encoding
        mock_encoding.return_value = None
        parser = create_cli_parser()
        args = parser.parse_args(["import"])
        assert args.encoding == "cp1252"  # Should fallback to default


class TestHandleExcelVbaCommand:
    """Test suite for handle_excel_vba_command function."""

    @patch('vba_edit.excel_vba.setup_logging')
    @patch('vba_edit.excel_vba.get_active_office_document')
    @patch('vba_edit.excel_vba.get_document_paths')
    @patch('vba_edit.excel_vba.ExcelVBAHandler')
    def test_basic_command_handling(self, mock_handler, mock_get_paths, mock_get_doc, mock_logging):
        """Test basic command handling without xlwings."""
        # Setup mocks
        mock_get_doc.return_value = None
        mock_get_paths.return_value = (Path("test.xlsm"), Path("."))
        mock_handler_instance = Mock()
        mock_handler.return_value = mock_handler_instance
        
        # Create mock args
        args = Mock()
        args.xlwings = False
        args.command = "export"
        args.file = "test.xlsm"
        args.vba_directory = None
        args.verbose = False
        args.logfile = None
        args.detect_encoding = False
        args.encoding = "cp1252"
        args.save_headers = False
        args.save_metadata = True
        
        # Test export command
        handle_excel_vba_command(args)
        
        # Verify handler was created and called
        mock_handler.assert_called_once()
        mock_handler_instance.export_vba.assert_called_once_with(save_metadata=True, overwrite=True)

    @patch('vba_edit.excel_vba.handle_xlwings_command')
    def test_xlwings_option(self, mock_xlwings_handler):
        """Test xlwings option handling."""
        args = Mock()
        args.xlwings = True
        
        with patch('builtins.__import__') as mock_import:
            mock_xlwings = Mock()
            mock_xlwings.__version__ = "0.24.0"
            mock_import.return_value = mock_xlwings
            
            # This should call xlwings handler and return early
            with patch('vba_edit.excel_vba.setup_logging'):
                handle_excel_vba_command(args)
            
            mock_xlwings_handler.assert_called_once_with(args)

    @patch('vba_edit.excel_vba.setup_logging')
    def test_xlwings_import_error(self, mock_logging):
        """Test xlwings import error handling."""
        args = Mock()
        args.xlwings = True
        args.verbose = False
        args.logfile = None
        
        with patch('builtins.__import__', side_effect=ImportError):
            with pytest.raises(SystemExit):
                handle_excel_vba_command(args)


if __name__ == "__main__":
    pytest.main([__file__])