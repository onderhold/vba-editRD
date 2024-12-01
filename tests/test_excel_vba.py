import pytest
from unittest.mock import patch, MagicMock
import sys
from vba_edit.excel_vba import vba_edit, vba_import, vba_export, main

# Test base functions without xlwings
def test_vba_edit_without_xlwings():
    with pytest.raises(NotImplementedError) as exc_info:
        vba_edit("test.xlsm")
    assert "VBA editing without xlwings is not implemented yet" in str(exc_info.value)

def test_vba_import_without_xlwings():
    with pytest.raises(NotImplementedError) as exc_info:
        vba_import("test.xlsm")
    assert "VBA import without xlwings is not implemented yet" in str(exc_info.value)

def test_vba_export_without_xlwings():
    with pytest.raises(NotImplementedError) as exc_info:
        vba_export("test.xlsm")
    assert "VBA export without xlwings is not implemented yet" in str(exc_info.value)

# # Test CLI argument parsing
# @patch('vba_edit.excel_vba.vba_edit')
# def test_main_edit_command_without_xlwings(mock_vba_edit):
#     test_args = ['excel_vba.py', 'edit', '-f', 'test.xlsm']
#     with patch.object(sys, 'argv', test_args):
#         with patch('vba_edit.excel_vba.USE_XLWINGS', False):
#             main()
#             mock_vba_edit.assert_called_once()

# @patch('excel_vba.vba_import')
# def test_main_import_command_without_xlwings(mock_vba_import):
#     test_args = ['excel_vba.py', 'import', '-f', 'test.xlsx']
#     with patch.object(sys, 'argv', test_args):
#         with patch('excel_vba.USE_XLWINGS', False):
#             main()
#             mock_vba_import.assert_called_once()

# @patch('excel_vba.vba_export')
# def test_main_export_command_without_xlwings(mock_vba_export):
#     test_args = ['excel_vba.py', 'export', '-f', 'test.xlsx']
#     with patch.object(sys, 'argv', test_args):
#         with patch('excel_vba.USE_XLWINGS', False):
#             main()
#             mock_vba_export.assert_called_once()

# # Test with xlwings integration
# @pytest.fixture
# def mock_xlwings():
#     with patch.dict('sys.modules', {'xlwings': MagicMock()}):
#         xlwings_mock = MagicMock()
#         xlwings_mock.__version__ = '0.30.0'
#         xlwings_mock.cli.vba_edit = MagicMock()
#         xlwings_mock.cli.vba_import = MagicMock()
#         xlwings_mock.cli.vba_export = MagicMock()
#         yield xlwings_mock

# @patch('excel_vba.USE_XLWINGS', True)
# def test_main_edit_command_with_xlwings(mock_xlwings):
#     test_args = ['excel_vba.py', 'edit', '-f', 'test.xlsx']
#     with patch.object(sys, 'argv', test_args):
#         main()
#         mock_xlwings.cli.vba_edit.assert_called_once()

# @patch('excel_vba.USE_XLWINGS', True)
# def test_main_import_command_with_xlwings(mock_xlwings):
#     test_args = ['excel_vba.py', 'import', '-f', 'test.xlsx']
#     with patch.object(sys, 'argv', test_args):
#         main()
#         mock_xlwings.cli.vba_import.assert_called_once()

# @patch('excel_vba.USE_XLWINGS', True)
# def test_main_export_command_with_xlwings(mock_xlwings):
#     test_args = ['excel_vba.py', 'export', '-f', 'test.xlsx']
#     with patch.object(sys, 'argv', test_args):
#         main()
#         mock_xlwings.cli.vba_export.assert_called_once()

# # Test verbose flag
# @patch('builtins.print')
# def test_verbose_flag(mock_print):
#     test_args = ['excel_vba.py', 'edit', '-f', 'test.xlsx', '--verbose']
#     with patch.object(sys, 'argv', test_args):
#         with patch('excel_vba.USE_XLWINGS', False):
#             main()
#             mock_print.assert_called()

# Test version flag
def test_version_flag():
    test_args = ['excel_vba.py', '--version']
    with pytest.raises(SystemExit) as exc_info:
        with patch.object(sys, 'argv', test_args):
            main()
    assert exc_info.value.code == 0