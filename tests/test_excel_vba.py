import pytest
import subprocess
from unittest.mock import patch  # , MagicMock
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

def test_excel_vba_help():
    result = subprocess.run(['excel-vba', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba" in result.stdout
    assert "Commands:" in result.stdout
    assert "edit" in result.stdout
    assert "import" in result.stdout
    assert "export" in result.stdout

def test_excel_vba_edit():
    result = subprocess.run(['excel-vba', 'edit', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba [-h] [-f FILE] [-v] [--version] {edit" in result.stdout
    assert "--verbose" in result.stdout


def test_excel_vba_import():
    result = subprocess.run(['excel-vba', 'import', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba [-h] [-f FILE] [-v] [--version] {edit,import" in result.stdout
    assert "--file" in result.stdout

def test_excel_vba_export():
    result = subprocess.run(['excel-vba', 'export', '-h'], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba [-h] [-f FILE] [-v] [--version] {edit,import,export}" in result.stdout
    assert "--version" in result.stdout

# Test version flag
def test_version_flag():
    test_args = ["excel_vba.py", "--version"]
    with pytest.raises(SystemExit) as exc_info:
        with patch.object(sys, "argv", test_args):
            main()
    assert exc_info.value.code == 0
