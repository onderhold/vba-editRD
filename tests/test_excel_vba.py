import subprocess
import sys
import pytest


@pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
def test_word_vba_help():
    result = subprocess.run(["excel-vba", "-h"], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba" in result.stdout
    assert "Commands:" in result.stdout
    assert "edit" in result.stdout
    assert "import" in result.stdout
    assert "export" in result.stdout


@pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
def test_word_vba_edit():
    result = subprocess.run(["excel-vba", "edit", "-h"], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba edit" in result.stdout
    assert "--verbose" in result.stdout


@pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
def test_word_vba_import():
    result = subprocess.run(["excel-vba", "import", "-h"], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba import" in result.stdout
    assert "--verbose" in result.stdout


@pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
def test_word_vba_export():
    result = subprocess.run(["excel-vba", "export", "-h"], capture_output=True, text=True)
    assert result.returncode == 0
    assert "usage: excel-vba export" in result.stdout
    assert "--verbose" in result.stdout
