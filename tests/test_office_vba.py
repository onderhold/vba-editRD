"""Tests for Office VBA handling."""

import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from vba_edit.office_vba import (
    DocumentClosedError,
    ExcelVBAHandler,
    RPCError,
    VBAAccessError,
    VBAComponentHandler,
    VBAModuleType,
    WordVBAHandler,
)
from vba_edit.utils import (
    EncodingError,
    VBAFileChangeHandler,
    detect_vba_encoding,
    is_office_app_installed,
)


def pytest_configure(config):
    """Add custom markers."""
    config.addinivalue_line("markers", "integration: mark test as integration test")


# Fixtures
@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_vba_files(temp_dir):
    """Create sample VBA files for testing."""
    # Create standard module
    standard_module = temp_dir / "TestModule.bas"
    standard_module.write_text(
        'Attribute VB_Name = "TestModule"\n' "Sub Test()\n" '    Debug.Print "Hello"\n' "End Sub"
    )

    # Create class module
    class_module = temp_dir / "TestClass.cls"
    class_module.write_text(
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "TestClass"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = False\n"
        "Attribute VB_Exposed = False\n"
        "Public Sub TestMethod()\n"
        '    Debug.Print "Class Method"\n'
        "End Sub"
    )

    # Create document module
    doc_module = temp_dir / "ThisDocument.cls"
    doc_module.write_text(
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "ThisDocument"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = True\n"
        "Attribute VB_Exposed = True\n"
        "Private Sub Document_Open()\n"
        '    Debug.Print "Document Opened"\n'
        "End Sub"
    )

    return temp_dir


@pytest.fixture
def mock_word_handler(temp_dir):
    """Create a WordVBAHandler with mocked COM objects."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        mock_app = Mock()
        mock_doc = Mock()
        mock_app.Documents.Open.return_value = mock_doc
        mock_dispatch.return_value = mock_app

        handler = WordVBAHandler(doc_path=str(temp_dir / "test.docm"), vba_dir=str(temp_dir))
        handler.app = mock_app
        handler.doc = mock_doc
        yield handler


@pytest.fixture
def mock_excel_handler(temp_dir):
    """Create an ExcelVBAHandler with mocked COM objects."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        mock_app = Mock()
        mock_wb = Mock()
        mock_app.Workbooks.Open.return_value = mock_wb
        mock_dispatch.return_value = mock_app

        handler = ExcelVBAHandler(doc_path=str(temp_dir / "test.xlsm"), vba_dir=str(temp_dir))
        handler.app = mock_app
        handler.doc = mock_wb
        yield handler


# Test VBA Component Handler
def test_get_module_type():
    """Test correct identification of VBA module types."""
    handler = VBAComponentHandler()

    test_cases = [
        ("TestModule.bas", VBAModuleType.STANDARD),
        ("TestClass.cls", VBAModuleType.CLASS),
        ("TestForm.frm", VBAModuleType.FORM),
        ("ThisDocument.cls", VBAModuleType.DOCUMENT),
        ("ThisWorkbook.cls", VBAModuleType.DOCUMENT),
        ("Hoja1.cls", VBAModuleType.DOCUMENT),  # Spanish Excel sheet
        ("Sheet1.cls", VBAModuleType.DOCUMENT),  # English Excel sheet
        ("Tabelle1.cls", VBAModuleType.DOCUMENT),  # German Excel sheet
    ]

    for filename, expected_type in test_cases:
        result = handler.get_module_type(Path(filename))
        assert result == expected_type, f"Failed for {filename}"


def test_split_vba_content(sample_vba_files):
    """Test splitting VBA content into header and code sections."""
    handler = VBAComponentHandler()

    # Test class module splitting
    class_content = (sample_vba_files / "TestClass.cls").read_text()
    header, code = handler.split_vba_content(class_content)

    assert "VERSION 1.0 CLASS" in header
    assert "Attribute VB_Name" in header
    assert "Public Sub TestMethod()" in code
    assert "Debug.Print" in code

    # Test standard module splitting
    standard_content = (sample_vba_files / "TestModule.bas").read_text()
    header, code = handler.split_vba_content(standard_content)

    assert "Attribute VB_Name" in header
    assert "Sub Test()" in code


def test_encoding_detection(temp_dir):
    """Test VBA file encoding detection."""
    test_file = temp_dir / "test.bas"

    # Test UTF-8 content
    test_file.write_text('\' UTF-8 content\nSub Test()\nMsgBox "Hello 世界"\nEnd Sub', encoding="utf-8")

    encoding, confidence = detect_vba_encoding(str(test_file))
    assert encoding.lower() in ["utf-8", "ascii"]
    assert confidence > 0.7

    # Test invalid file
    invalid_file = temp_dir / "invalid.bas"
    with pytest.raises(EncodingError):
        detect_vba_encoding(str(invalid_file))


# Test Word VBA Handler
def test_word_handler_initialization(mock_word_handler):
    """Test WordVBAHandler initialization."""
    assert mock_word_handler.app_name == "Word"
    assert mock_word_handler.app_progid == "Word.Application"


def test_word_handler_document_access(mock_word_handler):
    """Test document access checking."""
    # Test normal access
    mock_word_handler.doc.Name = Mock()
    mock_word_handler.doc.Name.return_value = "test.docm"
    mock_word_handler.doc.FullName = str(mock_word_handler.doc_path)
    assert mock_word_handler.is_document_open()

    # Test closed document by simulating an RPC error
    mock_word_handler.doc.Name = Mock()
    mock_word_handler.doc.Name.side_effect = Exception("RPC server is unavailable")
    with pytest.raises((DocumentClosedError, RPCError)):
        mock_word_handler.is_document_open()


# Test Excel VBA Handler
def test_excel_handler_initialization(mock_excel_handler):
    """Test ExcelVBAHandler initialization."""
    assert mock_excel_handler.app_name == "Excel"
    assert mock_excel_handler.app_progid == "Excel.Application"


def test_excel_handler_vba_access(mock_excel_handler):
    """Test VBA project access checking."""

    # Create mock VBProject that raises on access
    def raise_access_error():
        raise Exception("Programmatic access to Visual Basic Project is not trusted")

    mock_vbproject = Mock()
    type(mock_vbproject).VBComponents = property(lambda self: raise_access_error())
    type(mock_excel_handler.doc).VBProject = property(lambda self: mock_vbproject)

    with pytest.raises(VBAAccessError):
        mock_excel_handler.get_vba_project()


def test_file_change_handler(temp_dir):
    """Test VBA file change handling."""
    doc_path = str(temp_dir / "test.docm")
    vba_dir = str(temp_dir)

    handler = VBAFileChangeHandler(doc_path, vba_dir)

    # Test initialization
    assert handler.doc_path == Path(doc_path).resolve()
    assert handler.vba_dir == Path(vba_dir).resolve()
    assert handler.encoding == "cp1252"  # Default encoding


@pytest.mark.skipif(not is_office_app_installed("word"), reason="Word is not installed")
def test_word_export_import_cycle(temp_dir):
    """Integration test for Word VBA export/import cycle.

    Note: This test requires Word installation and should be run selectively.
    """
    print("Running Word export/import cycle test ... not implemented yet")

    # doc_path = temp_dir / "test.docm"
    # Implementation would go here for actual Word automation
    # This is marked as integration test and skipped by default


if __name__ == "__main__":
    pytest.main(["-v"])
