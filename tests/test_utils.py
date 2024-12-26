# tests/test_utils.py
import pytest
from pathlib import Path
import tempfile
from unittest.mock import Mock, patch

try:
    # Proper import path for com_error
    import pywintypes

    COM_ERROR = pywintypes.com_error
except ImportError:
    # For test environments without win32com
    COM_ERROR = Exception

from vba_edit.utils import (
    detect_vba_encoding,
    get_document_path,
    is_office_app_installed,
    VBAFileChangeHandler,
    DocumentNotFoundError,
    EncodingError,
)


def test_get_document_path():
    """Test document path resolution with different inputs."""
    with tempfile.TemporaryDirectory() as tmpdir:
        # Test with explicit file path
        test_doc = Path(tmpdir) / "test.docm"
        test_doc.touch()
        assert str(test_doc.resolve()) == get_document_path(str(test_doc))

        # Test with nonexistent file
        with pytest.raises(DocumentNotFoundError):
            get_document_path("nonexistent.docm")

def test_is_office_app_installed_validation():
    """Test input validation of is_office_app_installed."""
    # Test invalid app names
    with pytest.raises(ValueError, match="Unsupported application"):
        is_office_app_installed("invalid_app")
        
    with pytest.raises(ValueError):
        is_office_app_installed("")
        
    # Test case insensitivity
    result = is_office_app_installed("EXCEL")
    assert isinstance(result, bool)  # Just verify return type, not actual value

def test_is_office_app_installed_mock():
    """Test Office detection with mocked COM objects."""
    with patch("win32com.client.GetActiveObject") as mock_active, \
         patch("win32com.client.Dispatch") as mock_dispatch:
        
        # Test successful detection scenarios
        mock_active.return_value = Mock(Name="Excel")
        assert is_office_app_installed("excel") is True
        
        # Test fallback to Dispatch when no running instance
        mock_active.side_effect = COM_ERROR
        mock_app = Mock(Name="Excel")
        mock_dispatch.return_value = mock_app
        assert is_office_app_installed("word") is True
        
        # Test detection of non-installed apps
        mock_active.side_effect = COM_ERROR
        mock_dispatch.side_effect = COM_ERROR
        assert is_office_app_installed("powerpoint") is False


def test_vba_file_change_handler():
    """Test VBA file change handling."""
    with tempfile.TemporaryDirectory() as tmpdir:
        handler = VBAFileChangeHandler(doc_path=str(Path(tmpdir) / "test.docm"), vba_dir=tmpdir)

        # Test initialization
        assert handler.encoding == "cp1252"  # Default encoding
        assert handler.vba_dir == Path(tmpdir).resolve()


def test_detect_vba_encoding_edge_cases():
    """Test encoding detection with various content types."""
    with tempfile.TemporaryDirectory() as tmpdir:
        test_file = Path(tmpdir) / "test.bas"

        # Test empty file
        test_file.write_text("")
        with pytest.raises(EncodingError):
            detect_vba_encoding(str(test_file))

        # Test binary content
        test_file.write_bytes(b"\x00\x01\x02\x03")
        encoding, confidence = detect_vba_encoding(str(test_file))
        assert encoding is not None
