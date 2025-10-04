"""Tests for Office VBA handling."""

import tempfile
import pythoncom
from pathlib import Path
from unittest.mock import Mock, patch, PropertyMock
from contextlib import contextmanager

import threading
import time

import pytest

from vba_edit.office_cli import OfficeVBACLI
from vba_edit.office_vba import (
    VBAComponentHandler,
    WordVBAHandler,
    ExcelVBAHandler,
    AccessVBAHandler,
    VBAModuleType,
)
from vba_edit.exceptions import DocumentNotFoundError, DocumentClosedError, RPCError


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
    standard_module.write_text('Attribute VB_Name = "TestModule"\nSub Test()\n    Debug.Print "Hello"\nEnd Sub')

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


class MockCOMError(Exception):
    """Mock COM error for testing without causing Windows fatal exceptions."""

    def __init__(self, hresult, text, details, helpfile=None):
        self.args = (hresult, text, details, helpfile)


@contextmanager
def com_initialized():
    """Context manager for COM initialization/cleanup."""
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


class BaseOfficeMock:
    """Base class for Office mock fixtures."""

    def __init__(self, handler_class, temp_dir, mock_document, file_extension):
        self.handler_class = handler_class
        self.temp_dir = temp_dir
        self.mock_document = mock_document
        self.file_extension = file_extension
        self.handler = None
        self.mock_app = None

    def setup(self):
        """Setup the mock handler and app."""
        doc_path = self.temp_dir / f"test{self.file_extension}"
        doc_path.touch()

        self.mock_app = Mock()
        self._configure_mock_app()

        with patch("win32com.client.Dispatch") as mock_dispatch:
            mock_dispatch.return_value = self.mock_app
            self.handler = self.handler_class(doc_path=str(doc_path), vba_dir=str(self.temp_dir))
            self.handler.app = self.mock_app
            self.handler.doc = self.mock_document

    def cleanup(self):
        """Cleanup mock objects and references."""
        try:
            if hasattr(self, "handler") and self.handler:
                # Force close any real COM connections
                if hasattr(self.handler, "doc") and self.handler.doc:
                    try:
                        # Try to close document if it's real
                        if hasattr(self.handler.doc, "Close"):
                            self.handler.doc.Close()
                    except Exception:
                        pass
                    self.handler.doc = None

                if hasattr(self.handler, "app") and self.handler.app:
                    try:
                        # Try to quit application if it's real
                        if hasattr(self.handler.app, "Quit"):
                            self.handler.app.Quit()
                    except Exception:
                        pass
                    self.handler.app = None

                self.handler = None

            # Clear mock references
            self.mock_app = None

            # Force garbage collection
            import gc

            gc.collect()

        except Exception as e:
            import warnings

            warnings.warn(f"Warning: Error during COM cleanup: {e}")

    def _configure_mock_app(self):
        """Configure app-specific mock behavior. Override in subclasses."""
        raise NotImplementedError


@pytest.fixture
def mock_word_handler(temp_dir, mock_document):
    """Create a WordVBAHandler with mocked COM objects."""

    class WordMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.Documents.Open.return_value = self.mock_document

    with com_initialized():
        mock = WordMock(WordVBAHandler, temp_dir, mock_document, ".docm")
        mock.setup()
        yield mock.handler
        mock.cleanup()


@pytest.fixture
def mock_document():
    """Create a mock document with VBA project."""
    mock_doc = Mock()
    mock_vbproj = Mock()
    mock_doc.VBProject = mock_vbproj
    mock_components = Mock()
    mock_vbproj.VBComponents = mock_components
    return mock_doc


@pytest.fixture
def mock_excel_handler(temp_dir, mock_document):
    """Create an ExcelVBAHandler with mocked COM objects."""

    class ExcelMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.Workbooks.Open.return_value = self.mock_document

    with com_initialized():
        mock = ExcelMock(ExcelVBAHandler, temp_dir, mock_document, ".xlsm")
        mock.setup()
        yield mock.handler
        mock.cleanup()


@pytest.fixture
def mock_access_handler(temp_dir, mock_document):
    """Create an AccessVBAHandler with mocked COM objects."""

    class AccessMock(BaseOfficeMock):
        def _configure_mock_app(self):
            self.mock_app.CurrentDb.return_value = self.mock_document
            # Access-specific configuration
            self.mock_app.VBE = Mock()
            self.mock_app.VBE.ActiveVBProject = self.mock_document.VBProject

    with com_initialized():
        mock = AccessMock(AccessVBAHandler, temp_dir, mock_document, ".accdb")
        mock.setup()
        yield mock.handler
        mock.cleanup()


def create_mock_component():
    """Create a fresh mock component with code module."""
    mock_component = Mock()
    mock_code_module = Mock()
    mock_code_module.CountOfLines = 0
    mock_component.CodeModule = mock_code_module
    return mock_component, mock_code_module


@pytest.mark.office
def test_path_handling(temp_dir):
    """Test path handling in VBA handlers."""
    # Create test document
    doc_path = temp_dir / "test.docm"
    doc_path.touch()
    vba_dir = temp_dir / "vba"

    # Test normal initialization
    handler = WordVBAHandler(doc_path=str(doc_path), vba_dir=str(vba_dir))
    assert handler.doc_path == doc_path.resolve()
    assert handler.vba_dir == vba_dir.resolve()
    assert vba_dir.exists()

    # Test with nonexistent document
    nonexistent = temp_dir / "nonexistent.docm"
    with pytest.raises(DocumentNotFoundError) as exc_info:
        WordVBAHandler(doc_path=str(nonexistent), vba_dir=str(vba_dir))
    assert "not found" in str(exc_info.value).lower()


@pytest.mark.com
@pytest.mark.office
def test_vba_error_handling(mock_word_handler):
    """Test VBA-specific error conditions."""
    # # Create a mock COM error that simulates VBA access denied
    # mock_error = MockCOMError(
    #     -2147352567,  # DISP_E_EXCEPTION
    #     "Exception occurred",
    #     (0, "Microsoft Word", "VBA Project access is not trusted", "wdmain11.chm", 25548, -2146822220),
    #     None,
    # )

    # with patch.object(mock_word_handler.doc, "VBProject", new_callable=PropertyMock) as mock_project:
    #     # Use our mock error instead of pywintypes.com_error
    #     mock_project.side_effect = mock_error
    #     with pytest.raises(VBAAccessError) as exc_info:
    #         mock_word_handler.get_vba_project()
    #     assert "Trust access to the VBA project" in str(exc_info.value)

    # Test RPC server error
    with patch.object(mock_word_handler.doc, "Name", new_callable=PropertyMock) as mock_name:
        mock_name.side_effect = Exception("RPC server is unavailable")
        with pytest.raises(RPCError) as exc_info:
            mock_word_handler.is_document_open()
        assert "lost connection" in str(exc_info.value).lower()

    # # Test general VBA error
    # with patch.object(mock_word_handler.doc, "VBProject", new_callable=PropertyMock) as mock_project:
    #     mock_project.side_effect = Exception("Some unexpected VBA error")
    #     with pytest.raises(VBAError) as exc_info:
    #         mock_word_handler.get_vba_project()
    #     assert "wdmain11.chm" in str(exc_info.value).lower()


@pytest.mark.office
def test_component_handler():
    """Test VBA component handler functionality."""
    handler = VBAComponentHandler()

    # Test module type identification
    assert handler.get_module_type(Path("test.bas")) == VBAModuleType.STANDARD
    assert handler.get_module_type(Path("test.cls")) == VBAModuleType.CLASS
    assert handler.get_module_type(Path("test.frm")) == VBAModuleType.FORM
    assert handler.get_module_type(Path("ThisDocument.cls")) == VBAModuleType.DOCUMENT
    assert handler.get_module_type(Path("ThisWorkbook.cls")) == VBAModuleType.DOCUMENT
    assert handler.get_module_type(Path("Sheet1.cls")) == VBAModuleType.DOCUMENT

    # Test invalid extension
    with pytest.raises(ValueError):
        handler.get_module_type(Path("test.invalid"))


def test_component_header_handling():
    """Test VBA component header handling."""
    handler = VBAComponentHandler()

    # Test header splitting
    content = 'Attribute VB_Name = "TestModule"\nOption Explicit\nSub Test()\nEnd Sub'
    header, code = handler.split_vba_content(content)
    assert 'Attribute VB_Name = "TestModule"' in header
    assert "Option Explicit" in code
    assert "Sub Test()" in code

    # Test minimal header creation
    header = handler.create_minimal_header("TestModule", VBAModuleType.STANDARD)
    assert 'Attribute VB_Name = "TestModule"' in header

    class_header = handler.create_minimal_header("TestClass", VBAModuleType.CLASS)
    assert "VERSION 1.0 CLASS" in class_header
    assert "MultiUse = -1" in class_header


@pytest.mark.com
@pytest.mark.word
@pytest.mark.office
def test_word_handler_functionality(mock_word_handler, sample_vba_files):
    """Test Word VBA handler specific functionality."""
    handler = mock_word_handler

    # Test basic properties
    assert handler.app_name == "Word"
    assert handler.app_progid == "Word.Application"
    assert handler.get_document_module_name() == "ThisDocument"

    # Test document status checking
    type(handler.doc).Name = PropertyMock(return_value="test.docm")
    type(handler.doc).FullName = PropertyMock(return_value=str(handler.doc_path))
    assert handler.is_document_open()

    # Test document module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("ThisDocument", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


@pytest.mark.com
@pytest.mark.excel
@pytest.mark.office
def test_excel_handler_functionality(mock_excel_handler, sample_vba_files):
    """Test Excel VBA handler specific functionality."""
    handler = mock_excel_handler

    # Test basic properties
    assert handler.app_name == "Excel"
    assert handler.app_progid == "Excel.Application"
    assert handler.get_document_module_name() == "ThisWorkbook"

    # Test document status checking
    type(handler.doc).Name = PropertyMock(return_value="test.xlsm")
    type(handler.doc).FullName = PropertyMock(return_value=str(handler.doc_path))
    assert handler.is_document_open()

    # Test workbook module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("ThisWorkbook", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


@pytest.mark.com
@pytest.mark.access
@pytest.mark.office
def test_access_handler_functionality(mock_access_handler, sample_vba_files):
    """Test Access VBA handler specific functionality."""
    handler = mock_access_handler

    # Test basic properties
    assert handler.app_name == "Access"
    assert handler.app_progid == "Access.Application"
    assert handler.get_document_module_name() == ""

    # Test database status checking
    handler.doc.Name = str(handler.doc_path)
    assert handler.is_document_open()

    # Test module update using local mocks
    mock_component, mock_code_module = create_mock_component()
    components = Mock()
    components.return_value = mock_component

    handler._update_document_module("TestModule", "' Test Code", components)
    mock_code_module.AddFromString.assert_called_once_with("' Test Code")


@patch("vba_edit.office_vba.get_document_paths")
@pytest.mark.office
def test_rubberduck_folders_passed_to_handler(mock_get_paths, vba_app, office_app_config, temp_dir):
    """Test that rubberduck_folders option is passed to the handler."""

    config = office_app_config[vba_app]

    # Create the appropriate file for this Office app
    doc_path = temp_dir / f"test{config['extension']}"
    doc_path.touch()

    # Mock the path resolution
    mock_get_paths.return_value = (doc_path, temp_dir)

    # Create args with rubberduck_folders enabled
    cli = OfficeVBACLI(vba_app)
    parser = cli.create_cli_parser()
    args = parser.parse_args(["export", "--rubberduck-folders", "--file", str(doc_path)])

    # Mock the handler class directly on the CLI instance
    with patch.object(cli, "handler_class") as mock_handler_class:
        # Create mock handler instance
        mock_handler_instance = Mock()
        mock_handler_class.return_value = mock_handler_instance

        # Handle the command
        cli.handle_office_vba_command(args)

        # Verify handler was called with use_rubberduck_folders=True
        mock_handler_class.assert_called_once()
        call_kwargs = mock_handler_class.call_args[1]
        assert call_kwargs["use_rubberduck_folders"] is True


@pytest.mark.com
@pytest.mark.office
def test_watch_changes_handling(vba_app, request):
    """Test that watch_changes method exists and can be interrupted."""
    # Get the appropriate handler fixture based on the app
    handler_fixture_name = f"mock_{vba_app}_handler"
    handler = request.getfixturevalue(handler_fixture_name)

    # Test that the method exists and is callable
    assert hasattr(handler, "watch_changes")
    assert callable(handler.watch_changes)

    # Test that the handler has the expected attributes for watching
    assert hasattr(handler, "vba_dir")
    assert handler.vba_dir is not None

    # Rather than actually calling watch_changes (which can hang),
    # just verify the method signature and basic setup
    import inspect

    sig = inspect.signature(handler.watch_changes)
    assert len(sig.parameters) == 0  # Should take no parameters besides self

    # Test passes if we can inspect the method without issues


@pytest.mark.com
@pytest.mark.office
def test_watch_changes_with_threading_v1(vba_app, temp_dir, request):
    """Test file watching functionality using threading."""
    # Get the appropriate handler fixture based on the app
    handler_fixture_name = f"mock_{vba_app}_handler"
    handler = request.getfixturevalue(handler_fixture_name)

    # Create a test VBA file
    test_module = temp_dir / "TestModule.bas"
    test_module.write_text('Attribute VB_Name = "TestModule"\nSub Test()\nEnd Sub')

    # Track when changes are detected
    changes_detected = threading.Event()
    watch_error = None

    # Mock the import_single_file method to signal when a change is processed
    original_import = handler.import_single_file

    def mock_import_with_signal(file_path):
        changes_detected.set()  # Signal that change was detected
        return original_import(file_path)

    handler.import_single_file = Mock(side_effect=mock_import_with_signal)

    def watcher_thread():
        """Run the watcher in a separate thread."""
        nonlocal watch_error
        try:
            handler.watch_changes()
        except Exception as e:
            watch_error = e

    def file_modifier_thread():
        """Modify the file after a short delay."""
        time.sleep(0.5)  # Give watcher time to start

        # Modify the test file
        modified_content = test_module.read_text() + "\n' Added comment"
        test_module.write_text(modified_content)

        # Wait for change detection or timeout
        if changes_detected.wait(timeout=5):
            # Success! Changes were detected
            # Now signal the watcher to stop by mocking document closure
            handler.is_document_open = Mock(return_value=False)
        else:
            # Timeout - force stop the watcher
            handler.is_document_open = Mock(return_value=False)

    # Start both threads
    watcher = threading.Thread(target=watcher_thread, daemon=True)
    modifier = threading.Thread(target=file_modifier_thread, daemon=True)

    watcher.start()
    modifier.start()

    # Wait for both threads to complete (with timeout)
    modifier.join(timeout=10)
    watcher.join(timeout=2)  # Shorter timeout for watcher since it should exit quickly

    # Verify the test results
    if watch_error and not isinstance(watch_error, (DocumentClosedError, KeyboardInterrupt)):
        raise watch_error

    # Verify that changes were actually detected
    assert changes_detected.is_set(), "File changes should have been detected"

    # Verify that import_single_file was called with the correct file
    handler.import_single_file.assert_called_with(test_module)


@pytest.mark.com
@pytest.mark.office
def test_watch_changes_with_threading(vba_app, tmp_path, request):
    """Test file watching functionality using threading."""
    # Get the appropriate handler fixture based on the app
    handler_fixture_name = f"mock_{vba_app}_handler"
    handler = request.getfixturevalue(handler_fixture_name)

    # Create a test VBA file
    vba_dir = tmp_path
    test_module = tmp_path / "TestModule.bas"
    original_content = 'Attribute VB_Name = "TestModule"\nSub Test()\nEnd Sub'
    test_module.write_text(original_content)

    # Update the handler to use the pytest temporary directory
    handler.vba_dir = vba_dir

    # Log file path information for debugging
    import logging

    logger = logging.getLogger("test_office_vba")
    logger.debug("=== File Path Information ===")
    logger.debug(f"Test file name: {test_module.name}")
    logger.debug(f"Test file path: {test_module}")
    logger.debug(f"Absolute path: {test_module.absolute()}")
    logger.debug(f"Parent directory: {test_module.parent}")
    logger.debug(f"File exists: {test_module.exists()}")
    logger.debug(f"VBA directory: {tmp_path}")
    logger.debug(f"Handler VBA dir: {handler.vba_dir}")
    logger.debug("================================")

    # Track when changes are detected and file content verification
    changes_detected = threading.Event()
    content_verified = threading.Event()
    watch_error = None
    file_content_matches = False

    # Mock the import_single_file method to signal when a change is processed
    original_import = handler.import_single_file

    def mock_import_with_signal(file_path):
        changes_detected.set()  # Signal that change was detected
        return original_import(file_path)

    handler.import_single_file = Mock(side_effect=mock_import_with_signal)

    def watcher_thread():
        """Run the watcher in a separate thread."""
        nonlocal watch_error, file_content_matches
        try:
            handler.watch_changes()
        except Exception as e:
            watch_error = e

        # When watcher exits, verify the file content
        current_content = test_module.read_text()
        expected_addition = "' Added comment"
        file_content_matches = expected_addition in current_content

        logger.debug("File content verification:")
        logger.debug(f"Original content: {repr(original_content)}")
        logger.debug(f"Current content: {repr(current_content)}")
        logger.debug(f"Looking for: {repr(expected_addition)}")
        logger.debug(f"Content matches: {file_content_matches}")

        content_verified.set()  # Signal that content verification is done

    def file_modifier_thread():
        """Modify the file after a short delay."""
        time.sleep(0.5)  # Give watcher time to start

        # Modify the test file
        modified_content = test_module.read_text() + "\n' Added comment"
        test_module.write_text(modified_content)
        logger.debug(f"File modified. New content: {repr(modified_content)}")

        # Wait for change detection with shorter timeout than watcher
        if changes_detected.wait(timeout=3):
            logger.debug("Changes detected! Waiting a bit more for processing...")
            time.sleep(0.5)  # Give watcher time to process the change

            # Signal the watcher to stop by mocking document closure
            handler.is_document_open = Mock(return_value=False)
            logger.debug("Signaled watcher to stop")
        else:
            logger.debug("Timeout waiting for change detection - forcing stop")
            # Timeout - force stop the watcher
            handler.is_document_open = Mock(return_value=False)

    # Start both threads
    watcher = threading.Thread(target=watcher_thread, daemon=True)
    modifier = threading.Thread(target=file_modifier_thread, daemon=True)

    watcher.start()
    modifier.start()

    # Wait for modifier to complete first (shorter timeout)
    modifier.join(timeout=5)

    # Wait for watcher to complete content verification (longer timeout)
    watcher.join(timeout=8)

    # Wait for content verification to complete
    content_verified.wait(timeout=2)

    # Verify the test results
    if watch_error and not isinstance(watch_error, (DocumentClosedError, KeyboardInterrupt)):
        logger.error(f"Unexpected watch error: {watch_error}")
        raise watch_error

    # Verify that changes were actually detected
    assert changes_detected.is_set(), "File changes should have been detected within timeout"

    # Verify that import_single_file was called with the correct file
    handler.import_single_file.assert_called_with(test_module)

    # Verify that the file content was actually modified
    assert file_content_matches, "File should contain the added comment"

    # Additional verification: read the file directly to double-check
    final_content = test_module.read_text()
    assert "' Added comment" in final_content, f"File content verification failed. Content: {repr(final_content)}"

    logger.info("✓ All verifications passed!")


@pytest.mark.office
def test_watchfiles_integration():
    """Test that watchfiles is properly integrated and can be imported."""
    watchfiles = pytest.importorskip("watchfiles", reason="watchfiles not available")

    # Verify the Change enum has expected values
    assert hasattr(watchfiles.Change, "added")
    assert hasattr(watchfiles.Change, "modified")
    assert hasattr(watchfiles.Change, "deleted")


if __name__ == "__main__":
    pytest.main(["-v", __file__])
