import ctypes
import logging
import logging.handlers
import sys
from pathlib import Path
from typing import Optional, Tuple, Callable
import win32com.client
import chardet
from functools import wraps
from watchgod import Change, RegExpWatcher, watch  # noqa: F401

from vba_edit.office_vba import DocumentClosedError, RPCError, check_rpc_error

# Configure module logger
logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False, logfile: Optional[str] = None) -> None:
    """Configure root logger.

    Args:
        verbose: Enable verbose (DEBUG) logging if True
        logfile: Path to log file. If None, only console logging is enabled.
    """
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    # Remove any existing handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    root_logger.addHandler(console_handler)

    # File handler with rotation (only if logfile is specified)
    if logfile:
        try:
            file_handler = logging.handlers.RotatingFileHandler(
                logfile,
                maxBytes=1024 * 1024,  # 1MB
                backupCount=3,
                encoding="utf-8",
            )
            file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
            file_handler.setLevel(logging.DEBUG)
            root_logger.addHandler(file_handler)
        except Exception as e:
            logger.warning(f"Could not set up file logging: {e}")


def error_handler(func: Callable) -> Callable:
    """Decorator for consistent error handling across functions.

    Args:
        func: Function to wrap with error handling

    Returns:
        Wrapped function with error handling
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except OfficeError:
            logger.debug(f"Known error in {func.__name__}", exc_info=True)
            raise
        except Exception as e:
            logger.error(f"Unexpected error in {func.__name__}: {str(e)}")
            logger.debug("Detailed error information:", exc_info=True)
            raise OfficeError(f"Operation failed: {str(e)}") from e

    return wrapper


class OfficeError(Exception):
    """Base exception class for Office-related errors."""

    pass


class DocumentNotFoundError(OfficeError):
    """Exception raised when document cannot be found."""

    pass


class ApplicationError(OfficeError):
    """Exception raised when there are issues with Office applications."""

    pass


class EncodingError(OfficeError):
    """Exception raised when there are encoding-related issues."""

    pass


class VBAAccessError(OfficeError):
    """Exception raised when VBA project access is denied."""

    pass


class VBAFileChangeHandler:
    """Handler for VBA file changes."""

    def __init__(self, doc_path: str, vba_dir: str, encoding: Optional[str] = "cp1252"):
        """Initialize the VBA file change handler.

        Args:
            doc_path: Path to the Word document
            vba_dir: Directory containing VBA files
            encoding: Character encoding for VBA files
        """
        self.doc_path = Path(doc_path).resolve()
        self.vba_dir = Path(vba_dir).resolve()
        self.encoding = encoding
        self.word = None
        self.doc = None
        self.logger = logging.getLogger(__name__)

        self.logger.debug(f"Initialized VBAFileChangeHandler with document: {self.doc_path}")
        self.logger.debug(f"VBA directory: {self.vba_dir}")
        self.logger.debug(f"Using encoding: {self.encoding}")

    @error_handler
    def import_changed_file(self, file_path: Path) -> None:
        """Import a single VBA file that has changed.

        Args:
            file_path: Path to the changed VBA file

        Raises:
            VBAAccessError: If VBA project access is denied
            DocumentClosedError: If document is closed
            RPCError: If Word application is not available
            OfficeError: For other VBA-related errors
        """
        self.logger.info(f"Processing changes in {file_path.name}")
        temp_file = None

        try:
            if self.word is None:
                self.logger.debug("Initializing Word application")
                try:
                    self.word = win32com.client.Dispatch("Word.Application")
                    self.word.Visible = True
                    self.doc = self.word.Documents.Open(str(self.doc_path))
                except Exception as e:
                    if check_rpc_error(e):
                        raise RPCError(
                            "\nWord application is not available. Please ensure Word is running and "
                            "try again with 'word-vba import' to import your changes."
                        ) from e
                    raise

            try:
                vba_project = self.doc.VBProject
            except Exception as e:
                if check_rpc_error(e):
                    raise DocumentClosedError(
                        "\nThe Word document has been closed. The edit session will be terminated.\n"
                        "IMPORTANT: Any changes made after closing the document must be imported using\n"
                        "'word-vba import' before starting a new edit session, otherwise they will be lost."
                    ) from e
                raise VBAAccessError(
                    "Cannot access VBA project. Please ensure 'Trust access to the VBA "
                    "project object model' is enabled in Word Trust Center Settings."
                ) from e

            # Read content with UTF-8 encoding (as exported)
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            components = vba_project.VBComponents
            component_name = file_path.stem

            if component_name == "ThisDocument":
                self.logger.debug("Processing ThisDocument module")
                doc_component = components("ThisDocument")

                # Skip header section for ThisDocument
                content_lines = content.splitlines()
                if len(content_lines) > 9:
                    actual_code = "\n".join(content_lines[9:])
                else:
                    actual_code = ""

                # Convert content to specified encoding
                content_bytes = actual_code.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Read back with proper encoding
                with open(temp_file, "r", encoding=self.encoding) as f:
                    new_code = f.read()

                # Update existing ThisDocument module
                self.logger.debug("Updating ThisDocument module")
                doc_component.CodeModule.DeleteLines(1, doc_component.CodeModule.CountOfLines)
                if new_code.strip():
                    doc_component.CodeModule.AddFromString(new_code)

            else:
                self.logger.debug(f"Processing regular component: {component_name}")
                # Handle regular components
                content_bytes = content.encode(self.encoding)
                temp_file = file_path.with_suffix(".temp")

                with open(temp_file, "wb") as f:
                    f.write(content_bytes)

                # Remove existing component if it exists
                try:
                    existing = components(component_name)
                    self.logger.debug(f"Removing existing component: {component_name}")
                    components.Remove(existing)
                except Exception:
                    self.logger.debug(f"No existing component to remove: {component_name}")

                # Import the component
                self.logger.debug(f"Importing component: {component_name}")
                components.Import(str(temp_file))

            self.logger.info(f"Successfully imported: {file_path.name}")
            self.doc.Save()

        except DocumentClosedError as e:
            self.logger.error(str(e))
            # Don't raise here - let the watch loop continue
        except RPCError as e:
            self.logger.error(str(e))
            # Don't raise here - let the watch loop continue
        except Exception as e:
            self.logger.error(f"Failed to process {file_path.name}: {str(e)}")
            self.logger.debug("Detailed error information:", exc_info=True)
            raise
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except Exception as e:
                    self.logger.warning(f"Failed to remove temporary file {temp_file}: {e}")


@error_handler
def get_active_office_document(app_type: str) -> str:
    """Get the path of the currently active Office document.

    Args:
        app_type: The Office application type ('word', 'excel', 'access', 'powerpoint')

    Returns:
        Full path to the active document

    Raises:
        ValueError: If invalid application type is specified
        ApplicationError: If Office application is not running or no document is active
    """
    app_type = app_type.lower()
    app_mapping = {
        "word": ("Word.Application", "Documents", "ActiveDocument"),
        "excel": ("Excel.Application", "Workbooks", "ActiveWorkbook"),
        "access": ("Access.Application", "CurrentProject", "FullName"),
        "powerpoint": ("PowerPoint.Application", "Presentations", "ActivePresentation"),
    }

    if app_type not in app_mapping:
        raise ValueError(f"Invalid application type. Must be one of: {', '.join(app_mapping.keys())}")

    logger.debug(f"Getting active {app_type} document")
    app_class, collection_name, active_doc_property = app_mapping[app_type]

    try:
        app = win32com.client.GetObject(Class=app_class)

        # Special handling for Access
        if app_type == "access":
            active_doc = getattr(app, collection_name)
            if not active_doc:
                raise ApplicationError("No Access database is currently open")
            return getattr(active_doc, active_doc_property)

        # Handle Word, Excel, and PowerPoint
        collection = getattr(app, collection_name)
        if not collection.Count:
            raise ApplicationError(f"No {app_type.capitalize()} document is currently open")

        active_doc = getattr(app, active_doc_property)
        if not active_doc:
            raise ApplicationError(f"Could not get active {app_type.capitalize()} document")

        doc_path = active_doc.FullName
        logger.debug(f"Found active document: {doc_path}")
        return doc_path

    except ApplicationError:
        raise
    except Exception as e:
        raise ApplicationError(f"Could not connect to {app_type.capitalize()} or get active document: {e}")


@error_handler
def get_document_path(file_path: Optional[str] = None, app_type: str = "word") -> str:
    """Get the document path from either the provided file path or active Office document.

    Args:
        file_path: Optional path to the Office document
        app_type: Type of Office application ('word', 'excel', 'access', 'powerpoint')

    Returns:
        Path to the document

    Raises:
        DocumentNotFoundError: If document cannot be found
        ApplicationError: If Office application is not running or no document is active
    """
    logger.debug(f"Getting document path (file_path={file_path}, app_type={app_type})")

    if file_path:
        path = Path(file_path).resolve()
        if not path.exists():
            raise DocumentNotFoundError(f"Document not found: {path}")
        logger.debug(f"Using provided document path: {path}")
        return str(path)

    try:
        doc_path = get_active_office_document(app_type)
        logger.debug(f"Using active document path: {doc_path}")
        return doc_path
    except Exception as e:
        raise DocumentNotFoundError(f"Could not determine document path: {e}")


@error_handler
def detect_vba_encoding(file_path: str) -> Tuple[str, float]:
    """Detect the encoding of a VBA file using chardet.

    Args:
        file_path: Path to the file to analyze

    Returns:
        Tuple containing the detected encoding and confidence score

    Raises:
        EncodingError: If encoding detection fails
    """
    logger.debug(f"Detecting encoding for file: {file_path}")
    try:
        with open(file_path, "rb") as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)

            if not result["encoding"]:
                raise EncodingError(f"Could not detect encoding for file: {file_path}")

            logger.debug(f"Detected encoding: {result['encoding']} (confidence: {result['confidence']})")
            return result["encoding"], result["confidence"]
    except Exception as e:
        raise EncodingError(f"Failed to detect encoding: {e}")


@error_handler
def get_windows_ansi_codepage() -> Optional[str]:
    """Get the Windows ANSI codepage as a Python encoding string.

    Returns:
        Python encoding name (e.g., 'cp1252') or None if not on Windows
        or if codepage couldn't be determined
    """
    logger.debug("Getting Windows ANSI codepage")
    try:
        codepage = ctypes.windll.kernel32.GetACP()
        encoding = f"cp{codepage}"
        logger.debug(f"Windows ANSI codepage: {encoding}")
        return encoding
    except (AttributeError, OSError) as e:
        logger.debug(f"Could not get Windows ANSI codepage: {e}")
        return None
