# Description: Custom exceptions for the VBA editor.


class OfficeError(Exception):
    """Base exception class for Office-related errors."""

    pass


class PathError(OfficeError):
    """Exception raised for path-related errors."""

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


class VBAError(Exception):
    """Base exception class for all VBA-related errors.

    This exception serves as the parent class for all specific VBA error types
    in the module. It provides a common base for error handling and allows
    catching all VBA-related errors with a single except clause.
    """

    pass


class VBAAccessError(VBAError):
    """Exception raised when access to the VBA project is denied.

    This typically occurs when "Trust access to the VBA project object model"
    is not enabled in the Office application's Trust Center settings.
    """

    pass


class VBAImportError(VBAError):
    """Exception raised when importing VBA components fails.

    This can occur due to various reasons such as invalid file format,
    encoding issues, or problems with the VBA project structure.
    """

    pass


class VBAExportError(VBAError):
    """Exception raised when exporting VBA components fails.

    This can occur due to file system permissions, encoding issues,
    or problems accessing the VBA components.
    """

    pass


class DocumentClosedError(VBAError):
    """Exception raised when attempting to access a closed Office document.

    This exception includes a custom error message that provides guidance
    on how to handle changes made after document closure.

    Args:
        doc_type (str): Type of document (e.g., "workbook", "document")
    """

    def __init__(self, doc_type: str = "document"):
        super().__init__(
            f"\nThe Office {doc_type} has been closed. The edit session will be terminated.\n"
            f"IMPORTANT: Any changes made after closing the {doc_type} must be imported using\n"
            f"'*-vba import' or by saving the file again in the next edit session.\n"
            f"As of version 0.2.1, the '*-vba edit' command will no longer overwrite files\n"
            f"already present in the VBA directory."
        )


class RPCError(VBAError):
    """Exception raised when the RPC server becomes unavailable.

    This typically occurs when the Office application crashes or is forcefully closed.

    Args:
        app_name (str): Name of the Office application
    """

    def __init__(self, app_name: str = "Office application"):
        super().__init__(
            f"\nLost connection to {app_name}. The edit session will be terminated.\n"
            f"IMPORTANT: Any changes made after closing {app_name} must be imported using\n"
            f"'*-vba import' or by saving the file again in the next edit session.\n"
            f"As of version 0.2.1, the '*-vba edit' command will no longer overwrite files\n"
            f"already present in the VBA directory."
        )


def check_rpc_error(error: Exception) -> bool:
    """Check if an exception is related to RPC server unavailability.

    This function examines the error message for common indicators of RPC
    server connection issues.

    Args:
        error: The exception to check

    Returns:
        bool: True if the error appears to be RPC-related, False otherwise
    """
    error_str = str(error).lower()
    rpc_indicators = [
        "rpc server",
        "rpc-server",
        "remote procedure call",
        "0x800706BA",  # RPC server unavailable error code
        "-2147023174",  # Same error in decimal
    ]
    return any(indicator in error_str for indicator in rpc_indicators)
