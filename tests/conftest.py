"""Pytest configuration file."""


def pytest_configure(config):
    """Register custom marks."""
    markers = ["com: marks tests that require COM initialization", "integration: mark test as integration test"]
    for marker in markers:
        config.addinivalue_line("markers", marker)
