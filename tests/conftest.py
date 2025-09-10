"""Pytest configuration file."""

import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--apps",
        action="store",
        default="all",
        help="Comma-separated list of apps to test (excel,word,access) or 'all' for all available apps",
    )


def pytest_configure(config):
    """Register custom marks."""
    markers = [
        "excel: mark test as Excel-specific",
        "word: mark test as Word-specific",
        "access: mark test as Access-specific",
        "office: mark test as general Office test",
        "com: marks tests that require COM initialization",
        "integration: mark test as integration test",
    ]
    for marker in markers:
        config.addinivalue_line("markers", marker)


def pytest_collection_modifyitems(config, items):
    """Skip tests based on selected apps."""
    apps_option = config.getoption("--apps")
    if apps_option.lower() == "all":
        return  # Don't skip anything

    selected_apps = [app.strip().lower() for app in apps_option.split(",")]

    # Skip tests that don't match selected apps
    for item in items:
        # Check if test has app-specific markers
        test_apps = []
        if item.get_closest_marker("excel"):
            test_apps.append("excel")
        if item.get_closest_marker("word"):
            test_apps.append("word")
        if item.get_closest_marker("access"):
            test_apps.append("access")

        # If test has app-specific markers and none match selected apps, skip it
        if test_apps and not any(app in selected_apps for app in test_apps):
            item.add_marker(pytest.mark.skip(reason=f"Test requires {test_apps} but only {selected_apps} selected"))


@pytest.fixture
def selected_apps(request):
    """Get the list of apps selected for testing."""
    apps_option = request.config.getoption("--apps")
    if apps_option.lower() == "all":
        return ["excel", "word", "access"]
    else:
        # Parse comma-separated list and validate
        apps = [app.strip().lower() for app in apps_option.split(",")]
        valid_apps = ["excel", "word", "access"]
        invalid_apps = [app for app in apps if app not in valid_apps]
        if invalid_apps:
            raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")
        return apps


@pytest.fixture
def excel_only(request):
    """Check if running in Excel-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["excel"]


@pytest.fixture
def word_only(request):
    """Check if running in Word-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["word"]


@pytest.fixture
def access_only(request):
    """Check if running in Access-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["access"]
