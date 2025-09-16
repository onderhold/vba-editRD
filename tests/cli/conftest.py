"""Pytest configuration for CLI tests."""

import pytest
from .helpers import get_installed_apps


@pytest.fixture
def vba_app():
    """VBA application fixture - will be parametrized by pytest_generate_tests."""
    # This fixture body will never execute because pytest_generate_tests
    # will parametrize it with actual values
    pass


def pytest_generate_tests(metafunc):
    """Dynamically parametrize vba_app based on command line options."""
    if "vba_app" in metafunc.fixturenames:
        # Get selected apps from command line
        apps_option = metafunc.config.getoption("--apps")
        if apps_option.lower() == "all":
            selected_apps = ["excel", "word", "access"]
        else:
            selected_apps = [app.strip().lower() for app in apps_option.split(",")]
            valid_apps = ["excel", "word", "access"]
            invalid_apps = [app for app in selected_apps if app not in valid_apps]
            if invalid_apps:
                raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")

        apps = get_installed_apps(selected_apps=selected_apps)
        metafunc.parametrize("vba_app", apps, ids=lambda x: f"{x}-vba")