"""Pytest configuration file."""

import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--apps",
        action="store",
        default="all",
        help="Comma-separated list of apps to test (excel,word,access,powerpoint) or 'all' for all available apps",
    )
    parser.addoption(
        "--check-vba-trust",
        action="store_true",
        default=False,
        help="Check VBA trust access for selected Office applications at startup",
    )


@pytest.fixture
def office_app_config():
    """Get Office application configuration."""
    return {
        "excel": {"handler": "ExcelVBAHandler", "extension": ".xlsm"},
        "word": {"handler": "WordVBAHandler", "extension": ".docm"},
        "access": {"handler": "AccessVBAHandler", "extension": ".accdb"},
        "powerpoint": {"handler": "PowerPointVBAHandler", "extension": ".pptm"},
    }


def pytest_configure(config):
    """Register custom marks and check VBA trust access."""
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

    # Only check VBA trust access when explicitly requested
    if config.getoption("--check-vba-trust"):
        _check_vba_trust_access(config)


def _check_vba_trust_access(config):
    """Check VBA trust access for selected Office applications."""
    try:
        from vba_edit.utils import check_vba_trust_access
    except ImportError:
        print("⚠️  Cannot import VBA trust check utility. Skipping VBA verification.")
        return

    # Get selected apps using the same logic as pytest_generate_tests
    apps_option = config.getoption("--apps")
    if apps_option.lower() == "all":
        selected_apps = ["excel", "word", "access", "powerpoint"]
    else:
        selected_apps = [app.strip().lower() for app in apps_option.split(",")]

    print(f"\n{'=' * 60}")
    print("🔍 CHECKING VBA TRUST ACCESS")
    print(f"{'=' * 60}")
    print(f"Selected Office applications: {', '.join(selected_apps)}")

    failed_apps = {}
    success_apps = []

    for app in selected_apps:
        print(f"Checking {app.title()}...", end=" ", flush=True)

        # Capture both exceptions and output to detect failures
        import io
        from contextlib import redirect_stdout, redirect_stderr

        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()

        try:
            with redirect_stdout(stdout_capture), redirect_stderr(stderr_capture):
                check_vba_trust_access(app)

            # Check if the captured output contains warning messages
            captured_output = stdout_capture.getvalue() + stderr_capture.getvalue()

            # Look for common failure indicators in the output
            failure_indicators = [
                "access seems to be disabled",
                "Trust Access",
                "not enabled",
                "Trust Center",
                "VBA project object model",
            ]

            has_warning = any(indicator in captured_output for indicator in failure_indicators)

            if has_warning:
                print("❌ FAILED")
                # Extract the meaningful part of the warning message
                warning_lines = [line.strip() for line in captured_output.split("\n") if line.strip()]
                warning_msg = warning_lines[-1] if warning_lines else "VBA trust access issue detected"
                failed_apps[app] = warning_msg
            else:
                print("✅ OK")
                success_apps.append(app)

        except Exception as e:
            print("❌ FAILED")
            failed_apps[app] = str(e)

    print(f"{'=' * 60}")

    if failed_apps:
        print("🚨 VBA TRUST ACCESS ISSUES DETECTED")
        print(f"{'=' * 60}")
        print(f"✅ Working: {', '.join(success_apps) if success_apps else 'None'}")
        print(f"❌ Issues:  {', '.join(failed_apps.keys())}")

        print("\n📋 DETAILED ERRORS:")
        for app, error in failed_apps.items():
            print(f"   • {app.title()}: {error}")

        print("\n🔧 TO FIX THESE ISSUES:")
        print(f"   1. Open each affected Office application ({', '.join(failed_apps.keys())})")
        print("   2. Go to: File → Options → Trust Center → Trust Center Settings")
        print("   3. Select: Macro Settings")
        print("   4. Enable: 'Trust access to the VBA project object model'")
        print("   5. Click OK and restart the application")

        print("\n🔧 OR run these commands to check/configure:")
        for app in failed_apps.keys():
            print(f"   {app}-vba check")

        print("\n⚙️  TESTING OPTIONS:")
        print("   • Skip VBA tests:     pytest -m 'not com'")
        print("   • Skip this check:    pytest --skip-vba-check")

        # Show specific command to skip only the problematic apps
        working_apps = [app for app in selected_apps if app not in failed_apps]
        if working_apps:
            print(f"   • Test only working apps: pytest --apps {','.join(working_apps)}")

        print(f"{'=' * 60}")
        print("❌ VBA TRUST ACCESS REQUIRED")
        print("   Configure VBA trust settings above, or use testing options to skip.")

        # Store the failed apps in pytest config for later use
        config._vba_trust_failed_apps = failed_apps

        # Mark all tests to be skipped if VBA trust access is not available
        config._vba_trust_skip_all = True

    else:
        print("✅ ALL OFFICE APPLICATIONS CONFIGURED CORRECTLY!")
        print(f"   VBA trust access enabled for: {', '.join(success_apps)}")
        print(f"{'=' * 60}")

        # Store success info in config
        config._vba_trust_success_apps = success_apps


def pytest_generate_tests(metafunc):
    """Dynamically parametrize vba_app based on command line options."""
    if "vba_app" in metafunc.fixturenames:
        # Import here to avoid circular import issues
        from tests.cli.helpers import get_installed_apps

        # Get selected apps from command line
        apps_option = metafunc.config.getoption("--apps")
        if apps_option.lower() == "all":
            selected_apps = ["excel", "word", "access", "powerpoint"]
        else:
            selected_apps = [app.strip().lower() for app in apps_option.split(",")]
            valid_apps = ["excel", "word", "access", "powerpoint"]
            invalid_apps = [app for app in selected_apps if app not in valid_apps]
            if invalid_apps:
                raise ValueError(f"Invalid apps: {invalid_apps}. Valid options: {valid_apps}")

        apps = get_installed_apps(selected_apps=selected_apps)
        metafunc.parametrize("vba_app", apps, ids=lambda x: f"{x}-vba")


def pytest_collection_modifyitems(config, items):
    """Skip tests based on selected apps and VBA trust access."""
    # If VBA trust access is not available, skip all tests
    if getattr(config, "_vba_trust_skip_all", False):
        skip_reason = "VBA trust access configuration required"
        for item in items:
            item.add_marker(pytest.mark.skip(reason=skip_reason))
        return

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


@pytest.fixture(autouse=True, scope="function")
def com_cleanup():
    """Ensure COM objects are cleaned up after each test."""
    yield
    # Force cleanup of any remaining COM objects
    try:
        import gc
        import pythoncom

        gc.collect()
        # Try to uninitialize COM (may not always work)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    except ImportError:
        pass


@pytest.fixture
def vba_app():
    """VBA application fixture - will be parametrized by pytest_generate_tests."""
    # This fixture body will never execute because pytest_generate_tests
    # will parametrize it with actual values
    pass


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


@pytest.fixture
def powerpoint_only(request):
    """Check if running in PowerPoint-only mode."""
    selected = request.getfixturevalue("selected_apps")
    return selected == ["powerpoint"]
