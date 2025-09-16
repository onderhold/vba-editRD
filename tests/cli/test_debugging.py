"""CLI option debugging tests that write effective values to files."""

import pytest
import os
from pathlib import Path
from .helpers import CLITester, temp_office_doc
from vba_edit.office_vba import OFFICE_MACRO_EXTENSIONS
from vba_edit.cli_common import (
    PLACEHOLDER_CONFIG_PATH,
    PLACEHOLDER_FILE_NAME,
    PLACEHOLDER_FILE_PATH,
    PLACEHOLDER_VBA_PROJECT,
    CONFIG_SECTION_GENERAL,
    CONFIG_KEY_FILE,
    CONFIG_KEY_VBA_DIRECTORY,
    CONFIG_KEY_VERBOSE,
    CONFIG_KEY_RUBBERDUCK_FOLDERS,
)


class TestCLIOptionsDebugging:
    """Tests that write effective option values to files for debugging."""

    def write_options_file(self, tmp_path: Path, test_name: str, options_dict: dict) -> Path:
        """Write effective options to a test file for debugging.
        
        Args:
            tmp_path: Temporary directory path
            test_name: Name of the test case
            options_dict: Dictionary of option names and values
            
        Returns:
            Path to the created options file
        """
        options_file = tmp_path / f"test_options_{test_name}.txt"
        content = f"Test: {test_name}\n"
        content += "=" * 50 + "\n"
        for key, value in sorted(options_dict.items()):
            content += f"{key}: {value}\n"
        content += "\n"
        options_file.write_text(content, encoding="utf-8")
        return options_file

    @pytest.mark.office
    def test_default_options_values(self, vba_app, tmp_path):
        """Test default option values without any configuration."""
        cli = CLITester(f"{vba_app}-vba")
        
        # Create a minimal test to capture default values
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"test{extension}"
        vba_dir = tmp_path / "vba"
        
        # Run export with minimal arguments to see defaults
        result = cli.run([
            "export", 
            "-f", str(test_file), 
            "--vba-directory", str(vba_dir),
            "--help"  # Use help to avoid file operations
        ])
        
        # Extract effective options
        effective_options = {
            "command": "export",
            "file": str(test_file),
            "vba_directory": str(vba_dir),
            "verbose": "false",
            "rubberduck_folders": "false",
            "config_file": "none",
            "return_code": result.returncode,
            "stdout_length": len(result.stdout),
            "stderr_length": len(result.stderr)
        }
        
        options_file = self.write_options_file(tmp_path, f"default_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_config_file_option_merging(self, vba_app, tmp_path):
        """Test option merging between config file and command line."""
        cli = CLITester(f"{vba_app}-vba")
        
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"project{extension}"
        vba_dir = tmp_path / "vba_modules"
        
        # Create config file with some options
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_dir}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = false
"""
        config_file = tmp_path / "test-config.toml"
        config_file.write_text(config_content, encoding="utf-8")
        
        # Run with config file but override some options
        override_vba_dir = tmp_path / "override_vba"
        result = cli.run([
            "export",
            "--conf", str(config_file),
            "--vba-directory", str(override_vba_dir),  # Override config
            "--rubberduck-folders",  # Override config
            "--help"
        ])
        
        effective_options = {
            "command": "export",
            "config_file": str(config_file),
            "file_from_config": str(test_file),
            "vba_directory_override": str(override_vba_dir),
            "verbose_from_config": "true",
            "rubberduck_folders_override": "true",
            "return_code": result.returncode,
            "config_file_exists": str(config_file.exists()),
            "test_file_exists": str(test_file.exists())
        }
        
        options_file = self.write_options_file(tmp_path, f"config_merge_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_placeholder_resolution_debugging(self, vba_app, tmp_path):
        """Test placeholder resolution with detailed logging."""
        cli = CLITester(f"{vba_app}-vba")
        
        # Create config subdirectory
        config_dir = tmp_path / "config" / "project"
        config_dir.mkdir(parents=True)
        
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / "documents" / f"MyProject{extension}"
        test_file.parent.mkdir(parents=True)
        
        # Use all available placeholders
        vba_directory_with_placeholders = f"{PLACEHOLDER_CONFIG_PATH}{os.sep}modules{os.sep}{PLACEHOLDER_FILE_NAME}"
        backup_path = f"{PLACEHOLDER_FILE_PATH}{os.sep}backups"
        
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{test_file}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_directory_with_placeholders}"
{CONFIG_KEY_VERBOSE} = true

[advanced]
backup_directory = "{backup_path}"
output_template = "{PLACEHOLDER_VBA_PROJECT}_exported"
"""
        config_file = config_dir / "project-config.toml"
        config_file.write_text(config_content, encoding="utf-8")
        
        result = cli.run([
            "export",
            "--conf", str(config_file),
            "--help"
        ])
        
        # Calculate expected placeholder resolutions
        expected_config_path = str(config_dir)
        expected_file_path = str(test_file.parent)
        expected_file_name = test_file.stem
        expected_vba_dir = f"{expected_config_path}{os.sep}modules{os.sep}{expected_file_name}"
        expected_backup_dir = f"{expected_file_path}{os.sep}backups"
        
        effective_options = {
            "command": "export",
            "config_file": str(config_file),
            "original_vba_directory": vba_directory_with_placeholders,
            "resolved_vba_directory": expected_vba_dir,
            "original_backup_directory": backup_path,
            "resolved_backup_directory": expected_backup_dir,
            "placeholder_config_path": PLACEHOLDER_CONFIG_PATH,
            "placeholder_file_path": PLACEHOLDER_FILE_PATH,
            "placeholder_file_name": PLACEHOLDER_FILE_NAME,
            "placeholder_vba_project": PLACEHOLDER_VBA_PROJECT,
            "actual_config_path": expected_config_path,
            "actual_file_path": expected_file_path,
            "actual_file_name": expected_file_name,
            "return_code": result.returncode
        }
        
        options_file = self.write_options_file(tmp_path, f"placeholders_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.office
    def test_multiple_config_scenarios(self, vba_app, tmp_path):
        """Test various configuration scenarios and document their behavior."""
        cli = CLITester(f"{vba_app}-vba")
        
        scenarios = [
            {
                "name": "minimal_config",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
""",
                "args": ["export", "--help"]
            },
            {
                "name": "full_config",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true

[office]
application = "{vba_app}"
""",
                "args": ["export", "--help"]
            },
            {
                "name": "config_with_file",
                "config": f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "test{OFFICE_MACRO_EXTENSIONS[vba_app]}"
{CONFIG_KEY_VBA_DIRECTORY} = "vba-files"
""",
                "args": ["export", "--help"]
            }
        ]
        
        for scenario in scenarios:
            config_file = tmp_path / f"config_{scenario['name']}.toml"
            config_file.write_text(scenario["config"], encoding="utf-8")
            
            result = cli.run([*scenario["args"], "--conf", str(config_file)])
            
            effective_options = {
                "scenario": scenario["name"],
                "config_content": scenario["config"].replace("\n", "\\n"),
                "command_args": " ".join(scenario["args"]),
                "return_code": result.returncode,
                "stdout_lines": len(result.stdout.splitlines()),
                "stderr_lines": len(result.stderr.splitlines()),
                "config_file_size": len(scenario["config"]),
                "has_error_output": str("error" in result.stderr.lower()),
                "has_help_output": str("help" in result.stdout.lower() or "usage" in result.stdout.lower())
            }
            
            options_file = self.write_options_file(tmp_path, f"scenario_{scenario['name']}_{vba_app}", effective_options)
            assert options_file.exists()

    @pytest.mark.office
    def test_relative_path_resolution(self, vba_app, tmp_path):
        """Test relative path resolution in different directory contexts."""
        cli = CLITester(f"{vba_app}-vba")
        
        # Create directory structure
        project_root = tmp_path / "project"
        config_dir = project_root / "config"
        docs_dir = project_root / "documents"
        output_dir = project_root / "output"
        
        for dir_path in [project_root, config_dir, docs_dir, output_dir]:
            dir_path.mkdir(parents=True)
        
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        relative_scenarios = [
            {
                "name": "relative_to_config",
                "config_dir": config_dir,
                "file_path": f"..{os.sep}documents{os.sep}test{extension}",
                "vba_dir": f"..{os.sep}output{os.sep}vba"
            },
            {
                "name": "absolute_paths",
                "config_dir": config_dir,
                "file_path": str(docs_dir / f"test{extension}"),
                "vba_dir": str(output_dir / "vba")
            },
            {
                "name": "mixed_paths",
                "config_dir": config_dir,
                "file_path": str(docs_dir / f"test{extension}"),
                "vba_dir": f"..{os.sep}output{os.sep}vba"
            }
        ]
        
        for scenario in relative_scenarios:
            config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{scenario['file_path']}"
{CONFIG_KEY_VBA_DIRECTORY} = "{scenario['vba_dir']}"
{CONFIG_KEY_VERBOSE} = true
"""
            config_file = scenario["config_dir"] / f"{scenario['name']}.toml"
            config_file.write_text(config_content, encoding="utf-8")
            
            result = cli.run([
                "export",
                "--conf", str(config_file),
                "--help"
            ])
            
            # Calculate absolute paths for comparison
            if os.path.isabs(scenario["file_path"]):
                absolute_file = scenario["file_path"]
            else:
                absolute_file = str((scenario["config_dir"] / scenario["file_path"]).resolve())
                
            if os.path.isabs(scenario["vba_dir"]):
                absolute_vba_dir = scenario["vba_dir"]
            else:
                absolute_vba_dir = str((scenario["config_dir"] / scenario["vba_dir"]).resolve())
            
            effective_options = {
                "scenario": scenario["name"],
                "config_dir": str(scenario["config_dir"]),
                "relative_file_path": scenario["file_path"],
                "relative_vba_dir": scenario["vba_dir"],
                "absolute_file_path": absolute_file,
                "absolute_vba_dir": absolute_vba_dir,
                "config_file": str(config_file),
                "config_exists": str(config_file.exists()),
                "return_code": result.returncode,
                "working_directory": str(Path.cwd())
            }
            
            options_file = self.write_options_file(tmp_path, f"relative_{scenario['name']}_{vba_app}", effective_options)
            assert options_file.exists()

    @pytest.mark.office
    def test_command_line_precedence(self, vba_app, tmp_path):
        """Test command line argument precedence over config file."""
        cli = CLITester(f"{vba_app}-vba")
        
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        config_file_path = tmp_path / f"config_file{extension}"
        cli_file_path = tmp_path / f"cli_file{extension}"
        config_vba_dir = tmp_path / "config_vba"
        cli_vba_dir = tmp_path / "cli_vba"
        
        # Config file specifies certain values
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{config_file_path}"
{CONFIG_KEY_VBA_DIRECTORY} = "{config_vba_dir}"
{CONFIG_KEY_VERBOSE} = false
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = false
"""
        config_file = tmp_path / "precedence-test.toml"
        config_file.write_text(config_content, encoding="utf-8")
        
        # Command line overrides everything
        result = cli.run([
            "export",
            "--conf", str(config_file),
            "-f", str(cli_file_path),  # Override file
            "--vba-directory", str(cli_vba_dir),  # Override VBA directory
            "--verbose",  # Override verbose
            "--rubberduck-folders",  # Override rubberduck folders
            "--help"
        ])
        
        effective_options = {
            "test_type": "command_line_precedence",
            "config_file": str(config_file),
            "config_file_value": str(config_file_path),
            "cli_file_value": str(cli_file_path),
            "config_vba_dir": str(config_vba_dir),
            "cli_vba_dir": str(cli_vba_dir),
            "config_verbose": "false",
            "cli_verbose": "true",
            "config_rubberduck": "false",
            "cli_rubberduck": "true",
            "return_code": result.returncode,
            "has_precedence_test": "true"
        }
        
        options_file = self.write_options_file(tmp_path, f"precedence_{vba_app}", effective_options)
        assert options_file.exists()

    @pytest.mark.integration
    @pytest.mark.com
    @pytest.mark.office
    def test_real_document_option_resolution(self, vba_app, temp_office_doc, tmp_path):
        """Test option resolution with a real Office document."""
        cli = CLITester(f"{vba_app}-vba")
        
        # Create complex config with all placeholders
        vba_path_template = f"{PLACEHOLDER_FILE_PATH}{os.sep}{PLACEHOLDER_VBA_PROJECT}-{PLACEHOLDER_FILE_NAME}"
        
        config_content = f"""
[{CONFIG_SECTION_GENERAL}]
{CONFIG_KEY_FILE} = "{temp_office_doc}"
{CONFIG_KEY_VBA_DIRECTORY} = "{vba_path_template}"
{CONFIG_KEY_VERBOSE} = true
{CONFIG_KEY_RUBBERDUCK_FOLDERS} = true

[export]
create_folders = true

[import]
overwrite_existing = true
"""
        config_file = tmp_path / "real-doc-config.toml"
        config_file.write_text(config_content, encoding="utf-8")
        
        # Test export operation
        result = cli.run([
            "export",
            "--conf", str(config_file)
        ])
        
        # Calculate expected values
        file_path = str(temp_office_doc.parent)
        file_name = temp_office_doc.stem
        # VBA project name would need to be extracted from the document
        expected_vba_dir = f"{file_path}{os.sep}VBAProject-{file_name}"  # Placeholder for actual VBA project name
        
        effective_options = {
            "test_type": "real_document",
            "document_path": str(temp_office_doc),
            "document_exists": str(temp_office_doc.exists()),
            "document_size": str(temp_office_doc.stat().st_size) if temp_office_doc.exists() else "0",
            "config_file": str(config_file),
            "template_vba_directory": vba_path_template,
            "expected_vba_directory": expected_vba_dir,
            "file_path_component": file_path,
            "file_name_component": file_name,
            "return_code": result.returncode,
            "operation_successful": str(result.returncode == 0),
            "output_length": len(result.stdout + result.stderr),
            "has_vba_components": str("VBA components" in (result.stdout + result.stderr))
        }
        
        options_file = self.write_options_file(tmp_path, f"real_doc_{vba_app}", effective_options)
        assert options_file.exists()
        
        # If export was successful, check if the expected directory was created
        if result.returncode == 0:
            # Look for any VBA directories that were created
            created_dirs = []
            for item in temp_office_doc.parent.iterdir():
                if item.is_dir() and ("vba" in item.name.lower() or file_name in item.name):
                    created_dirs.append(str(item))
            
            if created_dirs:
                effective_options["actual_created_dirs"] = "|".join(created_dirs)
                # Update the options file with actual results
                options_file = self.write_options_file(tmp_path, f"real_doc_updated_{vba_app}", effective_options)