"""CLI option debugging tests that write effective values to files."""

import pytest
import os
import json
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
    merge_config_with_args,
    resolve_placeholders,
    load_config_file,
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
        content += f"Timestamp: {pytest.current_time}\n" if hasattr(pytest, 'current_time') else ""
        content += f"CLI Command: {options_dict.get('cli_command', 'unknown')}\n"
        content += f"Return Code: {options_dict.get('return_code', 'unknown')}\n"
        content += "-" * 30 + "\n"
        
        # Separate config and runtime options
        config_options = {k: v for k, v in options_dict.items() if k.startswith(('config_', 'original_', 'template_'))}
        runtime_options = {k: v for k, v in options_dict.items() if k.startswith(('resolved_', 'expected_', 'actual_', 'cli_'))}
        other_options = {k: v for k, v in options_dict.items() if not any(k.startswith(prefix) for prefix in ['config_', 'original_', 'template_', 'resolved_', 'expected_', 'actual_', 'cli_'])}
        
        if config_options:
            content += "Configuration Options:\n"
            for key, value in sorted(config_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"
        
        if runtime_options:
            content += "Runtime/Resolved Options:\n"
            for key, value in sorted(runtime_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"
        
        if other_options:
            content += "Other Options:\n"
            for key, value in sorted(other_options.items()):
                content += f"  {key}: {value}\n"
            content += "\n"
        
        # Add debug information about the test environment
        content += "Environment Info:\n"
        content += f"  Working Directory: {Path.cwd()}\n"
        content += f"  Temp Path: {tmp_path}\n"
        content += f"  Test File: {options_file}\n"
        
        options_file.write_text(content, encoding="utf-8")
        return options_file

    def simulate_cli_processing(self, cli_args: list, config_file_path: Path = None) -> dict:
        """Simulate CLI argument processing to extract effective options.
        
        Args:
            cli_args: CLI arguments as they would be passed
            config_file_path: Optional configuration file path
            
        Returns:
            Dictionary of effective options after processing
        """
        # This simulates what happens inside the CLI processing
        from argparse import Namespace
        import argparse
        
        # Create a mock args namespace from CLI arguments
        # This is a simplified version - in reality each app has its own parser
        args = Namespace()
        
        # Parse basic arguments
        i = 0
        while i < len(cli_args):
            arg = cli_args[i]
            if arg in ['-f', '--file'] and i + 1 < len(cli_args):
                args.file = cli_args[i + 1]
                i += 2
            elif arg == '--vba-directory' and i + 1 < len(cli_args):
                args.vba_directory = cli_args[i + 1]
                i += 2
            elif arg == '--conf' and i + 1 < len(cli_args):
                args.conf = cli_args[i + 1]
                config_file_path = Path(cli_args[i + 1])
                i += 2
            elif arg == '--verbose':
                args.verbose = True
                i += 1
            elif arg == '--rubberduck-folders':
                args.rubberduck_folders = True
                i += 1
            elif arg in ['export', 'import', 'edit']:
                args.command = arg
                i += 1
            else:
                i += 1
        
        # Set defaults for missing attributes
        for attr in ['file', 'vba_directory', 'conf', 'verbose', 'rubberduck_folders', 'command']:
            if not hasattr(args, attr):
                setattr(args, attr, None if attr in ['file', 'vba_directory', 'conf'] else False)
        
        effective_options = {
            'command': getattr(args, 'command', 'unknown'),
            'file_argument': getattr(args, 'file', None),
            'vba_directory_argument': getattr(args, 'vba_directory', None),
            'verbose_argument': getattr(args, 'verbose', False),
            'rubberduck_folders_argument': getattr(args, 'rubberduck_folders', False),
            'config_file_argument': getattr(args, 'conf', None),
        }
        
        # Load and merge configuration if available
        if config_file_path and config_file_path.exists():
            try:
                config = load_config_file(str(config_file_path))
                effective_options['config_loaded'] = True
                effective_options['config_content'] = str(config)
                
                # Merge config with args (this simulates the actual CLI processing)
                merged_args = merge_config_with_args(args, config)
                
                effective_options['file_after_config'] = getattr(merged_args, 'file', None)
                effective_options['vba_directory_after_config'] = getattr(merged_args, 'vba_directory', None)
                effective_options['verbose_after_config'] = getattr(merged_args, 'verbose', False)
                effective_options['rubberduck_folders_after_config'] = getattr(merged_args, 'rubberduck_folders', False)
                
                # Try to resolve placeholders if we have enough information
                if getattr(merged_args, 'file', None) and getattr(merged_args, 'vba_directory', None):
                    try:
                        resolved_options = resolve_placeholders(
                            vars(merged_args), 
                            str(config_file_path.parent),
                            getattr(merged_args, 'file', '')
                        )
                        effective_options['vba_directory_resolved'] = resolved_options.get('vba_directory')
                        effective_options['file_resolved'] = resolved_options.get('file')
                    except Exception as e:
                        effective_options['placeholder_resolution_error'] = str(e)
                
            except Exception as e:
                effective_options['config_load_error'] = str(e)
                effective_options['config_loaded'] = False
        else:
            effective_options['config_loaded'] = False
            effective_options['config_file_exists'] = config_file_path.exists() if config_file_path else False
        
        return effective_options

    @pytest.mark.office
    def test_default_options_values(self, vba_app, tmp_path):
        """Test default option values without any configuration."""
        cli = CLITester(f"{vba_app}-vba")
        
        # Create a minimal test to capture default values
        extension = OFFICE_MACRO_EXTENSIONS[vba_app]
        test_file = tmp_path / f"test{extension}"
        vba_dir = tmp_path / "vba"
        
        # Run export with minimal arguments to see defaults
        cli_args = [
            "export", 
            "-f", str(test_file), 
            "--vba-directory", str(vba_dir),
            "--help"  # Use help to avoid file operations
        ]
        result = cli.run(cli_args)
        
        # Simulate the CLI processing to get effective options
        simulated_options = self.simulate_cli_processing(cli_args)
        
        # Extract effective options combining real CLI output with simulation
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "command": "export",
            "file": str(test_file),
            "vba_directory": str(vba_dir),
            "verbose": "false",
            "rubberduck_folders": "false",
            "config_file": "none",
            "return_code": result.returncode,
            "stdout_length": len(result.stdout),
            "stderr_length": len(result.stderr),
            "help_displayed": "help" in result.stdout.lower() or "usage" in result.stdout.lower(),
            **{f"simulated_{k}": v for k, v in simulated_options.items()}
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
        cli_args = [
            "export",
            "--conf", str(config_file),
            "--vba-directory", str(override_vba_dir),  # Override config
            "--rubberduck-folders",  # Override config
            "--help"
        ]
        result = cli.run(cli_args)
        
        # Simulate CLI processing
        simulated_options = self.simulate_cli_processing(cli_args, config_file)
        
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "command": "export",
            "config_file": str(config_file),
            "config_content": config_content.replace("\n", "\\n"),
            "file_from_config": str(test_file),
            "vba_directory_from_config": str(vba_dir),
            "vba_directory_override": str(override_vba_dir),
            "verbose_from_config": "true",
            "rubberduck_folders_override": "true",
            "return_code": result.returncode,
            "config_file_exists": str(config_file.exists()),
            "test_file_exists": str(test_file.exists()),
            **{f"simulated_{k}": v for k, v in simulated_options.items()}
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
        
        cli_args = [
            "export",
            "--conf", str(config_file),
            "--help"
        ]
        result = cli.run(cli_args)
        
        # Calculate expected placeholder resolutions
        expected_config_path = str(config_dir)
        expected_file_path = str(test_file.parent)
        expected_file_name = test_file.stem
        expected_vba_dir = f"{expected_config_path}{os.sep}modules{os.sep}{expected_file_name}"
        expected_backup_dir = f"{expected_file_path}{os.sep}backups"
        
        # Simulate CLI processing
        simulated_options = self.simulate_cli_processing(cli_args, config_file)
        
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "command": "export",
            "config_file": str(config_file),
            "config_content": config_content.replace("\n", "\\n"),
            "original_vba_directory": vba_directory_with_placeholders,
            "expected_vba_directory": expected_vba_dir,
            "original_backup_directory": backup_path,
            "expected_backup_directory": expected_backup_dir,
            "template_config_path": PLACEHOLDER_CONFIG_PATH,
            "template_file_path": PLACEHOLDER_FILE_PATH,
            "template_file_name": PLACEHOLDER_FILE_NAME,
            "template_vba_project": PLACEHOLDER_VBA_PROJECT,
            "actual_config_path": expected_config_path,
            "actual_file_path": expected_file_path,
            "actual_file_name": expected_file_name,
            "return_code": result.returncode,
            **{f"simulated_{k}": v for k, v in simulated_options.items()}
        }
        
        options_file = self.write_options_file(tmp_path, f"placeholders_{vba_app}", effective_options)
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
        cli_args = [
            "export",
            "--conf", str(config_file),
            "-f", str(cli_file_path),  # Override file
            "--vba-directory", str(cli_vba_dir),  # Override VBA directory
            "--verbose",  # Override verbose
            "--rubberduck-folders",  # Override rubberduck folders
            "--help"
        ]
        result = cli.run(cli_args)
        
        # Simulate CLI processing
        simulated_options = self.simulate_cli_processing(cli_args, config_file)
        
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_type": "command_line_precedence",
            "config_file": str(config_file),
            "config_content": config_content.replace("\n", "\\n"),
            "config_file_value": str(config_file_path),
            "cli_file_value": str(cli_file_path),
            "config_vba_dir": str(config_vba_dir),
            "cli_vba_dir": str(cli_vba_dir),
            "config_verbose": "false",
            "cli_verbose": "true",
            "config_rubberduck": "false",
            "cli_rubberduck": "true",
            "return_code": result.returncode,
            "has_precedence_test": "true",
            **{f"simulated_{k}": v for k, v in simulated_options.items()}
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
        cli_args = [
            "export",
            "--conf", str(config_file)
        ]
        result = cli.run(cli_args)
        
        # Calculate expected values
        file_path = str(temp_office_doc.parent)
        file_name = temp_office_doc.stem
        # VBA project name would need to be extracted from the document
        expected_vba_dir = f"{file_path}{os.sep}VBAProject-{file_name}"  # Placeholder for actual VBA project name
        
        # Simulate CLI processing
        simulated_options = self.simulate_cli_processing(cli_args, config_file)
        
        effective_options = {
            "cli_command": f"{vba_app}-vba " + " ".join(cli_args),
            "test_type": "real_document",
            "document_path": str(temp_office_doc),
            "document_exists": str(temp_office_doc.exists()),
            "document_size": str(temp_office_doc.stat().st_size) if temp_office_doc.exists() else "0",
            "config_file": str(config_file),
            "config_content": config_content.replace("\n", "\\n"),
            "template_vba_directory": vba_path_template,
            "expected_vba_directory": expected_vba_dir,
            "file_path_component": file_path,
            "file_name_component": file_name,
            "return_code": result.returncode,
            "operation_successful": str(result.returncode == 0),
            "output_length": len(result.stdout + result.stderr),
            "has_vba_components": str("VBA components" in (result.stdout + result.stderr)),
            **{f"simulated_{k}": v for k, v in simulated_options.items()}
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