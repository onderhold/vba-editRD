## Command-Line Options for VBA-edit Executable Modules

| Module | Command | Option | Short | Scope | Description |
|--------|---------|--------|-------|-------|-------------|
| **All modules** | **All commands** | `--help` | `-h` | Global | Show help message and exit (auto-generated) |
| All modules | All commands | `--version` | | Global | Show version information and exit |
| All modules | All commands | `--conf` / `--config` | | Global | Path to configuration file (TOML format) |
| All modules | All commands | `--file` | `-f` | Global | Path to Office document |
| All modules | All commands | `--vba-directory` | | Global | Directory to export/import VBA files |
| All modules | All commands | `--verbose` | `-v` | Global | Enable verbose logging output |
| All modules | All commands | `--logfile` | `-l` | Global | Enable logging to file |
| All modules | edit/import/export | `--rubberduck-folders` | | edit/import/export | Organize folders per RubberduckVBA @Folder annotations |
| All modules | edit/import/export | `--open-folder` | | edit/import/export | Open export directory in file explorer after export |
| All modules | edit/import/export | `--encoding` | `-e` | edit/import/export | Encoding for reading/writing VBA files |
| All modules | edit/import/export | `--detect-encoding` | `-d` | edit/import/export | Auto-detect file encoding for VBA files |
| All modules | edit/import/export | `--save-headers` | | edit/import/export | Save VBA component headers to separate .header files |
| All modules | edit/import/export | `--in-file-headers` | | edit/import/export | Include VBA headers directly in code files |
| All modules | export only | `--save-metadata` | `-m` | export only | Save metadata file with character encoding information |
| **excel-vba** | edit/import/export | `--xlwings` | `-x` | Excel-specific | Use wrapper for xlwings vba commands |
| excel-vba | edit/import/export | `--pq-directory` | | Excel-specific | Directory to export PowerQuery M files (Excel only) |

### Command Structure

All executable modules support the same four commands:

- `edit` - Keep VBA code synchronized between Office document and filesystem
- `import` - Import VBA code from filesystem to Office document  
- `export` - Export VBA code from Office document to filesystem
- `check` - Check if Office document contains VBA code (no additional options)

### Module Names

The project provides these executable modules:

- `excel-vba` - Excel VBA operations
- `word-vba` - Word VBA operations  
- `access-vba` - Access VBA operations
- `powerpoint-vba` - PowerPoint VBA operations

### Notes

1. **Global options** (`--help`, `--version`, `--conf`, `--file`, `--vba-directory`, `--verbose`, `--logfile`) are available on the main command level for all modules.

2. **Command-specific options** are only available for the `edit`, `import`, and `export` commands. The `check` command accepts no additional options.

3. **Excel-specific options** (`--xlwings`, `--pq-directory`) are only available in the `excel-vba` module.

4. **Mutually exclusive options**: `--encoding` and `--detect-encoding` cannot be used together; `--save-headers` and `--in-file-headers` cannot be used together.

5. **Export-only option**: `--save-metadata` is only available for the `export` command.

The architecture uses a centralized argument system in `cli_common.py` with functions like `add_common_arguments`, `add_encoding_arguments`, etc., that are applied consistently across all modules through the `OfficeVBACLI` class.