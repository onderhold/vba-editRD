## Possible Additional Features and Improvements

- When in edit mode, watch the Office document separately as well. I it gets changed, pause the watcher on the exported files, and diff the VBA code in the Office file with the exported files.
When the code in the Office file differs, export it, overwriting existing ones,

    **or, if this gets to difficult** (e.g. we would to have to check on new and deleted modules in the Office file, and on changed attributes): --> stop edit mode (safe mode)

- --mixed-headers: with --in-file-headers, still use --save-headers logic for user forms. When in edit mode, track changes to header files, append the code files to them and do a reimport.

- _Startup logic_: Check if the document is already open in another instance of the application.

- Using the Office applications in _invisible/silent mode_ (currently, the application window is always shown) for certain commands. Will most probably require the definition of various exit codes as well.

- **implemented:** Add CLI option to use a _config file_ (e.g. `excel-vba.toml`).
   - The config file can be specified via a command-line option (e.g. `--conf path/to/config.toml`).
   - There's no provision for a default config file, if one shall be used, it always has to be specified explicitly.
   - The config file can contain options such as the input and output file paths, the Office application to use, and any additional command-line arguments.
   - Placeholders in the config file (e.g. `{file.name}`, `{file.path}`, `{file.fullname}`) will be resolved based on the actual file being processed.
   - The config file will support sections for different settings, such as `general`, `word`, `excel`, etc.
   - Command-line arguments will override config file settings.

- _export/import/edit PowerQuery M Code_ for Excel files, maybe by a dedicated _**excel-pq**_ module

- **implemented:** _replace watchgod_ with newer library `watchfiles`

- _Restructuring the option handling_, which will go together with restructuring/tagging of test modules:
   - Some options apply to all commands: check, all 'file manipulation' commands
      - --help, --verbose, --logfile
   - Some options apply to all the 'file manipulation' commands: import, export, edit
      - --file, --vba-directory, --rubberduck-folders, --encoding, --detect-encoding, --save-headers, --in-file-headers
   - Some options apply to only certain commands:
      - check: --silent (tbd)
      - import: --silent (tbd)
      - export: --silent (tbd), --open-folder, --save-metadata

   - Some options apply to certain executables only
      - excel-vba: --xlwings, --pq-directory

- _User manual_ in addition to the --help option


##
- (very optional:)
    - implement a (readonly) virtual filesystem that mirrors the structure of the VBA project
    - exporting additional Office document properties (e.g. custom document properties, custom XML parts, references, etc.) or even the entire document (e.g. using `python-docx`, `openpyxl`, `python-pptx`, `pywin32` etc.)

- execution mode / visiblity / silent flag: vba-edit may be used in two different use cases
  - as a one-off background tool to export Office program code to the filesystem in a version control/CI/CD pipeline or similar
  - as a short-running foreground process to import some code modules once (e.g. mix-in VBA libraries)
  - as a long-running process that keeps the file system and the VBA editor in sync (edit mode)
  
    The current implementation tries to cover all thes use cases. Might it be cleaner to have separate commands or modes for these scenarios?

    - as a one-off exporting tool, it might be useful to have a command that just opens the document in readonly mode using a separate invisible instance of the Office application (irrespective of whether the document is already open in another instance), exports the code, and then ends the application again (closing the document implicitly as well).
    - when used to add some code modules once (mix-in), the document shall be opened visibly. There will be the need for the user to recompile, so the application needs to be visible and stay opened.
    - since edit mode always means that changes to the code can potentially happen concurrently in two places (VBA editor and file system), there needs to be a clear strategy for resolving conflicts and keeping both sides in sync. This is the most complex use case and might benefit from a more robust implementation, e.g. using a virtual filesystem. But even with a virtual filesystem, there's no way to control in-memory changes at neither the single module file level nor in the VBA editor. So, a last writer-wins strategy is the only feasible approach. At least, there should be a mechanism to detect conflicts of saved content and wether unsaved changes exist, and to notify the user accordingly (e.g. using hash codes for each module for saved content, and just warning messages about unsaved changes).

### Configuration File

- The configuration file is a TOML file that allows users to customize the behavior of the CLI tool. Design decisions include:
  - File format: TOML is chosen for its simplicity and readability. It looks like INI files but supports nested structures.
  - Location: The config file is specified via a command-line option; if a path is omitted, it defaults to the current working directory.
  - Structure: The config file will have sections for different settings, such as default file paths, encoding options, and application-specific settings.

- Users can specify options such as the input and output file paths, the Office application to use, and any additional command-line arguments.

### Invisible Mode

- The invisible mode allows users to run the Office application in the background without displaying an application window. This mode will spawn a new instance of the application, with certain restrictions to the command being executed: the import command may not allow a new instance if the office file in question is already opened in the current session (reminder: _**or any other session!!**_)

## PowerQuery M Code

- PowerQuery M code is stored in a different part of the Office (Excel) document than VBA code. PowerQuery M code is stored in the `xl/query` folder of the `.xlsx`, `.xlsm` or `.xlsb` file.
- To export/import PowerQuery M code, the tool will need to read/write the relevant parts of the Office document. This can be done using libraries such as `zipfile` in order to manipulate the workbook directly. Another implementation option would be to temporarily inject VBA code - decision on these implementation options will be based upon the availabilty of official APIs callable from outside of the COM-object (Office SDK?) 

