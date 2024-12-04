# vba-edit
 Enable seamless MS Office VBA code editing in preferred editor or IDE (facilitating the use of coding assistants and version control workflows)

[![CI](https://github.com/markuskiller/vba-edit/workflows/test.yaml/badge.svg?branch=dev)](https://github.com/markuskiller/vba-edit/)
[![PyPI - Version](https://img.shields.io/pypi/v/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![PyPI - Downloads](https://img.shields.io/pypi/dm/vba-edit.svg)](https://pypi.org/project/vba-edit)
[![License](https://img.shields.io/badge/License-BSD_3--Clause-blue.svg)](https://opensource.org/licenses/BSD-3-Clause)

> [!NOTE] 
> This project is heavily inspired by code from ``xlwings vba edit``, actively maintained and developed by the [xlwings-Project](https://www.xlwings.org/) under the BSD 3-Clause License. We use the name ``xlwings`` solely to give credit to the orginal author and to refer to existing video tutorials on the subject of vba editing. This does not imply endorsement or sponsorship by the original authors or contributors.

## Quickstart

### Installation

To install `vba-edit`, you can use ``pip``:

```sh
pip install vba-edit
```

or ``uv pip``:

```sh
uv pip install vba-edit
```

### Usage

- [Working with MS Word VBA code](#working-with-ms-word-vba-code)
- [Working with MS Excel VBA code](#working-with-ms-excel-vba-code)
- ... (work in progress)


#### Working with MS Word VBA code

```
usage: word-vba [-h] {edit,import,export} ...

vba-edit v0.1.0 (word-vba)

A command-line tool suite for managing VBA content in MS Office documents.

WORD-VBA allows you to edit, import, and export VBA content from Word documents.
If no file is specified, the tool will attempt to use the currently active Word document.

Commands:
    edit    Edit VBA content in Word document
    import  Import VBA content into Word document
    export  Export VBA content from Word document

Examples:
    word-vba edit   <--- uses active Word document and current directory for exported 
                         VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                         active Word document

    word-vba import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
    word-vba export --file "C:/path/to/document.docx" --encoding cp850 --save-metadata

IMPORTANT: 
           [!] It's early days. Use with care and backup your imortant macro-enabled
               MS Office documents before using them with this tool!

               First tests have been very promissing. Feedback appreciated via
               github issues. 

           [!] This tool requires "Trust access to the VBA project object model" 
               enabled in Word.

positional arguments:
  {edit,import,export}
    edit                Edit VBA content in Word document
    import              Import VBA content into Word document
    export              Export VBA content from Word document

options:
  -h, --help            Show this help message and exit
```
##### EDIT

```sh
word-vba edit
```

Updates the VBA modules of the active (or specified) MS Word document from their local exports every time you hit save. If you run this for the first time, the modules will be exported from MS Word into your current working directory.

```
usage: word-vba edit [-h] [--encoding ENCODING | --detect-encoding] 
                     [--file FILE] [--vba-directory VBA_DIRECTORY] 
                     [--verbose]

options:
  -h, --help            show this help message and exit
  --encoding ENCODING, -e ENCODING
                        Encoding to be used when reading VBA files from Word document (default: cp1252)
  --detect-encoding, -d
                        Auto-detect input encoding for VBA files exported from Word document
  --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
  --vba-directory VBA_DIRECTORY
                        Directory to export VBA files to (optional, defaults to current directory)
  --verbose, -v         Enable verbose logging output
```

##### EXPORT

```
word-vba export
```
Overwrites the local version of the modules with those from the active (or specified) Word document.

```
usage: word-vba export [-h] [--save-metadata] [--encoding ENCODING | --detect-encoding] 
                       [--file FILE] [--vba-directory VBA_DIRECTORY] [--verbose]

options:
  -h, --help            show this help message and exit
  --save-metadata, -m   Save metadata file with character encoding information (default: False)
  --encoding ENCODING, -e ENCODING
                        Encoding to be used when reading VBA files from Word document (default: cp1252)
  --detect-encoding, -d
                        Auto-detect input encoding for VBA files exported from Word document
  --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
  --vba-directory VBA_DIRECTORY
                        Directory to export VBA files to (optional, defaults to current directory)
  --verbose, -v         Enable verbose logging output
```

##### IMPORT

```
word-vba import
```

Overwrites the VBA modules in the active (or specified) Word document with their local versions.

```
usage: word-vba import [-h] [--encoding ENCODING] [--file FILE] 
                       [--vba-directory VBA_DIRECTORY] [--verbose]

options:
  -h, --help            show this help message and exit
  --encoding ENCODING, -e ENCODING
                        Encoding to be used when writing VBA files back into Word document (default: cp1252)
  --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
  --vba-directory VBA_DIRECTORY
                        Directory to export VBA files to (optional, defaults to current directory)
  --verbose, -v         Enable verbose logging output
```


> [!NOTE]  
> Whenever you change something in the VBA editor (such as the layout of a form or the properties of a module), you have to run ``word-vba export``.

> [!IMPORTANT]  
> Requires "Trust access to the VBA project object model" enabled.

#### Working with MS Excel VBA code

```sh
excel-vba edit
```

Updates the VBA modules of the active (or specified) MS Excel document from their local exports every time you hit save. If you run this for the first time, the modules will be exported from MS Excel into your current working directory.

> [!NOTE] 
> The ``--file/-f`` flag allows you to specify a file path instead of using the active document.

```
excel-vba export
```
Overwrites the local version of the modules with those from the active (or specified) Excel document.

```
excel-vba import
```

Overwrites the VBA modules in the active (or specified) Excel document with their local versions.

> [!NOTE]  
> Whenever you change something in the VBA editor (such as the layout of a form or the properties of a module), you have to run ``excel-vba export``.

> [!IMPORTANT]  
> Requires "Trust access to the VBA project object model" enabled.

> [!TIP]
> Watch the excellent [``xlwings vba edit`` walkthrough on Youtube](https://www.youtube.com/watch?v=xoO-Fx0fTpM). The ``excel-vba edit`` command calls ``xlwings vba edit`` if ``xlwings`` is installed and provides a rudimentary fallback, in case it is not installed. If you often work with Excel-VBA-Code, make sure that [``xlwings``](https://www.xlwings.org/) is installed:
>
> ```sh
> pip install xlwings
> ```
> 
> or ``uv pip``:
>
> ```sh
> uv pip install xlwings
> ```