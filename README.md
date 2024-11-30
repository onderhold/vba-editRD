# vba-edit
 Enable seamless MS Office VBA code editing in preferred editor or IDE (facilitating the use of version control workflows)

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

```sh
word-vba edit
```

Updates the VBA modules of the active (or specified) MS Word document from their local exports every time you hit save. If you run this for the first time, the modules will be exported from MS Word into your current working directory.

> [!NOTE] 
> The ``--file/-f`` flag allows you to specify a file path instead of using the active document.

```
word-vba export
```
Overwrites the local version of the modules with those from the active (or specified) Word document.

```
word-vba import
```

Overwrites the VBA modules in the active (or specified) Word document with their local versions.

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
> Watch the excellent [``xlwings vba edit`` walkthrough on Youtube](https://www.youtube.com/watch?v=xoO-Fx0fTpM). The ``excel-vba edit`` command calls ``xlwings vba edit`` if the optional dependency ``xlwings`` is installed and provides a rudimentary fallback, in case it is not installed. If you often work with Excel-VBA-Code, make sure that [``xlwings``](https://www.xlwings.org/) is installed:
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