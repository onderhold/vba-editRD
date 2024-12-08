vba-edit
========

Enable seamless MS Office VBA code editing in preferred editor or IDE
(facilitating the use of coding assistants and version control
workflows)

|CI| |PyPI - Version| |PyPI - Python Version| |PyPI - Downloads|
|License|

.. note::

   This project is heavily inspired by code from ``xlwings vba edit``,
   actively maintained and developed by the
   `xlwings-Project <https://www.xlwings.org/>`__ under the BSD 3-Clause
   License. We use the name ``xlwings`` solely to give credit to the
   orginal author and to refer to existing video tutorials on the
   subject of vba editing. This does not imply endorsement or
   sponsorship by the original authors or contributors. [!IMPORTANT]
   It's early days. Use with care and backup your imortant macro-enabled
   MS Office documents before using them with this tool!

   First tests have been very promissing. Feedback appreciated via
   github issues.

Links
-----

- `Homepage <https://langui.ch/current-projects/vba-edit/>`__
- `Documentation <https://github.com/markuskiller/vba-edit/blob/main/README.md>`__
- `Source <https://github.com/markuskiller/vba-edit>`__
- `Changelog <https://github.com/markuskiller/vba-edit/blob/main/CHANGELOG.md>`__

Quickstart
----------

Installation
~~~~~~~~~~~~

To install ``vba-edit``, you can use ``pip``:

.. code:: sh

   pip install vba-edit

or ``uv pip``:

.. code:: sh

   uv pip install vba-edit

Overview command-line tools
~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code:: text

   vba-edit v0.2.0 (word-vba|excel-vba|...)

   A command-line tool suite for managing VBA content in MS Office documents.

   WORD|EXCEL|...-VBA allows you to edit, import, and export VBA content 
   from Office documents. If no file is specified, the tool will attempt
   to use the currently active document.

            usage: word-vba [-h] {edit,import,export} ...
            usage: excel-vba [-h] {edit,import,export} ...
            usage: access-vba [-h] {edit,import,export} ...      {planned in v0.3.0}
            usage: powerpoint-vba [-h] {edit,import,export} ...  {planned in v0.4.0}   

   Commands:
       edit    Edit VBA content in Office document
       import  Import VBA content into Office document
       export  Export VBA content from Office document

   Examples :                          
       word-vba  edit   <--- uses active Word document and current directory for exported 
                             VBA files (*.bas/*.cls/*.frm) & syncs changes back to the 
                             active Word document
       

       Options:

       word-vba  import -f "C:/path/to/document.docx" --vba-directory "path/to/vba/files"
       word-vba  export --file "C:/path/to/document.docx" --encoding cp850 --save-metadata
       word-vba  edit --vba-directory "path/to/vba/files" --logfile "path/to/logfile" --verbose

   positional arguments:
     {edit,import,export}
       edit                Edit VBA content in Office document
       import              Import VBA content into Office document
       export              Export VBA content from Office document

   options:
     -h, --help            Show this help message and exit

.. important::

   Requires "Trust access to the VBA project object model" enabled.
   |Trust Center Settings|

Usage
~~~~~

- `Working with MS Word VBA code <#working-with-ms-word-vba-code>`__
- `Working with MS Excel VBA code <#working-with-ms-excel-vba-code>`__
- ... (work in progress)

Working with MS Word VBA code
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

WORD-VBA EDIT COMMAND
'''''''''''''''''''''

.. code:: sh

   word-vba edit

Updates the VBA modules of the active (or specified) MS Word document
from their local exports every time you hit save. If you run this for
the first time, the modules will be exported from MS Word into your
current working directory.

.. code:: text

   usage: word-vba edit [-h] [--encoding ENCODING | --detect-encoding] 
                        [--file FILE] [--vba-directory VBA_DIRECTORY] 
                        [--verbose]

   options:
     -h, --help            show this help message and exit
     --encoding ENCODING, -e ENCODING
                           Encoding to be used when reading VBA files from Word document 
                           (default: cp1252)
     --detect-encoding, -d
                           Auto-detect input encoding for VBA files exported from Word 
                           document
     --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
     --vba-directory VBA_DIRECTORY
                           Directory to export VBA files to (optional, defaults to 
                           current directory)
     --verbose, -v         Enable verbose logging output
     --logfile [LOGFILE], -l [LOGFILE]
                           Enable logging to file. Optional path can be specified 
                           (default: vba_edit.log)

WORD-VBA EXPORT COMMAND
'''''''''''''''''''''''

.. code:: text

   word-vba export

Overwrites the local version of the modules with those from the active
(or specified) Word document.

.. code:: text

   usage: word-vba export [-h] [--save-metadata] [--encoding ENCODING | --detect-encoding] 
                          [--file FILE] [--vba-directory VBA_DIRECTORY] [--verbose]

   options:
     -h, --help            show this help message and exit
     --save-metadata, -m   Save metadata file with character encoding information 
                           (default: False)
     --encoding ENCODING, -e ENCODING
                           Encoding to be used when reading VBA files from Word 
                           document (default: cp1252)
     --detect-encoding, -d
                           Auto-detect input encoding for VBA files exported from 
                           Word document
     --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
     --vba-directory VBA_DIRECTORY
                           Directory to export VBA files to (optional, defaults to 
                           current directory)
     --verbose, -v         Enable verbose logging output
     --logfile [LOGFILE], -l [LOGFILE]
                           Enable logging to file. Optional path can be specified 
                           (default: vba_edit.log)

WORD-VBA IMPORT COMMAND
'''''''''''''''''''''''

.. code:: text

   word-vba import

Overwrites the VBA modules in the active (or specified) Word document
with their local versions.

.. code:: text

   usage: word-vba import [-h] [--encoding ENCODING] [--file FILE] 
                          [--vba-directory VBA_DIRECTORY] [--verbose]

   options:
     -h, --help            show this help message and exit
     --encoding ENCODING, -e ENCODING
                           Encoding to be used when writing VBA files back into Word 
                           document (default: cp1252)
     --file FILE, -f FILE  Path to Word document (optional, defaults to active document)
     --vba-directory VBA_DIRECTORY
                           Directory to export VBA files to (optional, defaults to 
                           current directory)
     --verbose, -v         Enable verbose logging output
     --logfile [LOGFILE], -l [LOGFILE]
                           Enable logging to file. Optional path can be specified 
                           (default: vba_edit.log)

.. note::

   Whenever you change something in the Word VBA editor (such as the
   layout of a form or the properties of a module), you have to run
   ``word-vba export``.

Working with MS Excel VBA code
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. code:: sh

   excel-vba edit

Updates the VBA modules of the active (or specified) MS Excel document
from their local exports every time you hit save. If you run this for
the first time, the modules will be exported from MS Excel into your
current working directory.

.. note::

   The ``--file/-f`` flag allows you to specify a file path instead of
   using the active document.

.. code:: text

   excel-vba export

Overwrites the local version of the modules with those from the active
(or specified) Excel document.

.. code:: text

   excel-vba import

Overwrites the VBA modules in the active (or specified) Excel document
with their local versions.

.. note::

   Whenever you change something in the VBA editor (such as the layout
   of a form or the properties of a module), you have to run
   ``excel-vba export``.

Video Tutorial
~~~~~~~~~~~~~~

.. tip::

   Watch the excellent `xlwings vba edit walkthrough on
   Youtube <https://www.youtube.com/watch?v=xoO-Fx0fTpM>`__. The
   ``excel-vba edit|import|export`` should do more or less what
   ``xlwings vba edit|import|edit`` does (with the exception that
   currently VBA files are not (yet) deleted from the Excel file if they
   are deleted in edit mode). If you prefer to use ``xlwings`` & it is
   installed, it is possible to use a wrapper for ``xlwings`` by typing
   ``excel-vba edit|import|export -x``. In this case, there is an
   additional command-line option ``--vba-directory``. If you often work
   with Excel-VBA-Code, make sure that
   `xlwings <https://www.xlwings.org/>`__ is installed:

   .. code:: sh

      pip install xlwings

   or ``uv pip``:

   .. code:: sh

      uv pip install xlwings

.. |CI| image:: https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml/badge.svg
   :target: https://github.com/markuskiller/vba-edit/actions/workflows/test.yaml
.. |PyPI - Version| image:: https://img.shields.io/pypi/v/vba-edit.svg
   :target: https://pypi.org/project/vba-edit
.. |PyPI - Python Version| image:: https://img.shields.io/pypi/pyversions/vba-edit.svg
   :target: https://pypi.org/project/vba-edit
.. |PyPI - Downloads| image:: https://img.shields.io/pypi/dm/vba-edit
   :target: https://pypi.org/project/vba-edit
.. |License| image:: https://img.shields.io/badge/License-BSD_3--Clause-blue.svg
   :target: https://opensource.org/licenses/BSD-3-Clause
.. |Trust Center Settings| image:: https://langui.ch/wp/wp-content/uploads/2024/12/trust_center_vba_object_model_settings.png
