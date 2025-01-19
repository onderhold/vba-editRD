vba-edit
========

Enable seamless Microsoft Office VBA code editing in your preferred
editor or IDE, facilitating the use of coding assistants and version
control workflows.

|CI| |PyPI - Version| |PyPI - Python Version| |PyPI - Downloads|
|License|

Features
--------

- Edit VBA code in your favorite code editor or IDE
- Automatically sync changes between your editor and Office applications
- Support for Word, Excel, and Access (PowerPoint support coming in
  v0.4.0)
- Preserve form layouts and module properties
- Handle different character encodings
- Integration with version control systems
- Support for UserForms and class modules

.. note::

   Inspired by code from ``xlwings vba edit``
   (`xlwings-Project <https://www.xlwings.org/>`__) under the BSD
   3-Clause License.

Quick Start
-----------

Installation
~~~~~~~~~~~~

.. code:: bash

   pip install vba-edit

Prerequisites
~~~~~~~~~~~~~

Enable "Trust access to the VBA project object model" in your Office
application's Trust Center Settings:

1. Open your Office application
2. Go to File > Options > Trust Center > Trust Center Settings
3. Select "Macro Settings"
4. Check "Trust access to the VBA project object model"

.. note::

   In MS Access, Trust Access to VBA project object model is always
   enabled if database is stored in trusted location.

Basic Usage
~~~~~~~~~~~

Excel Example
^^^^^^^^^^^^^

1. Open your Excel workbook with VBA code
2. In your terminal, run:

.. code:: bash

   excel-vba edit

3. Edit the exported .bas, .cls, or .frm files in your preferred editor
4. Changes are automatically synced back to Excel when you save

Word Example
^^^^^^^^^^^^

.. code:: bash

   # Export VBA modules from active document
   word-vba export --vba-directory ./VBA

   # Edit and sync changes automatically
   word-vba edit --vba-directory ./VBA

   # Import changes back to document
   word-vba import --vba-directory ./VBA

Access Example
^^^^^^^^^^^^^^

.. code:: bash

   # Export VBA modules
   access-vba export --vba-directory ./src

   # Import changes
   access-vba import --vba-directory ./src

Detailed Features
-----------------

Supported File Types
~~~~~~~~~~~~~~~~~~~~

- Standard Modules (.bas)
- Class Modules (.cls)
- UserForms (.frm)
- Document/Workbook Modules

Command Line Tools
~~~~~~~~~~~~~~~~~~

The package provides separate command-line tools for each Office
application:

- ``word-vba``
- ``excel-vba``
- ``access-vba``

Each tool supports three main commands:

- ``edit``: Live sync between editor and Office (Word/Excel only)
- ``export``: Export VBA modules to files
- ``import``: Import VBA modules from files
- ``check``: Check if 'Trust Access to the VBA project object model' is
  enabled

.. note::

   The command ``python -m vba_edit.utils`` can be used to troubleshoot
   Trust Access to VBA project object model, scanning and giving
   feedback on all supported MS Office apps

Common Options
~~~~~~~~~~~~~~

.. code:: text

   --file, -f             Path to Office document (optional)
   --vba-directory        Directory for VBA files
   --encoding, -e         Specify character encoding
   --detect-encoding, -d  Auto-detect encoding
   --save-headers         Save module headers separately
   --verbose, -v          Enable detailed logging
   --logfile, -l         Enable file logging

Excel-Specific Features
~~~~~~~~~~~~~~~~~~~~~~~

For Excel users who also have xlwings installed:

.. code:: bash

   excel-vba edit -x  # Use xlwings wrapper

Best Practices
--------------

1. Always backup your Office files before using vba-edit
2. Use version control (git) to track your VBA code
3. Run ``export`` after changing form layouts or module properties
4. Use ``--save-headers`` when working with UserForms
5. Consider using ``--detect-encoding`` for non-English VBA code

Known Limitations
-----------------

- Access support is limited to import/export (no live editing)
- UserForms require ``--save-headers`` option
- PowerPoint support coming in v0.4.0
- ``--in-file-headers`` option coming soon

Links
-----

- `Homepage <https://langui.ch/current-projects/vba-edit/>`__
- `Documentation <https://github.com/markuskiller/vba-edit/blob/main/README.md>`__
- `Source Code <https://github.com/markuskiller/vba-edit>`__
- `Changelog <https://github.com/markuskiller/vba-edit/blob/main/CHANGELOG.md>`__
- `Changelog of latest dev
  version <https://github.com/markuskiller/vba-edit/blob/dev/CHANGELOG.md>`__
- `Video Tutorial <https://www.youtube.com/watch?v=xoO-Fx0fTpM>`__
  (xlwings walkthrough, with similar functionality)

License
-------

BSD 3-Clause License

Acknowledgments
---------------

This project is heavily inspired by code from ``xlwings vba edit``,
maintained by the `xlwings-Project <https://www.xlwings.org/>`__ under
the BSD 3-Clause License.

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
