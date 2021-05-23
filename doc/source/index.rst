.. figure:: _static/logo.png
   :align: center

   pylightxl
   `pypi <https://pypi.org/project/pylightxl/>`_ | `github <https://github.com/PydPiper/pylightxl>`_


.. |br| raw:: html

   <br />

Welcome to pylightxl documentation
----------------------------------

A light weight, zero dependency (only standard libs used), to the point (no bells and whistles) Microsoft Excel
reader/writer python 2.7.18 - 3+ library. Although there are
several excellent read/write options out there (`python-excel.org <https://www.python-excel.org/>`_ or
`excelpython.org <https://www.excelpython.org/>`_) pylightxl focused on the following key features:

- **Zero non-standard library dependencies!** |br|

    - No compatibility/version control issues.

- **Python2.7.18 to Python3+ support for life!** |br|

    - Don't worry about which python version you are using, pylightxl will support it for life

- **Light-weight single source code file** |br|

    - Want your project remain truly dependent-less? Copy the single source file into your project without any extra
      dependency issues or setup.
    - Do you struggle with other libraries weighing your projects down due to their very large size? Pylightxl's
      single source file size and zero dependency will not weight your project down (preferable for django apps)
    - Do you struggle with ``pyinstaller`` or other ``exe`` wrappers failing to build or building to very large
      packages? Pylightxl will not cause any build errors and will not add to your build size since it has zero
      dependencies and a small lib size.
    - Do you struggle with download restrictions at your company? Copy the entire pylightxl source from 1 single file
      and use it in your project.

- **100% test-driven development for highest reliability/maintainability that aims for 100% coverage on all supported versions**

    - Pylightxl aims to test all of its features, however unforeseen edge cases can occur when receiving excel files
      created by non-microsoft excel. We actively monitor issues to add support for these edge cases should they arise.

- **API aimed to be user friendly and intuitive and well documented. Structure: database > worksheet > indexing**

    - ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1')``
    - ``db.ws('Sheet1').row(1)`` or ``db.ws('Sheet1').col(1)``

High-Level Feature Summary
--------------------------
- **Reader**

    - supports Microsoft Excel 2004+ files (``.xlsx``, ``.xlsm``) and ``.csv`` files
    - read files via str path, pathlib path or file objects
    - read all or selective sheets
    - read type converted cell value (string, int, float), formula, comments, and named ranges

- **Database**

    - call cell value by row/col ID, excel address, or range
    - call an entire row/col or a semi-structured table based on user-defined headers

- **Writer**

    - write to new excel file (write excel files without having excel on your machine)
    - write to existing excel files (see limitations below)

Limitations
-----------
Although every effort was made to support a variety of users, the following limitations should be read carefully:

- Does not support ``.xls`` files (Microsoft Excel 2003 and older files)

- Writer does not support anything other than cell data (no graphs, images, macros, formatting)

- Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation), please use 64-bit if
  more data storage is required.


.. toctree::
   :maxdepth: 2
   :numbered:

   installation
   quickstart
   sourcecode/index
   solutions
   revlog
   license
   codeofconduct


Support Content Creator
-----------------------
If you have enjoy using this library please consider supporting it by one or more of the following ways:

- Star us on github! `Github <https://github.com/PydPiper/pylightxl>`_

- Sponsor via `Tidelift <https://tidelift.com/>`_

- Sponsor via `Patreon <https://www.patreon.com/pylightxl>`_
