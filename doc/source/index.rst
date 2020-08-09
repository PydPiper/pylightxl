.. figure:: _static/logo.png
   :align: center

   pylightxl
   `pypi <https://pypi.org/project/pylightxl/>`_ | `github <https://github.com/PydPiper/pylightxl>`_


.. |br| raw:: html

   <br />

Welcome to pylightxl documentation
----------------------------------

A light weight Microsoft Excel File reader. Although there are
several excellent read/write options out there (`python-excel.org <https://www.python-excel.org/>`_ or
`excelpython.org <https://www.excelpython.org/>`_) pylightxl focused on the following key features:

- **Zero non-standard library dependencies** |br|
  No compatibility/version control issues.

- **Light-weight single source code file that supports both Python37 and Python27.** |br|
  Single source files that can easily be copied directly into a project for true zero-dependency. |br|
  Great for those that have installation/download restrictions. |br|
  In addition the library's size and zero dependency makes this library pyinstaller compilation small and easy!

- **100% test-driven development for highest reliability/maintainability with 100% coverage on all supported versions**

- **API aimed to be user friendly and intuitive. Structure: database > worksheet > indexing**
  example: ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1')``

High-Level Feature Summary
--------------------------
- Read excel files (``.xlsx``, ``.xlsm``), all sheets or selective few for speed/memory management

- Index cell data by row/col number or address

- Calling an entire row/col of data returns an easy to use list output: ``db.ws('Sheet1').row(1)`` or ``db.ws('Sheet1').rows``

- Worksheet data size is consistent for each row/col. Any data that is empty will return a '' (default empty cell can be updated)

- Write to existing or now spreadsheets


Limitations
-----------
Although every effort was made to support a variety of users, the following limitations should be read carefully:

- Does not support ``.xls`` files (Microsoft Excel 2003 and older files)

- Writer does not support anything other than cell data (no graphs, images, macros, formatting)

- Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation)


.. toctree::
   :maxdepth: 2
   :numbered:

   installation
   quickstart
   sourcecode/index
   solutions
   revlog
   license


Support Content Creator
-----------------------
If you enjoyed this library, please consider supporting its creators! `Help Today <https://www.paypal.com/pools/c/8l2sqm1a6V>`_
