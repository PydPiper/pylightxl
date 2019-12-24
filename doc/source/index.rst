Welcome to pylightxl's documentation!
=====================================

.. |br| raw:: html

   <br />

Welcome to pylightxl documentation. A light weight Microsoft Excel File reader. Although there are
several excellent read/write options out there (`python-excel.org <python-excel.org>`_) pylightxl focused
on the following key features:

- Zero non-standard library dependencies (standard libs used: ``zipfile``, ``re``, ``os``). |br|
  No compatibility/version control issues.

- Single source code that supports both Pytohn37 and Python27.
  The light weight library is only 3 source files that can be easily copied directly into a project for
  those that have installation/download restrictions. In addition the library's size and zero dependency makes
  pyinstaller compilation small and easy!

- 100% test-driven development for highest reliability/maintainability with 100% coverage on all supported versions

- API aimed to be user friendly, intuitive and to the point with no bells and whistles. Structure: database > worksheet > indexing |br|
  example: ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1)``


High-Level Feature Summary
--------------------------
- Read excel files (``.xlsx``, ``.xlsm``), all sheets or selective few for speed/memory management

- Index cells data by row/col number or address

- Calling an entire row/col of data returns an easy to use list output: |br| ``db.ws('Sheet1').row(1)`` or ``db.ws('Sheet1').rows``

- Worksheet data size is consistent for each row/col. Any data that is empty will return a ''

- Writer coming soon!


Limitations
-----------
Although every effort was made to support a variety of users, the following limitations should be read carefully:

- Does not support ``.xls`` files (Microsoft Excel 2003 and older files)

- Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation)


.. toctree::
   :maxdepth: 2
   :numbered:

   installation
   quickstart
   sourcecode/index
   license


Support Content Creator
-----------------------
If you enjoyed this library, please consider supporting its creators! `Help Today <https://www.paypal.com/pools/c/8l2sqm1a6V>`_
