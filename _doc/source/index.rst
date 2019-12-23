Welcome to pylightxl's documentation!
=====================================
Welcome to pylightxl documentation. A light weight Microsoft Excel File reader. Although there are
several excellent read/write options out there (`python-excel.org <python-excel.org>`_) pylightxl focused
on the following key features:

- Zero non-standard library dependencies (standard libs used: ``zipfile``, ``re``, ``os``)

- Supported on Pytohn37 and Python27!

- 100% test-driven development for highest reliability/maintainability with 100% coverage on all supported versions

- Read excel files (``.xlsx``, ``.xlsm``), all sheets or selective few

- Easy to use and intuitive worksheet by worksheet data handling
  ``db.ws('Sheet1').index(row=,col=)``  or ``db.ws('Sheet1').address(address=)``

- Iterate through row/col data: ``db.ws('Sheet1').row(1)`` or ``db.ws('Sheet1').rows``

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
