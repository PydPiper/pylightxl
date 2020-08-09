Example Solutions
=================

Reading Semi Structured data
----------------------------
- Question posted on `stackoverflow <https://stackoverflow.com/questions/59533824/python-extract-data-from-a-semi-structured-xlsx-file/59534919#59534919>`_

- Problem: read groups of 2D data from a single sheet that can begin at any row/col and has any
  number of rows/columns per data group, see figure below.

.. figure:: _static/ex_readsemistrdata.png

- Solution: note that ``ssd`` function takes any key-word argument as your KEYROWS/KEYCOLS flag and
  multiple tables are read the same way as you would read a book. Top left-to-right, then down.

.. code-block:: python

    import pylightxl
    db = pylightxl.readxl('Book1.xlsx')

    # request a semi-structured data (ssd) output
    ssd = db.ws('Sheet1').ssd(keycols="KEYCOLS", keyrows="KEYROWS")

    ssd[0]
    >>> {'keyrows': ['r1', 'r2', 'r3'], 'keycols': ['c1', 'c2', 'c3'], 'data': [[1, 2, 3], [4, '', 6], [7, 8, 9]]}
    ssd[1]
    >>> {'keyrows': ['rr1', 'rr2', 'rr3', 'rr4'], 'keycols': ['cc1', 'cc2', 'cc3'], 'data': [[10, 20, 30], [40, 50, 60], [70, 80, 90], [100, 110, 120]]}
