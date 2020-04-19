Quick Start Guide
=================
Get up and running in less than 5 minutes with pylightxl!

.. figure:: _static/readme_demo.gif
   :align: center



Read Excel File
---------------

.. code-block:: python

    import pylightxl as xl

    # readxl returns a pylightxl database that holds all worksheets and its data
    db = xl.readxl('folder1/folder2/excelfile.xlsx')

    # read only selective sheetnames
    db = xl.readxl('folder1/folder2/excelfile.xlsx', ('Sheet1','Sheet3'))

    # return all sheetnames
    db.ws_names
    >>> ['Sheet1', 'Sheet3']

Access Worksheet and Cell Data
------------------------------
The following example assumes ``excelfile.xlsx`` contains a worksheet named ``Sheet1`` and it has the
following cell content:

+----+----+----+----+
|    | A  | B  | C  |
+----+----+----+----+
| 1  | 10 | 20 |    |
+----+----+----+----+
| 2  |    | 30 | 40 |
+----+----+----+----+

.. code-block:: python

    import pylightxl as xl

    db = xl.readxl('excelfile.xlsx')

- access by worksheet name (tab name) and cell address

.. code-block:: python

    db.ws('Sheet1').address('A1')
    >>> 10

- access by worksheet name (tab name) and cell index

.. code-block:: python

    db.ws('Sheet1').index(row=1,col=2)
    >>> 20

- access an entire row/col (note: empty cells are returned as '')

.. code-block:: python

    db.ws('Sheet1').row(1)
    >>> [10,20,'']

    db.ws('Sheet1').col(1)
    >>> [10,'']

- get an entire row/col based on key-value (note: key is type sensitive)

.. code-block:: python

    # lets say we would like to return the column that has a cell value = 20 in row1
    db.ws('Sheet1').keycol(key=20)
    >>> [20,30]

    # we can also specify a custom keyindex (not just row1), note that we now are matched based on row2
    db.ws('Sheet1').keycol(key=30, keyindex=2)
    >>> [20,30]

    # similarly done for keyrow
    db.ws('Sheet1').keyrow(key='')
    >>> ['',30,40]

- get the size of a worksheet

.. code-block:: python

    db.ws('Sheet1').size
    >>> [2,3]


- iterate through rows/cols

.. code-block:: python

    for row in db.ws('Sheet1').rows:
        print(row)

    >>> [10,20,'']
    >>> ['',30,40]

    for col in db.ws('Sheet1').cols:
        print(col)

    >>> [10,'']
    >>> [20,30]
    >>> ['',40]

Write out a pylightxl.Database as an excel file
-----------------------------------------------
Pylightxl support excel writing without having excel installed on the machine. However it is not without
its limitations. The writer only supports cell data writing (ie.: does not support graphs, formatting, images,
macros, etc) simply just strings/numbers/equations in cells.

Note that equations typed by the user will not calculate for its value until the excel sheet is opened in excel.

.. code-block:: python

   import pylightxl as xl

   # read in an existing worksheet and change values of its cells (same worksheet as above)
   db = xl.readxl('excelfile.xlsx')
   # overwrite existing number value
   db.ws('Sheet1').index(row=1, col=1)
   >>> 10
   db.ws('Sheet1').update_index(row=1, col=1, val=100)
   db.ws('Sheet1').index(row=1, col=1)
   >>> 100
   # write text
   db.ws('Sheet1').update_index(row=1, col=2, val='twenty')
   # write equations
   db.ws('Sheet1').update_address(address='A3', val='=A1')

   xl.writexl(db, 'updated.xlsx')



