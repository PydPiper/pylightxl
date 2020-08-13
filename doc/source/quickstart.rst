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

Read Semi-Structured Data
-------------------------
.. figure:: _static/ex_readsemistrdata.png

- note that ``ssd`` function takes any key-word argument as your KEYROWS/KEYCOLS flag
- multiple tables are read the same way as you would read a book. Top left-to-right, then down

.. code-block:: python

    import pylightxl
    db = pylightxl.readxl('Book1.xlsx')

    # request a semi-structured data (ssd) output
    ssd = db.ws('Sheet1').ssd(keycols="KEYCOLS", keyrows="KEYROWS")

    ssd[0]
    >>> {'keyrows': ['r1', 'r2', 'r3'], 'keycols': ['c1', 'c2', 'c3'], 'data': [[1, 2, 3], [4, '', 6], [7, 8, 9]]}
    ssd[1]
    >>> {'keyrows': ['rr1', 'rr2', 'rr3', 'rr4'], 'keycols': ['cc1', 'cc2', 'cc3'], 'data': [[10, 20, 30], [40, 50, 60], [70, 80, 90], [100, 110, 120]]}



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


Write a new excel file from python data
---------------------------------------
For new python data that did not come from an existing excel speadsheet.

.. code-block:: python

    import pylightxl as xl

    # take this list for example as our input data that we want to put in column A
    mydata = [10,20,30,40]

    # create a black db
    db = xl.Database()

    # add a blank worksheet to the db
    db.add_ws(sheetname="Sheet1", data={})

    # loop to add our data to the worksheet
    for row_id, data in enumerate(mydata, start=1)
        db.ws("Sheet1").update_index(row=row_id, col=1, val=data)

    # write out the db
    xl.writexl(db, "output.xlsx")

