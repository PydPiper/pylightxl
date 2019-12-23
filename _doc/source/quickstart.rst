Quick Start Guide
=================
Get up and running in less than 5 minutes with pylightxl!


Read Excel File
---------------

.. code-block:: python

    import pylightxl as xl

    # readxl returns a pylightxl database that holds all worksheets and its data
    db = xl.readlxl('folder1/folder2/excelfile.xlsx')

    # read only selective sheetnames
    db = xl.readlxl('folder1/folder2/excelfile.xlsx', ('Sheet1','Sheet3'))

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

