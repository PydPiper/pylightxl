Revision Log
============

pypi version 1.45 (in-work)
------------
- added support for cell values that have multiple formats within a single cell.
  previous versions did not support this functionality since it is logged differently in sharedString.xml
- added support for updating formulas and viewing them:

    - view formula: ``db.ws('Sheet1').address('A1', formula=True)``
    - edit formula: ``db.ws('Sheet1').update_address('A1', val='=A1+10')``

- updated the following function arguments to drive commonality:

    - was: ``readxl(fn, sheetnames)`` new: ``readxl(fn, ws)``
    - was: ``writexl(db, path)`` new: ``writexl(db, fn)``
    - was: ``db.ws(sheetname)`` new: ``db.ws(ws)``
    - was: ``db.add_ws(sheetname, data)`` new: ``db.add_ws(ws, data)``

pypi version 1.44
-----------------
- bug fix: accounted for num2letter roll-over issue
- new feature: added a pylightxl native function for handling semi-structured data

pypi version 1.43
-----------------
- bug fix: accounted for reading error'ed out cell "#N/A"
- bug fix: accounted for bool TRUE/FALSE cell values not registering on readxl
- bug fix: accounted for edge case that was prematurely splitting cell tags <c r /> by formula closing
  bracket <f />
- bug fix: accounted for cell address roll-over

pypi version 1.42
-----------------
- added support for pathlib file reading
- bug fix: previous version did not handle merged cells properly
- bug fix: database updates did not update maxcol maxrow if new data addition was larger than the initial
  dataset
- bug fix: writexl that use linefeeds did not read in properly into readxl (fixed regex)
- bug fix: writexl filepath issues

pypi version 1.41
-------------------
- new-feature: write new excel file from pylightxl.Database
- new-feature: write to existing excel file from pylightxl.Database
- new-feature: db.update_index(row, col, val) for user defined cell values
- new-feature: db.update_address(address, val) for user defined cell values
- bug fix for reading user defined sheets
- bug fix for mis-alignment of reading user defined sheets and xml files

pypi version 1.3
----------------
- new-feature: add the ability to call rows/cols via key-value ex: ``db.ws('Sheet1').keycol('my column header')``
  will return the entire column that has 'my column header' in row 1

- fixed-bug: fixed leading/trailing spaced cell text values that are marked ``<t xml:space="preserve">`` in the
  sharedString.xml

pypi version 1.2
----------------
- fixed-bug: fixed Sheet number to custom Sheet name matching for 10+ sheets that were previously only sorting alphabetical
  which resulted with sorting: Sheet1, Sheet10, Sheet11, Sheet2... and so on.

pypi version 1.1
----------------
- initial release
