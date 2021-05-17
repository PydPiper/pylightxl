Revision Log
============

pypi version 1.55 (in-work)
---------------------------
- added comment parsing, see issue `#41 <https://github.com/PydPiper/pylightxl/issues/41>`_
- DEPRECATION WARNING: all indexing method that use "formula" as an argument will be replaced
  with "output" in future version. Please update your codebase to use "output" instead of "formula".
  This was done to simplify indexing the value (``output='v'``), the formula (``output='f'``) or the
  comment (``output='c'``).
- added file stream reading for ``readxl`` that now supports ``with block`` for reading. See issue `#25 <https://github.com/PydPiper/pylightxl/issues/25>`_

pypi version 1.54
-----------------
- added handling for datetime parsing

pypi version 1.53
-----------------
- bug fix: writing to existing file previously would only write to the current working directory, it
  now can handle subdirs. In addition inadvertently discovered a bug in python source code ElementTree.iterparse
  where ``source`` passed as a string was not closing the file properly. We submitted a issue to python issue tracker.

pypi version 1.52
-----------------
- updated reading error'ed cells "#N/A"
- updated workbook indexing bug from program generated workbooks that did not index from 1

pypi version 1.51
---------------------------
- license update within setup.py

pypi version 1.50
-----------------
- hot-fix: added python2 support for encoding with cgi instead of html

pypi version 1.49
-----------------
- bug-fix: updated encoding for string cells that contained xml-like data (ex: cell A1 "<cell content>")

pypi version 1.48
-----------------
- add feature to ``writecsv`` to be able to handle ``pathlib`` object and ``io.StreamIO`` object
- refactored readxl to remove regex, now readxl is all cElementTree
- refactored readxl/writexl to able to handle excel files written by openpyxl that is generated
  differently than how excel write files.

pypi version 1.47
-----------------
- added new function: ``db.nr('table1')`` returns the contents of named range "table1"
- added new function: ``db.ws('Sheet1').range('A1:C3')`` that returns the contents of a range
  it also has the ability to return the formulas of the range
- updated ``db.ws('Sheet1').row()`` and ``db.ws('Sheet1').col()`` to take in a new argument ``formual``
  that returns the formulas of a row or col
- bugfix: write to existing without named ranges was throwing a "repair" error. Fixed typo on xml for it
  and added unit tests to capture it
- added new function: ``xl.readcsv(fn, delimiter, ws)`` to read csv files and create a pylightxl db out
  of it (type converted)
- added new function: ``xl.writecsv(db, fn, ws, delimiter)`` to write out a pylightxl worksheet as a csv


pypi version 1.46
------------------
- bug fix: added ability to input an empty string into the cell update functions
  (previously entering val='') threw and error

pypi version 1.45
-----------------
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

- added new feature to be able to read-in NamedRanges, store it in the Database, update it, remove it,
  and write it. NamedRanges were integrated with existing function to handle semi-structured-data

    - ``db.add_nr(name'range1', ws='sheet1', address='A1:C2')``
    - ``db.remove_nr(name='range1')``
    - ``db.nr_names``

- add feature to remove worksheet: ``db.remove_ws(ws='Sheet1')``
- add feature to rename worksheet: ``db.rename_ws(old='sh1', new='sh2')``
- added a cleanup function upon writing to delete _pylightxl_ temp folder in case an error left them
- added feature to write to file that is open by excel by appending a "new_" tag to the file name and
  a warning message that file is opened by excel so a file was saved as "new_" + filename

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
