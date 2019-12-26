Revision Log
============

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
