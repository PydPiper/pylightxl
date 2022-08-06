<p align="center">
    <img src="doc/source/_static/header_logo.gif" />
</p>

<p align="center">
    <a href="https://app.circleci.com/pipelines/github/PydPiper/pylightxl" alt="build">
        <img src="https://img.shields.io/circleci/build/gh/PydPiper/pylightxl" />
    </a>
    <a href="https://codecov.io/gh/PydPiper/pylightxl" alt="codecov">
        <img src="https://img.shields.io/codecov/c/github/PydPiper/pylightxl/master" />
    </a>
    <a href="https://pypi.org/project/pylightxl/" alt="pypi">
        <img src="https://img.shields.io/pypi/v/pylightxl" />
    </a>
    <a href="https://pypi.org/project/pylightxl/" alt="downloads">
        <img src="https://img.shields.io/pypi/dm/pylightxl" />
    </a>
    <a alt="python">
        <img src="https://img.shields.io/pypi/pyversions/pylightxl" />
    </a>
    <a alt="license">
        <img src="https://img.shields.io/github/license/PydPiper/pylightxl" />
    </a>
</p>

<h2 align="center">
    <p>
        pylightxl - A Light Weight Excel Reader/Writer
    </p>
    <a href="https://pylightxl.readthedocs.io">
        Documentation
    </a>
</h2>

<p align="center">
    <a>
        A light weight, zero dependency (only standard libs used), to the point (no bells and whistles) 
        Microsoft Excel reader/writer python 2.7.18 - 3+ library.
    </a>
    <img src="doc/source/_static/readme_demo.gif" />
</p>


**Please help us spread the word about pylightxl to the community by voting for pylightxl to be added
to python-awesome list. Follow the [LINK](https://github.com/vinta/awesome-python/pull/1449) and upvote
the pull request in the top right corner**

The project is associated with [Tidelift](https://tidelift.com/). Please consider supporting open-source
contributions by using the packages with Tidelift's subscription today!

Project featured on [PythonBytes Podcast Episode #165](https://pythonbytes.fm/episodes/show/165/ranges-as-dictionary-keys-oh-my)

---

#### **High-Level Feature Summary**

- **Reader**

    - supports Microsoft Excel 2004+ files (``.xlsx``, ``.xlsm``) and ``.csv`` files
    - read files via str path, pathlib path or file objects
    - read all or selective sheets
    - read type converted cell value (string, int, float), formula, comments, and named ranges

- **Database**

    - call cell value by row/col ID, excel address, or range
    - call an entire row/col or a semi-structured table based on user-defined headers

- **Writer**

    - write to new excel file (write excel files without having excel on your machine)
    - write to existing excel files (see limitations below)

#### **Limitations**

Although every effort was made to support a variety of users, the following limitations should be read carefully:

- Does not support ``.xls`` files (Microsoft Excel 2003 and older files)

- Writer does not support anything other than cell data (no graphs, images, macros, formatting)

- Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation), please use 64-bit if
  more data storage is required.

---

#### **Why pylightxl over pandas/openpyxl/xlrd**

- **Zero non-standard library dependencies!** 

    - No compatibility/version control issues.

- **Python2.7.18 to Python3+ support for life!** 

    - Don't worry about which python version you are using, pylightxl will support it for life

- **Light-weight single source code file** 

    - Want your project remain truly dependent-less? Copy the single source file into your project without any extra
      dependency issues or setup.
    - Do you struggle with other libraries weighing your projects down due to their very large size? Pylightxl's
      single source file size and zero dependency will not weigh your project down (preferable for django apps)
    - Do you struggle with ``pyinstaller`` or other ``exe`` wrappers failing to build or building to very large
      packages? Pylightxl will not cause any build errors and will not add to your build size since it has zero
      dependencies and a small lib size.
    - Do you struggle with download restrictions at your company? Copy the entire pylightxl source from 1 single file
      and use it in your project.

- **100% test-driven development for highest reliability/maintainability that aims for 100% coverage on all supported versions**

    - Pylightxl aims to test all of its features, however unforeseen edge cases can occur when receiving excel files
      created by non-microsoft excel. We actively monitor issues to add support for these edge cases should they arise.

- **API aimed to be user friendly and intuitive and well documented. Structure: database > worksheet > indexing**

    - ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1')``
    - ``db.ws('Sheet1').row(1)`` or ``db.ws('Sheet1').col(1)``

---

#### **Setup**
pylightxl is officially published on [pypi.org](https://pypi.org), however one of the
key features of pylightxl is that it is packed light in case the user has pip
and/or download restrictions, see [docs - installation](https://pylightxl.readthedocs.io/en/latest/installation.html)

```pip install pylightxl```

---

#### **pypi version 1.60**

pypi version 1.60
-----------------
- added feature: ability to update NamedRanges `wb.update_nr(name, val)`, see issue [#72](https://github.com/PydPiper/pylightxl/issues/72)
- added feature: ability to find where a NamedRange is `wb.nr_loc(name)`
- added feature: ability to fill a range with a single value: `wb.ws('Sheet1').update_range(address='A1:B3', val=10)`
- update: NamedRanges now add the worksheets if they are not already in the workbook. Note that using `readxl` with worksheet names specified will also ignore NamedRanges from being read in from the sheet that are not read in.
- update: updated quickstart docs with the new feature demo scripts


See full history log of revisions: [Here](https://pylightxl.readthedocs.io/en/latest/revlog.html)

---

#### **Contact/Questions/Suggestions**
If you have any questions or feedback, we would love to hear from you - send us 
a post directly on [GitHub](https://github.com/PydPiper/pylightxl/issues/new?assignees=&labels=&template=pylightxl-issue-template.md&title=).

We try to keep an active lookout for users trying to solve Microsoft Excel related problems with
python on Stack Overflow. Please help us build on the great community that python already is by
helping others get up to speed with pylightxl!

From everyone in the pylightxl family, thank you for visiting!

![logo](doc/source/_static/logo.png)
