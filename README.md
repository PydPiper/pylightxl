<p align="center">
    <img src=doc/source/_static/header_logo.gif>
</p>

<p align="center">

<a>

[![Build](https://img.shields.io/circleci/build/gh/PydPiper/pylightxl)](https://app.circleci.com/pipelines/github/PydPiper/pylightxl)

</a>

<a>

![Codecov branch](https://img.shields.io/codecov/c/github/PydPiper/pylightxl/master)

</a>

<a>

![PyPI](https://img.shields.io/pypi/v/pylightxl)

</a>

<a>

[![PyPI - Downloads](https://img.shields.io/pypi/dm/pylightxl)](https://pypi.org/project/pylightxl/)

</a>

<a>

![PyPI - Python Version](https://img.shields.io/pypi/pyversions/pylightxl)

</a>

<a>

![GitHub](https://img.shields.io/github/license/PydPiper/pylightxl)

</a>
    
</p>

<p align="center">

<a>

pylightxl - A Light Weight Excel Reader/Writer
A light weight, zero dependency (only standard libs used), to the point (no bells and whistles) 
Microsoft Excel reader/writer python 2.7-3+ library. See documentation: [docs](https://pylightxl.readthedocs.io)

</a>

<a>

![Example Code](doc/source/_static/readme_demo.gif)

</a>

<a>

[docs - quick start guide](https://pylightxl.readthedocs.io/en/latest/quickstart.html)

</a>
    
</p>

**Please help us spread the word about pylightxl to the community by voting for pylightxl to be added
to python-awesome list. Follow the [LINK](https://github.com/vinta/awesome-python/pull/1449) and upvote
the pull request in the top right corner (emoji)**

---

#### **Supports**:
 - Reader supports .xlsx and .xlsm file extensions.
 - Writer only supports .xlsx (no macros/buttons/graphs/formatting) 

#### **Limitations**:
 - Does not support .xls (excel 97-2003 worksheet).
 - Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation).
 - Writer does not support anything other than writing values/formulas/strings.
 - Writing to existing workbooks will remove any macros/buttons/graphs/formatting!

---

#### **Why pylightxl over pandas/openpyxl/xlrd**

- **Zero non-standard library dependencies** 
  No compatibility/version control issues.

- **Light-weight single source code file that supports both Python37 and Python27.** 
  Single source files that can easily be copied directly into a project for true zero-dependency. 
  Great for those that have installation/download restrictions. 
  In addition the library's size and zero dependency makes this library pyinstaller compilation small and easy!

- **100% test-driven development for highest reliability/maintainability with 100% coverage on all supported versions**

- **API aimed to be user friendly and intuitive. Structure: database > worksheet > indexing example:**
   ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1')``

---

#### **Setup**
pylightxl is officially published on [pypi.org](pypi.org), however one of the
key features of pylightxl is that it is packed light in case the user has pip
and/or download restrictions, see [docs - installation](https://pylightxl.readthedocs.io/en/latest/installation.html)

```pip install pylightxl```

---

#### **Future Tasks**
- additional database indexing features
- performance

#### **pypi version 1.44**
- bug fix: accounted for num2letter roll-over issue
- new feature: added a pylightxl native function for handling semi-structured data

See full history log of revisions: [Here](https://pylightxl.readthedocs.io/en/latest/revlog.html)

---

#### **Contact/Questions/Suggestions**
If you have any questions or feedback, we would love to hear from you - send us 
an realpydpiper@gmail.com or post directly on [GitHub](https://github.com/PydPiper/pylightxl).

We try to keep an active lookout for users trying to solve Microsoft Excel related problems with
python on Stack Overflow. Please help us build on the great community that python already is by
helping others get up to speed with pylightxl!

From everyone in the pylightxl family, thank you for visiting!

![logo](doc/source/_static/logo.png)
