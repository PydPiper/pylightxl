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
    <a href="https://pylightxl.readthedocs.io/en/latest/quickstart.html">
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

Project featured on [PythonBytes Podcast Episode #165](https://pythonbytes.fm/episodes/show/165/ranges-as-dictionary-keys-oh-my)

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

- **Light-weight single source code file that supports both Python3 and Python2.7.18.** 
  Single source file that can easily be copied directly into a project for true zero-dependency. 
  Great for those that have installation/download restrictions. 
  In addition the library's size and zero dependency makes this library pyinstaller compilation small and easy!

- **100% test-driven development for highest reliability/maintainability with 100% coverage on all supported versions**

- **API aimed to be user friendly and intuitive. Structure: database > worksheet > indexing example:**
   ``db.ws('Sheet1').index(row=1,col=2)``  or ``db.ws('Sheet1').address(address='B1')``

---

#### **Setup**
pylightxl is officially published on [pypi.org](https://pypi.org), however one of the
key features of pylightxl is that it is packed light in case the user has pip
and/or download restrictions, see [docs - installation](https://pylightxl.readthedocs.io/en/latest/installation.html)

```pip install pylightxl```

---

#### **pypi version 1.53**

- bug fix: writing to existing file previously would only write to the current working directory, it 
  now can handle subdirs. In addition inadvertenally discovered a bug in python source code ElementTree.iterparse
  where ``source`` passed as a string was not closing the file properly. We submitted a issue to python issue tracker.
  
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
