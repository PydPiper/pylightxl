![logo](doc/source/_static/header_logo.gif)
# pylightxl - A Light Weight Excel Reader/Writer
A light weight, zero dependency (only standard libs used), to the point (no bells and whistles) 
Microsoft Excel reader/writer(coming soon) python 2.7-3+ library. See documentation: [docs](https://pylightxl.readthedocs.io)

Sample - see [docs - quick start gudide](https://pylightxl.readthedocs.io/en/latest/quickstart.html) for more examples:
![Example Code](doc/source/_static/readme_demo.gif)

---

#### **Supports**:
 - Supports .xlsx and .xlsm file extensions. 

#### **Limitations**:
 - Does not support .xls (excel 97-2003 worksheet).
 - Does not support worksheet cell data more than 536,870,912 cells (32-bit list limitation)
 
#### **In-work version 1.4**
- new-feature: write new excel file from pylightxl.Database
- new-feature: write to existing excel file from pylightxl.Database

#### **Why pylightxl over pandas/openpyxl**
- (compatability +1, small lib +1) pylightxl has no external dependencies (only uses python builtin standard libs)
- (compatability +1) pylightxl was written to be compatible for python 2.7-3+, it does not impose rules on users to switch versions
- (small lib +1) pylightxl was written to simply read/write, thereby making the library small without any bells or whiles which makes
  it easy to compile with pyinstaller and other packagers
- (user friendly +1) pylightxl was written to be as pythonic and easy to use as possible. Core developers actively survey stackoverflow 
  questions on working with excel files to tailor the API for most common problems.

---

#### **Setup**
pylightxl is officially published on [pypi.org](pypi.org), however one of the
key features of pylightxl is that it is packed light in case the user has pip
and/or download restrictions, see [docs - installation](https://pylightxl.readthedocs.io/en/latest/installation.html)

```pip install pylightxl```

---

#### **Contact/Questions/Suggestions**
If you have any questions or feedback, we would love to hear from you - send us 
an [email](pylightxl@gmail.com) or post directly on [GitHub](https://github.com/PydPiper/pylightxl).

We try to keep an active lookout or users trying to solve Microsoft Excel related problems with
python on StackoverFlow. Please help us build on the great community that python already is by
helping others get up to speed with pylightxl!

From everyone in the pylightxl family, thank you for visiting!

![logo](doc/source/_static/logo.png)