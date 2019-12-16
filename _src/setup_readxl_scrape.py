from distutils.core import setup
from Cython.Build import cythonize
from
from Cython.Distutils import build_ext

ext_modules = [Extension]

setup(name='readxl_scrape', ext_modules=cythonize('readxl_scrape.pyx'))
