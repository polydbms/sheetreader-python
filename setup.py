from setuptools import setup, Extension
import numpy
from numpy.distutils.misc_util import Configuration, get_info
from numpy.distutils.core import setup

sheetreader = Extension(
    "sheetreader",
    sources=[
        "src/interface.cpp",
        "src/sheetreader-core/src/XlsxSheet.cpp",
        "src/sheetreader-core/src/XlsxFile.cpp",
        "src/sheetreader-core/src/miniz/miniz.cpp",
    ],
    extra_compile_args=["/std:c++17", "/W4", "/DTARGET_PYTHON"])#, "/DEBUG:FULL"], extra_link_args=["/DEBUG:FULL"])

setup(name="SheetReader", ext_modules=[sheetreader], options={"build": {"build_lib": "."}})
