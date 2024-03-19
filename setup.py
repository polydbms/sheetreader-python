from setuptools import setup, Extension
import numpy
from numpy.distutils.misc_util import Configuration, get_info
from numpy.distutils.core import setup
import platform

extra_compile_args = []
extra_link_args = []

if platform.system() == 'Windows':
    extra_compile_args = ["/std:c++17", "/W4", "/DTARGET_PYTHON"]
    extra_link_args = []
elif platform.system() == 'Linux':
    extra_compile_args = ["-std=c++17", "-Wall", "-DTARGET_PYTHON"]
    extra_link_args = []

sheetreader = Extension(
    "sheetreader",
    sources=[
        "src/interface.cpp",
        "src/sheetreader-core/src/XlsxSheet.cpp",
        "src/sheetreader-core/src/XlsxFile.cpp",
        "src/sheetreader-core/src/miniz/miniz.cpp",
    ],
    extra_compile_args=extra_compile_args,
    extra_link_args=extra_link_args
)

setup(name="SheetReader", ext_modules=[sheetreader], options={"build": {"build_lib": "."}})

