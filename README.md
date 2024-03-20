# Python Bindings for SheetReader

SheetReader allows you to read your spreadsheets (.xlsx) files blazingly fast. This repository contains the Python bindings, as the core library is implemented in C++.

SheetReader is part of the [PolyDB Project](https://polydbms.org/). Feel free to also checkout our [SheetReader core parser](https://github.com/polydbms/sheetreader-core) and our [SheetReader R bindings](https://github.com/fhenz/SheetReader-r)

## Build Instructions
To build the Python bindings (while we prepare `pip install sheetreader`), you should start by cloning this repo, including its submodules:
```
git clone --recurse-submodules https://github.com/polydbms/sheetreader-python.git
```

The bindings require numpy as a dependency. If you do not have it installed, install it, e.g. through pip:
```
pip3 install numpy
```

Finally, build the bindings
```
python3 setup.py build
```
which will create a shared library that you can then copy it into your desired directory, e.g.:

```
cp sheetreader.cpython-38-x86_64-linux-gnu.so /usr/local/lib/python3.8/dist-packages
```

## Usage
Now you can parse your favorite spreadsheets by:
```
import sheetreader
sheet = sheetreader.read_xlsx("my_favorite_sheet.xlsx")
```

