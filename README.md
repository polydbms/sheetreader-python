# Python Bindings for SheetReader

SheetReader allows you to read your spreadsheets (.xlsx) files blazingly fast. This repository contains the Python bindings, as the core library is implemented in C++.

SheetReader is part of the [PolyDB Project](https://polydbms.org/). Feel free to also checkout our [SheetReader core parser](https://github.com/polydbms/sheetreader-core) and our [SheetReader R bindings](https://github.com/fhenz/SheetReader-r)

## Build Instructions
Contains `pyproject.toml`, so buildable (after `git clone --recurse-submodules https://github.com/polydbms/sheetreader-python.git`) via e.g. `python -m build .` to generate the wheel file.  
Also installable via pip (`pip install .`).

## Usage
Now you can parse your favorite spreadsheets by:
```
import sheetreader
sheet = sheetreader.read_xlsx("my_favorite_sheet.xlsx")
```

### Parameters:
**path** *string* The path of the `xlsx` file to parse.  
**sheet** *integer or string* Which sheet of the file to parse, can be either the index (starting at 1) or the name. default *1*  
**headers** *boolean* Whether to interpret the first parsed row as headers. default *True*  
**skip_rows** *integer* How many rows to skip before parsing data. default *0*  
**skip_columns** *integer* How many columns to skip before parsing data. default *0*  
**num_threads** *integer* How many threads to use for parsing. default *-1(auto)*  
**col_types** *dict or list* How to interpret parsed data, either by names (dict) or by position (list). Types must be of `numeric`,`text`,`logical`,`date`,`skip`,`guess`. default *None*  
