# SheetReader Python Bindings

SheetReader allows to read your Excel spreadsheet files (.xlsx) blazingly fast. This repository contains the Python bindings, as the [core library](https://github.com/polydbms/sheetreader-core) is implemented in C++.

## Quickstart
Sheetreader is available through:
```shell
pip install pysheetreader
```
After successful installation, spreadsheets can be loaded:
```python
import pysheetreader as sr
sheet = sr.read_xlsx("my_favorite_sheet.xlsx")
```
To convert a spreadsheet into a `pandas` Dataframe:
```python
import pysheetreader as sr
import pandas as pd
sheet = sr.read_xlsx("my_favorite_sheet.xlsx")
df = pd.DataFrame.from_dict(sheet[0])
```

### Parameters:
| Parameter     | Type                | Description                                                                                               | Default  |
|---------------|---------------------|-----------------------------------------------------------------------------------------------------------|----------|
| `path`        | `string`            | The path of the `.xlsx` file to parse.                                                                     | -        |
| `sheet`       | `integer` or `string`| The sheet of the file to parse, can be either the index (starting at 1) or the name.                        | `1`      |
| `headers`     | `boolean`           | Whether to interpret the first parsed row as headers.                                                      | `True`   |
| `skip_rows`   | `integer`           | How many rows to skip before parsing data.                                                                 | `0`      |
| `skip_columns`| `integer`           | How many columns to skip before parsing data.                                                              | `0`      |
| `num_threads` | `integer`           | How many threads to use for parsing. Use `-1` for automatic threading.                                     | `-1`     |
| `col_types`   | `dict` or `list`    | How to interpret parsed data, either by names (`dict`) or by position (`list`). Types: `numeric`, `text`, `logical`, `date`, `skip`, `guess`. | `None`   |
  


## Build Instructions
First install the submodules, which contain the `sheetreader-core` dependency with: 
```shell
git clone --recurse-submodules https://github.com/polydbms/sheetreader-python.git
```
To build from source, this repository provides a `pyproject.toml`.
The SheetReader wheel file can be generated through:
```shell
python -m build .
```
or installed with pip through:
```shell
pip install .
```

## More resources
SheetReader is part of the [PolyDB Project](https://polydbms.org/). We also provide bindings/extensions for several other environments:
- [R language](https://github.com/fhenz/SheetReader-r/): Load spreadsheets into dataframes, also available via [CRAN](https://cran.r-project.org/package=SheetReader).
- [PostgreSQL FDW](https://github.com/polydbms/pg_sheet_fdw): Foreign data wrapper for PostgreSQL; allows to register spreadsheets as foreign tables. 
- [DuckDB Extension](https://github.com/polydbms/sheetreader-duckdb): Extension for DuckDB that allows loading spreadsheets into tables. Also available as a [community extension](https://community-extensions.duckdb.org/extensions/sheetreader.html).

## Paper
SheetReader was published in the [Information Systems Journal](https://www.sciencedirect.com/science/article/abs/pii/S0306437923000194). Cite as:
```
@article{DBLP:journals/is/GavriilidisHZM23,
  author       = {Haralampos Gavriilidis and
                  Felix Henze and
                  Eleni Tzirita Zacharatou and
                  Volker Markl},
  title        = {SheetReader: Efficient Specialized Spreadsheet Parsing},
  journal      = {Inf. Syst.},
  volume       = {115},
  pages        = {102183},
  year         = {2023},
  url          = {https://doi.org/10.1016/j.is.2023.102183},
  doi          = {10.1016/J.IS.2023.102183},
  timestamp    = {Mon, 26 Jun 2023 20:54:32 +0200},
  biburl       = {https://dblp.org/rec/journals/is/GavriilidisHZM23.bib},
  bibsource    = {dblp computer science bibliography, https://dblp.org}
}