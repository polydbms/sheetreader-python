#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include <Python.h>
#include "numpy/ndarraytypes.h"
#include "numpy/npy_3kcompat.h"

#include <thread>
#include <iostream>
#include <cmath>

#include "sheetreader-core/src/XlsxFile.h"
#include "sheetreader-core/src/XlsxSheet.h"

PyObject* cells_to_python(const XlsxFile& file, XlsxSheet& sheet) {
	PyObject* tuple = PyTuple_New(1);

	size_t nColumns = sheet.mDimension.first;
	size_t nRows = sheet.mDimension.second;
	nRows = nRows - sheet.mHeaders - sheet.mSkipRows;
	nColumns = nColumns - sheet.mSkipColumns;
	if (nRows == 0) {
		return tuple;
	}

	std::vector<PyArrayObject*> proxies;
	std::vector<std::tuple<XlsxCell, CellType, size_t>> headerCells;
	if (nColumns > 0) {
		proxies.reserve(nColumns);
		headerCells.reserve(nColumns);
	}
	std::vector<CellType> coltypes(nColumns, CellType::T_NONE);
	std::vector<CellType> coerce(nColumns, CellType::T_NONE);

	const npy_intp d1[] = {nRows};
	const double nanv = nan("");

	unsigned long currentColumn = 0;
	long long currentRow = -1;
	std::vector<size_t> currentLocs(sheet.mCells.size(), 0);
	const size_t maxBuffers = sheet.mCells.size() > 0 ? sheet.mCells[0].size() : 0;
	for (size_t buf = 0; buf < maxBuffers; ++buf) {
		for (size_t ithread = 0; ithread < sheet.mCells.size(); ++ithread) {
			if (sheet.mCells[ithread].size() == 0) {
				break;
			}
			std::cout << buf << ", " << ithread << "/" << sheet.mCells.size() << std::endl;
			const std::vector<XlsxCell> cells = sheet.mCells[ithread].front();
			const std::vector<LocationInfo>& locs = sheet.mLocationInfos[ithread];
			size_t& currentLoc = currentLocs[ithread];

			// icell <= cells.size() because there might be location info after last cell
			for (size_t icell = 0; icell <= cells.size(); ++icell) {
				while (currentLoc < locs.size() && locs[currentLoc].buffer == buf && locs[currentLoc].cell == icell) {
					std::cout << "loc " << currentLoc << "/" << locs.size() << ": " << locs[currentLoc].buffer << " vs " << buf << ", " << locs[currentLoc].cell << " vs " << icell << " (" << locs[currentLoc].column << "/" << locs[currentLoc].row << ")" << std::endl;
					currentColumn = locs[currentLoc].column;
					if (locs[currentLoc].row == -1ul) {
						++currentRow;
					} else {
						currentRow = locs[currentLoc].row;
					}
					++currentLoc;
				}
				if (icell >= cells.size()) break;
				const auto adjustedColumn = currentColumn;
				const auto adjustedRow = currentRow - sheet.mSkipRows;
				const XlsxCell& cell = cells[icell];
				const CellType type = cell.type;

				//std::cout << adjustedColumn << "/" << adjustedRow << std::endl;

				if (adjustedRow >= sheet.mHeaders) {
					// normal (non-header) cell
					if (coltypes.size() <= adjustedColumn) {
						coltypes.resize(adjustedColumn + 1, CellType::T_NONE);
						coerce.resize(adjustedColumn + 1, CellType::T_NONE);
					}
					if (coltypes[adjustedColumn] == CellType::T_NONE || coltypes[adjustedColumn] == CellType::T_ERROR) {
						PyArrayObject* robj = reinterpret_cast<PyArrayObject*>(PyArray_SimpleNew(1, d1, NPY_DOUBLE));
						for (unsigned long i = 0; i < nRows; ++i) *reinterpret_cast<double*>(PyArray_GETPTR1(robj, i)) = nanv;
						if (type == CellType::T_NUMERIC) {
							//NOOP
						} else if (type == CellType::T_STRING_REF || type == CellType::T_STRING || type == CellType::T_STRING_INLINE) {
							robj = reinterpret_cast<PyArrayObject*>(PyArray_SimpleNew(1, d1, NPY_OBJECT));
						} else if (type == CellType::T_BOOLEAN) {
							robj = reinterpret_cast<PyArrayObject*>(PyArray_SimpleNew(1, d1, NPY_BOOL));
						} else if (type == CellType::T_DATE) {
							PyObject *date_type = Py_BuildValue("s", "M8[ms]");
							PyArray_Descr *descr;
							PyArray_DescrConverter(date_type, &descr);
							Py_XDECREF(date_type);

							robj = reinterpret_cast<PyArrayObject*>(PyArray_SimpleNewFromDescr(1, d1, descr));
							for (unsigned long i = 0; i < nRows; ++i) *reinterpret_cast<long long*>(PyArray_GETPTR1(robj, i)) = NPY_DATETIME_NAT;
						}
						if (type != CellType::T_NONE && robj != nullptr) {
							if (proxies.size() < adjustedColumn) {
								proxies.reserve(adjustedColumn + 1);
								proxies.resize(adjustedColumn);
								proxies.push_back(robj);
							} else if (proxies.size() > adjustedColumn) {
								proxies[adjustedColumn] = robj;
							} else {
								proxies.push_back(robj);
							}
							coltypes[adjustedColumn] = type;
						}
					}
					if (coltypes[adjustedColumn] != CellType::T_NONE && type != CellType::T_NONE && type != CellType::T_ERROR) {
						const CellType col_type = coltypes[adjustedColumn];
						const bool compatible = ((type == col_type)
							|| (type == CellType::T_STRING_REF && col_type == CellType::T_STRING)
							|| (type == CellType::T_STRING_REF && col_type == CellType::T_STRING_INLINE)
							|| (type == CellType::T_STRING && col_type == CellType::T_STRING_REF)
							|| (type == CellType::T_STRING && col_type == CellType::T_STRING_INLINE)
							|| (type == CellType::T_STRING_INLINE && col_type == CellType::T_STRING_REF)
							|| (type == CellType::T_STRING_INLINE && col_type == CellType::T_STRING));
						const unsigned long i = adjustedRow - sheet.mHeaders;
						if (coerce[adjustedColumn] == CellType::T_STRING) {
							PyArrayObject* robj = proxies[adjustedColumn];
							//coerceString(file, ithread, robj, i, cell, type);
						} else if (compatible) {
							PyArrayObject* robj = proxies[adjustedColumn];
							if (type == CellType::T_NUMERIC) {
								*reinterpret_cast<double*>(PyArray_GETPTR1(robj, i)) = cell.data.real;
							} else if (type == CellType::T_STRING_REF) {
								auto* str = file.getString(cell.data.integer);
								*reinterpret_cast<PyObject**>(PyArray_GETPTR1(robj, i)) = str;
							} else if (type == CellType::T_STRING || type == CellType::T_STRING_INLINE) {
								auto& str = file.getDynamicString(ithread, cell.data.integer);
								*reinterpret_cast<PyObject**>(PyArray_GETPTR1(robj, i)) = PyUnicode_FromString(str.c_str());
							} else if (type == CellType::T_BOOLEAN) {
								*reinterpret_cast<bool*>(PyArray_GETPTR1(robj, i)) = cell.data.boolean;
							} else if (type == CellType::T_DATE) {
								*reinterpret_cast<long long*>(PyArray_GETPTR1(robj, i)) = cell.data.real * 1000;
							}
						} else if (coerce[adjustedColumn] == CellType::T_NONE) {
							/*coerce[adjustedColumn] = CellType::T_STRING;
							if (col_type != CellType::T_STRING && col_type != CellType::T_STRING_REF && col_type != CellType::T_STRING_INLINE) {
								// convert existing
								Rcpp::RObject& robj = proxies[adjustedColumn];
								Rcpp::RObject newObj = Rcpp::CharacterVector(nRows, Rcpp::CharacterVector::get_na());
								for (size_t i = 0; i < nRows; ++i) {
									if (col_type == CellType::T_NUMERIC) {
										if (Rcpp::NumericVector::is_na(static_cast<Rcpp::NumericVector>(robj)[i])) continue;
										static_cast<Rcpp::CharacterVector>(newObj)[i] = formatNumber(static_cast<Rcpp::NumericVector>(robj)[i]);
									} else if (col_type == CellType::T_BOOLEAN) {
										if (Rcpp::LogicalVector::is_na(static_cast<Rcpp::LogicalVector>(robj)[i])) continue;
										static_cast<Rcpp::CharacterVector>(newObj)[i] = static_cast<Rcpp::LogicalVector>(robj)[i] ? "TRUE" : "FALSE";
									} else if (col_type == CellType::T_DATE) {
										if (Rcpp::DatetimeVector::is_na(static_cast<Rcpp::DatetimeVector>(robj)[i])) continue;
										static_cast<Rcpp::CharacterVector>(newObj)[i] = formatDatetime(static_cast<Rcpp::DatetimeVector>(robj)[i]);
									}
								}
								proxies[adjustedColumn] = newObj;
							}
							coerceString(file, ithread, proxies[adjustedColumn], i, cell, type);*/
						}
					}
				} else {
					// header cell
					if (headerCells.size() <= adjustedColumn) {
						headerCells.resize(adjustedColumn + 1);
					}
					headerCells[adjustedColumn] = std::make_tuple(cell, type, ithread);
				}
				++currentColumn;
			}
			sheet.mCells[ithread].pop_front();
		}
	}
	size_t numCols = std::max(proxies.size(), headerCells.size());
	PyObject* dict = PyDict_New();

	std::cout << "Transform " << proxies.size() << ", " << headerCells.size() << std::endl;
	for (size_t i = 0; i < numCols; ++i) {
		// header
		std::string colName = "Column" + std::to_string(i);
		if (i < headerCells.size() && std::get<1>(headerCells[i]) != CellType::T_NONE) {
			auto& cell = std::get<0>(headerCells[i]);
			auto& type = std::get<1>(headerCells[i]);
			if (type == CellType::T_NUMERIC) {
				colName = cell.data.real;
			} else if (type == CellType::T_STRING_REF) {
				colName = PyUnicode_AsUTF8(file.getString(cell.data.integer));
			} else if (type == CellType::T_STRING || type == CellType::T_STRING_INLINE) {
				colName = file.getDynamicString(std::get<2>(headerCells[i]), cell.data.integer);
			} else if (type == CellType::T_BOOLEAN) {
				colName = cell.data.boolean;
			} else if (type == CellType::T_DATE) {
				colName = cell.data.real;
			}
		}
		// data
		std::cout << colName << std::endl;
		if (i < proxies.size()) {
			PyDict_SetItemString(dict, colName.c_str(), reinterpret_cast<PyObject*>(proxies[i]));
		} else {
			PyArrayObject* robj = reinterpret_cast<PyArrayObject*>(PyArray_SimpleNew(1, d1, NPY_DOUBLE));
			for (unsigned long i = 0; i < nRows; ++i) *reinterpret_cast<double*>(PyArray_GETPTR1(robj, i)) = nanv;
			PyDict_SetItemString(dict, colName.c_str(), reinterpret_cast<PyObject*>(robj));
		}
	}

	PyTuple_SetItem(tuple, 0, dict);
	std::cout << "Return" << std::endl;
	return tuple;
}

static PyObject* read_xlsx(PyObject* self, PyObject* args, PyObject* kw) {
    char* path = nullptr;
	PyObject* sheet = nullptr;
    bool headers = true;
    int skip_rows = 0;
    int skip_columns = 0;
	char* method = nullptr;
	int num_threads = -1;
	static char* kwlist[] = { // Casting safe aslong as parsing function doesnt modify
		(char*)"path",
		(char*)"sheet",
		(char*)"headers",
		(char*)"skip_rows",
		(char*)"skip_columns",
		(char*)"num_threads",
		NULL
	};
	if (!PyArg_ParseTupleAndKeywords(args, kw, "s|Obiisi", kwlist,
		&path,
		&sheet,
		&headers,
        &skip_rows,
        &skip_columns,
        &method,
        &num_threads)) {
		return NULL;
	}

	std::string sheetName;
	int sheetNumber = 0;
	if (sheet == nullptr) {
		sheetNumber = 1;
	} else {
		if (PyFloat_Check(sheet)) {
			sheetNumber = PyFloat_AsDouble(sheet);
		} else if (PyLong_Check(sheet)) {
			sheetNumber = PyLong_AsLong(sheet);
		} else if (PyUnicode_Check(sheet)) {
			std::cout << "sheet string" << std::endl;
			//TODO: turn into sheetName
		} else {
			PyErr_SetString(PyExc_RuntimeError, "'sheet' must be a string or positive number");
			return NULL;
		}
		if (sheetNumber < 0) {
			PyErr_SetString(PyExc_RuntimeError, "'sheet' must be a string or positive number");
			return NULL;
		}
	}

	if (skip_rows < 0) skip_rows = 0;
	if (skip_columns < 0) skip_columns = 0;

	bool parallel = true;
	if (num_threads == -1) {
		// automatically decide number of threads
		num_threads = std::thread::hardware_concurrency();
		if (num_threads <= 0) {
			num_threads = 1;
		}
		// limit impact on user machine
		if (num_threads > 6 && num_threads <= 10) num_threads = 6;
		// really diminishing returns with higher number of threads
		if (num_threads > 10) num_threads = 10;
	}
	if (num_threads <= 1) {
        num_threads = 1;
        parallel = false;
    }

	try {
		XlsxFile file(path);
		file.mParallelStrings = parallel;
		file.parseSharedStrings();

		XlsxSheet fsheet = sheetNumber > 0 ? file.getSheet(sheetNumber) : file.getSheet(sheetName);
		fsheet.mHeaders = headers;
		// if parallel we need threads for string parsing
		// if "efficient", both sheet & strings need additional thread for decompression (meaning min is 2)
		int act_num_threads = num_threads - parallel * 2 - (num_threads > 1);
		if (act_num_threads <= 0) act_num_threads = 1;
		bool success = fsheet.interleaved(skip_rows, skip_columns, act_num_threads);
		file.finalize();
		if (!success) {
			PyErr_SetString(PyExc_RuntimeError, "There were errors while reading the file, please check output for consistency.");
			return NULL;
		}

		return cells_to_python(file, fsheet);
	} catch (const std::exception& e) {
		PyErr_SetString(PyExc_RuntimeError, e.what());
		return NULL;
	}
}

static PyMethodDef methods[] = {
	{"read_xlsx", reinterpret_cast<PyCFunction>(read_xlsx), METH_VARARGS | METH_KEYWORDS, "TODO Documentation"},
	{NULL, NULL} /* sentinel */
};

static struct PyModuleDef module = {
	PyModuleDef_HEAD_INIT,
	"sheetreader", /* name of module */
	NULL, /* module documentation, may be NULL */
	-1,
	methods
};

PyMODINIT_FUNC PyInit_sheetreader(void) {
	import_array(); // numpy
    if (PyErr_Occurred()) {
        return NULL;
    }
	return PyModule_Create(&module);
}

int main(int argc, char **argv) {
	wchar_t *program = Py_DecodeLocale(argv[0], NULL);
	if (program == NULL) {
		fprintf(stderr, "Fatal error: cannot decode argv[0]\n");
		exit(1);
	}

	/* Add a built-in module, before Py_Initialize */
	PyImport_AppendInittab("sheetreader", PyInit_sheetreader);

	/* Pass argv[0] to the Python interpreter */
	Py_SetProgramName(program);

	/* Initialize the Python interpreter.  Required. */
	Py_Initialize();

	PyMem_RawFree(program);
	return 0;
}