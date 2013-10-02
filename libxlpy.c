#include <Python.h>
#include <libxl.h>

extern PyTypeObject XLPyBookType;
extern PyTypeObject XLPySheetType;
extern PyTypeObject XLPyFormatType;
extern PyTypeObject XLPyFontType;

#define INIT_TYPE(Type) \
	Type.tp_new = PyType_GenericNew; \
	if (PyType_Ready(&Type) < 0) return; \
	Py_INCREF(&Type)

void initlibxlpy(void)
{
    PyObject* mod;
    mod = Py_InitModule3("libxlpy", NULL, "A libxl python wrapper");
    if (mod == NULL) return;
     
    // Init Classes
	INIT_TYPE(XLPyBookType);
    INIT_TYPE(XLPySheetType);
	INIT_TYPE(XLPyFormatType);
	INIT_TYPE(XLPyFontType);

    // Module's Public API
    PyModule_AddObject(mod, "Book", (PyObject* )&XLPyBookType);

	PyModule_AddIntConstant(mod, "SHEETTYPE_SHEET", SHEETTYPE_SHEET);
    PyModule_AddIntConstant(mod, "SHEETTYPE_CHART", SHEETTYPE_CHART);
    PyModule_AddIntConstant(mod, "SHEETTYPE_UNKNOWN", SHEETTYPE_UNKNOWN);
}
