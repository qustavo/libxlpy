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

#define ADD_INT_CONSTANT(v) PyModule_AddIntConstant(mod, #v, v)
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

	ADD_INT_CONSTANT(SHEETTYPE_SHEET);
    ADD_INT_CONSTANT(SHEETTYPE_CHART);
    ADD_INT_CONSTANT(SHEETTYPE_UNKNOWN);

    ADD_INT_CONSTANT(CELLTYPE_NUMBER);
    ADD_INT_CONSTANT(CELLTYPE_STRING);
    ADD_INT_CONSTANT(CELLTYPE_BOOLEAN);
    ADD_INT_CONSTANT(CELLTYPE_BLANK);
    ADD_INT_CONSTANT(CELLTYPE_ERROR);
}
