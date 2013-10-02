#include <Python.h>
#include <libxl.h>

#include "book.h"
#include "sheet.h"
#include "format.h"
#include "font.h"

extern PyTypeObject XLPySheetType;
extern PyTypeObject XLPyFormatType;
extern PyTypeObject XLPyFontType;

static int
init(XLPyBook *self)
{
	self->handler = xlCreateBook();
    return 0;
}

static void
dealloc(XLPyBook *self)
{
	xlBookRelease(self->handler);
	self->ob_type->tp_free((PyObject*)self);
}

static PyObject *
load(XLPyBook *self, PyObject *args)
{
	const char* fileName;
	int size;
	PyObject *raw = NULL;
	if(!PyArg_ParseTuple(args, "s#|O!", &fileName, &size, &PyBool_Type, &raw))
		return NULL;

	if (raw && PyObject_IsTrue(raw)) {
		// fileName is treated as the buffer
		if(!xlBookLoadRaw(self->handler, fileName, size)) {
			Py_RETURN_FALSE;
		}
	}
	else {
		if(!xlBookLoad(self->handler, fileName)) {
			Py_RETURN_FALSE;
		}
	}

	Py_RETURN_TRUE;
}

static PyObject *
save(XLPyBook *self, PyObject *args)
{
	const char *fileName = NULL;
	if(!PyArg_ParseTuple(args, "|s", &fileName)) return NULL;

	// No argument, saveRaw
	if(!fileName) {
		const char *raw;
		unsigned int size;
		if(!xlBookSaveRaw(self->handler, &raw, &size)) {
			Py_RETURN_FALSE;
		}
		return Py_BuildValue("s#", raw, size);
	}

	// If argument provided, save to file
	if(!xlBookSave(self->handler, fileName)) {
		Py_RETURN_FALSE;
	}

	Py_RETURN_TRUE;
}

static PyObject *
add_sheet(XLPyBook *self, PyObject *args)
{
	const char *name;
	PyObject *initSheet = NULL;
	if(!PyArg_ParseTuple(args, "s|O!", &name, &XLPySheetType, &initSheet))
		return NULL;
	
	SheetHandle sheet = xlBookAddSheet(self->handler, name,
			(NULL == initSheet) ? NULL : ((XLPySheet *)initSheet)->handler);

	if (!sheet) Py_RETURN_NONE;

	XLPySheet *obj = PyObject_New(XLPySheet, &XLPySheetType);
	obj->handler = sheet;
	return (PyObject *)obj;
}

static PyObject *
get_sheet(XLPyBook *self, PyObject *args)
{
	int num;
	if(!PyArg_ParseTuple(args, "i", &num))
		return NULL;

	SheetHandle sheet = xlBookGetSheet(self->handler, num);
	if (!sheet)
		Py_RETURN_NONE;

	XLPySheet *obj = PyObject_New(XLPySheet, &XLPySheetType);
	obj->handler = sheet;
	return (PyObject *)obj;
}

static PyObject *
sheet_type(XLPyBook *self, PyObject *args)
{
	int num;
	if(!PyArg_ParseTuple(args, "i", &num)) return NULL;
	return Py_BuildValue("i", xlBookSheetType(self->handler, num));
}

static PyObject *
del_sheet(XLPyBook *self, PyObject *args)
{
	int num;
	if(!PyArg_ParseTuple(args, "i", &num)) return NULL;

	if(!xlBookDelSheet(self->handler, num))
		Py_RETURN_FALSE;
	Py_RETURN_TRUE;
}

static PyObject *
sheet_count(XLPyBook *self)
{
	int count = xlBookSheetCount(self->handler);
	return Py_BuildValue("i", count);
}

static PyObject *
add_format(XLPyBook *self, PyObject *args)
{
	PyObject *initFormat = NULL;
	if(!PyArg_ParseTuple(args, "|O!", &XLPyFormatType, &initFormat)) return NULL;
	FormatHandle format = xlBookAddFormat(self->handler,
			(NULL == initFormat) ? NULL : ((XLPyFormat *)initFormat)->handler);

	if (!format) Py_RETURN_NONE;
	
	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = format;
	return (PyObject *)obj;
}

static PyObject *
add_font(XLPyBook *self, PyObject *args)
{
	PyObject *initFont = NULL;
	if(!PyArg_ParseTuple(args, "|O!", &XLPyFontType, &initFont)) return NULL;
	FontHandle font = xlBookAddFont(self->handler,
			(NULL == initFont) ? NULL : ((XLPyFont *)initFont)->handler);

	if (!font) Py_RETURN_NONE;
	
	XLPyFont *obj = PyObject_New(XLPyFont, &XLPyFontType);
	obj->handler = font;
	return (PyObject *)obj;
}

static PyObject *
format(XLPyBook *self, PyObject *args)
{
	int num;
	if(!PyArg_ParseTuple(args, "i", &num)) return NULL;
	FormatHandle format = xlBookFormat(self->handler, num);
	if(!format) Py_RETURN_NONE;

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = format;
	return (PyObject *)obj;
}

static PyObject *
format_size(XLPyBook *self)
{
	return Py_BuildValue("i",
			xlBookFormatSize(self->handler));
}

static PyObject *
active_sheet(XLPyBook *self)
{
	return Py_BuildValue("i",
			xlBookActiveSheet(self->handler));
}

static PyObject *
set_active_sheet(XLPyBook *self, PyObject *args)
{
	int num;
	if(!PyArg_ParseTuple(args, "i", &num)) return NULL;
	xlBookSetActiveSheet(self->handler, num);
	Py_RETURN_NONE;
}

static PyObject *
picture_size(XLPyBook *self, PyObject *args)
{
	return Py_BuildValue("i",
			xlBookPictureSize(self->handler));
}

static PyObject *
set_key(XLPyBook *self, PyObject *args)
{
	const char *name, *key;
	if(!PyArg_ParseTuple(args, "ss", &name, &key)) return NULL;

	xlBookSetKey(self->handler, name, key);
	Py_RETURN_NONE;
}

static PyObject *
set_locale(XLPyBook *self, PyObject *args)
{
	const char *locale;
	if(!PyArg_ParseTuple(args, "s", &locale)) return NULL;

	if (xlBookSetLocale(self->handler, locale))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
error_message(XLPyBook *self)
{
	return Py_BuildValue("s", xlBookErrorMessage(self->handler));
}

static PyMethodDef methods[] = {
	{"load", (PyCFunction) load, METH_VARARGS,
		"When second arg is True, it loads a xls-file from buffer."
		"When is not present or False, it loads xls-file from file path."
		"Returns False if error occurs."},
	{"save", (PyCFunction) save, METH_VARARGS,
		"When arg is a string, saves current workbook into xls-file."
		"When no args, returns xls-file as a buffer."
		"Returns False if error occurs"},
	{"addSheet", (PyCFunction) add_sheet, METH_VARARGS,
		"Adds a new sheet to this book, returns the sheet object. "
		"Use initSheet parameter if you wish to copy an existing sheet. "
		"Note initSheet must be only from this book. Returns None if error occurs."},
	{"getSheet", (PyCFunction) get_sheet, METH_VARARGS,
		"Gets pointer to a sheet with specified index. "
		"Returns None if error occurs."},
	{"sheetType", (PyCFunction) sheet_type, METH_VARARGS,
		"Returns type of sheet with specified index."},
	{"delSheet", (PyCFunction) del_sheet, METH_VARARGS,
		"Deletes a sheet with specified index. Returns false if error occurs. "},
	{"sheetCount", (PyCFunction) sheet_count, METH_VARARGS,
		"Returns a number of sheets in this book."},
	{"format", (PyCFunction) format, METH_VARARGS,
		"Returns a format with defined index. "
		"Index must be less than return value of formatSize() method."},
	{"addFormat", (PyCFunction) add_format, METH_VARARGS,
		"Adds a new format to the workbook, initial parameters can be copied from other format. "
		"Returns None if error occurs."},
	{"addFont", (PyCFunction) add_font, METH_VARARGS,
		"Adds a new font to the workbook, initial parameters can be copied from other font. "
		"Returns None if error occurs."},
	{"formatSize", (PyCFunction) format_size, METH_NOARGS,
		"Returns a number of formats in this book."},
	{"activeSheet", (PyCFunction) active_sheet, METH_NOARGS,
		"Returns an active sheet index in this workbook."},
	{"setActiveSheet", (PyCFunction) set_active_sheet, METH_VARARGS,
		"Sets an active sheet index in this workbook."},
	{"pictureSize", (PyCFunction) picture_size, METH_NOARGS,
		"Returns a number of pictures in this workbook."},
	{"setKey", (PyCFunction) set_key, METH_VARARGS,
		"Sets customer's license key."},
	{"setLocale", (PyCFunction) set_locale, METH_VARARGS,
		"Sets the locale for this library. "
		"The locale argument is the same as the locale argument in setlocale() function from C Run-Time Library. "
		"For example, value \"en_US.UTF-8\" allows to use non-ascii characters in Linux or Mac. "
        "It accepts the special value \"UTF-8\" for using UTF-8 character encoding in Windows and other operating systems. "
		"It has no effect for unicode projects with wide strings (with _UNICODE preprocessor variable). "
		"Returns True if a valid locale argument is given."},
	{"errorMessage", (PyCFunction) error_message, METH_NOARGS,
		"Returns the last error message."},
	{NULL}
};

PyTypeObject
XLPyBookType = {
   PyObject_HEAD_INIT(NULL)
   0,                         /* ob_size */
   "XLPyBook",                /* tp_name */
   sizeof(XLPyBook),          /* tp_basicsize */
   0,                         /* tp_itemsize */
   (destructor)dealloc,/* tp_dealloc */
   0,                         /* tp_print */
   0,                         /* tp_getattr */
   0,                         /* tp_setattr */
   0,                         /* tp_compare */
   0,                         /* tp_repr */
   0,                         /* tp_as_number */
   0,                         /* tp_as_sequence */
   0,                         /* tp_as_mapping */
   0,                         /* tp_hash */
   0,                         /* tp_call */
   0,                         /* tp_str */
   0,                         /* tp_getattro */
   0,                         /* tp_setattro */
   0,                         /* tp_as_buffer */
   Py_TPFLAGS_DEFAULT | Py_TPFLAGS_BASETYPE, /* tp_flags*/
   "XLPy Book",                 /* tp_doc */
   0,                         /* tp_traverse */
   0,                         /* tp_clear */
   0,                         /* tp_richcompare */
   0,                         /* tp_weaklistoffset */
   0,                         /* tp_iter */
   0,                         /* tp_iternext */
   methods,            /* tp_methods */
   0,                         /* tp_members */
   0,                         /* tp_getset */
   0,                         /* tp_base */
   0,                         /* tp_dict */
   0,                         /* tp_descr_get */
   0,                         /* tp_descr_set */
   0,                         /* tp_dictoffset */
   (initproc)init,     /* tp_init */
   0,                         /* tp_alloc */
   0,                         /* tp_new */
};
