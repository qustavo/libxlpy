#include <Python.h>
#include <libxl.h>

#include "font.h"

typedef void(set_t)(FontHandle, int);
typedef int (get_t)(FontHandle);

static int
init(XLPyFont *self)
{
    return 0;
}

static void
dealloc(XLPyFont *self)
{
	self->ob_type->tp_free((PyObject*)self);
}

static PyObject *
set(XLPyFont *self, PyObject *args, set_t fn)
{
	int val;
	PyArg_ParseTuple(args, "i", &val);
	fn(self->handler, val);
	Py_RETURN_NONE;
}

static PyObject *
get(XLPyFont *self, get_t fn, int boolean)
{
	int val = fn(self->handler);
	if(!boolean) return Py_BuildValue("i", val);
	if(val) Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

#define GET(name, fn, b) static PyObject *\
	name(XLPyFont *self) {\
		get(self, fn, b);\
	}

#define SET(name, fn) static PyObject *\
	set_##name(XLPyFont *self, PyObject *args) {\
		return set(self, args, fn);\
	}

#define GETSET(name, get, set, boolean) \
	GET(name, get, boolean) \
	SET(name, set)

GETSET(size,       xlFontSize,      xlFontSetSize,      0);
GETSET(italic,     xlFontItalic,    xlFontSetItalic,    1);
GETSET(strike_out, xlFontStrikeOut, xlFontSetStrikeOut, 1);
GETSET(color,      xlFontColor,     xlFontSetColor,     0);
GETSET(bold,       xlFontBold,      xlFontSetBold,      1);
GETSET(script,     xlFontScript,    xlFontSetScript,    0);
GETSET(underline,  xlFontUnderline, xlFontSetUnderline, 0);

static PyObject *
name(XLPyFont *self)
{
	return Py_BuildValue("s", xlFontName(self->handler));
}

static PyObject *
set_name(XLPyFont *self, PyObject *args)
{
	const char *font;
	if(!PyArg_ParseTuple(args, "s", &font)) return NULL;
	xlFontSetName(self->handler, font);
	Py_RETURN_NONE;
}

static PyMethodDef methods[] = {
	{"size", (PyCFunction) size, METH_NOARGS,
		"Returns the handle of the current font. Returns NULL if error occurs."},
	{"setSize", (PyCFunction) set_size, METH_VARARGS,
		"Sets the font for the format. To create a new font use Book::AddFont."},
	{"italic", (PyCFunction) italic, METH_NOARGS,
		"Returns whether the font is italic."},
	{"setItalic", (PyCFunction) set_italic, METH_VARARGS,
		"Turns on/off the italic font,"},
	{"strikeOut", (PyCFunction) strike_out, METH_NOARGS,
		"Returns whether the font is strikeout"},
	{"setStrikeOut", (PyCFunction) set_strike_out, METH_VARARGS,
		"Turns on/off the strikeout font."},
	{"color", (PyCFunction) color, METH_NOARGS,
		"Returns the color of the font."},
	{"setColor", (PyCFunction) set_color, METH_VARARGS,
		"Sets the color of the font."},
	{"bold", (PyCFunction) bold, METH_NOARGS,
		"Returns whether the font is bold."},
	{"setBold", (PyCFunction) set_bold, METH_VARARGS,
		"Turns on/off the bold font."},
	{"script", (PyCFunction) script, METH_NOARGS,
		"Returns the script style of the font."},
	{"setScript", (PyCFunction) set_script, METH_VARARGS,
		"Sets the script style of the font."},
	{"underline", (PyCFunction) underline, METH_NOARGS,
		"Returns the underline style of the font."},
	{"setUnderline", (PyCFunction) set_underline, METH_VARARGS,
		"Sets the underline style of the font."},
	{"name", (PyCFunction) name, METH_NOARGS,
		"Returns the name of the font."},
	{"setName", (PyCFunction) set_name, METH_VARARGS,
		"Sets the name of the font. Default name is \"Arial\"."},
	{NULL}
};

PyTypeObject XLPyFontType = {
   PyObject_HEAD_INIT(NULL)
   0,                         /* ob_size */
   "XLPyFont",              /* tp_name */
   sizeof(XLPyFont),          /* tp_basicsize */
   0,                         /* tp_itemsize */
   (destructor)dealloc,       /* tp_dealloc */
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
   "XLPy Font"  ,             /* tp_doc */
   0,                         /* tp_traverse */
   0,                         /* tp_clear */
   0,                         /* tp_richcompare */
   0,                         /* tp_weaklistoffset */
   0,                         /* tp_iter */
   0,                         /* tp_iternext */
   methods,                   /* tp_methods */
   0,                         /* tp_members */
   0,                         /* tp_getset */
   0,                         /* tp_base */
   0,                         /* tp_dict */
   0,                         /* tp_descr_get */
   0,                         /* tp_descr_set */
   0,                         /* tp_dictoffset */
   (initproc)init,            /* tp_init */
   0,                         /* tp_alloc */
   0,                         /* tp_new */
};
