#include <Python.h>
#include <libxl.h>

#include "format.h"
#include "font.h"

extern PyTypeObject XLPyFontType;

typedef void(op_t)(FormatHandle, int);

static int
init(XLPyFormat *self)
{
    return 0;
}

static void
dealloc(XLPyFormat *self)
{
	self->ob_type->tp_free((PyObject*)self);
}

static PyObject *
generic_set(XLPyFormat *self, PyObject *args, op_t op)
{
	int val;
	PyArg_ParseTuple(args, "i", &val);
	op(self->handler, val);
	Py_RETURN_NONE;
}

static PyObject *
font(XLPyFormat *self)
{
	FontHandle font = xlFormatFont(self->handler);
	if(!font) Py_RETURN_NONE;

	XLPyFont *obj = PyObject_New(XLPyFont, &XLPyFontType);
	obj->handler = font;
	return (PyObject *)obj;
}

static PyObject *
set_font(XLPyFormat *self, PyObject *args)
{
	PyObject *font;
	if(!PyArg_ParseTuple(args, "O!", &XLPyFontType, &font)) return NULL;

	if(xlFormatSetFont(self->handler, ((XLPyFont *)font)->handler))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
num_format(XLPyFormat *self)
{
    return Py_BuildValue("i",
            xlFormatNumFormat(self->handler));
}

static PyObject *
set_num_format(XLPyFormat *self, PyObject *args)
{
	return generic_set(self, args, xlFormatSetNumFormat);
}

static PyObject *
align_h(XLPyFormat *self)
{
	return Py_BuildValue("i", xlFormatAlignH(self->handler));
}

static PyObject *
set_align_h(XLPyFormat *self, PyObject *args)
{
	return generic_set(self, args, xlFormatSetAlignH);
}

static PyObject *
align_v(XLPyFormat *self)
{
	return Py_BuildValue("i", xlFormatAlignV(self->handler));
}

static PyObject *
set_align_v(XLPyFormat *self, PyObject *args)
{
	return generic_set(self, args, xlFormatSetAlignV);
}

static PyMethodDef methods[] = {
	{"font", (PyCFunction) font, METH_NOARGS,
		"Returns the handle of the current font. "
		"Returns None if error occurs."},
	{"setFont", (PyCFunction) set_font, METH_VARARGS,
		"Sets the font for the format. "
		"To create a new font use Book::addFont()"},
    {"numFormat", (PyCFunction) num_format, METH_NOARGS,
        "Returns the number format identifier."},
	{"setNumFormat", (PyCFunction) set_num_format, METH_VARARGS,
		"Sets the number format identifier. "
		"The identifier must be a valid built-in number format identifier or the identifier of a custom number format. "
		"To create a custom format use Book::AddCustomNumFormat()"},
	{"alignH", (PyCFunction) align_h, METH_NOARGS,
		"Returns the horizontal alignment."},
	{"setAlignH", (PyCFunction) set_align_h, METH_VARARGS,
		"Sets the horizontal alignment."},
	{"alignV", (PyCFunction) align_v, METH_NOARGS,
		"Returns the vertical alignment."},
	{"setAlignV", (PyCFunction) set_align_v, METH_VARARGS,
		"Sets the vertical alignment."},

	{NULL}
};

PyTypeObject XLPyFormatType = {
   PyObject_HEAD_INIT(NULL)
   0,                         /* ob_size */
   "XLPyFormat",              /* tp_name */
   sizeof(XLPyFormat),        /* tp_basicsize */
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
   "XLPy Format",             /* tp_doc */
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
