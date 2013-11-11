#include <Python.h>
#include <libxl.h>

#include "sheet.h"
#include "format.h"

extern PyTypeObject XLPyFormatType;

static int
init(XLPySheet *self)
{
    return 0;
}

static void
dealloc(XLPySheet *self)
{
	self->ob_type->tp_free((PyObject*)self);
}

static PyObject *
cell_type(XLPySheet *self, PyObject *args)
{
    int row, col;
    if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;
    return Py_BuildValue("i",
            xlSheetCellType(self->handler, row, col));
}

static PyObject *
is_formula(XLPySheet *self, PyObject *args)
{
    int row, col;
    if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;
    if (xlSheetIsFormula(self->handler, row, col))
        Py_RETURN_TRUE;
    Py_RETURN_FALSE;
}

static PyObject *
cell_format(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = xlSheetCellFormat(self->handler, row, col);
	if(!fmt) Py_RETURN_NONE;

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return (PyObject *)obj;
}

static PyObject *
set_cell_format(XLPySheet *self, PyObject *args)
{
	const int row, col;
	XLPyFormat *fmt;
	if(!PyArg_ParseTuple(args, "iiO!", &row, &col, &XLPyFormatType, &fmt))
		return NULL;

	xlSheetSetCellFormat(self->handler, row, col, fmt->handler);
	Py_RETURN_NONE;
}

static PyObject *
read_str(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = NULL;
	const char *str = xlSheetReadStr(self->handler, row, col, &fmt);
	if(!str) Py_RETURN_NONE;

	if(!fmt) return Py_BuildValue("(sO)", str, Py_None);

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return Py_BuildValue("(sO)", str, obj);
}

static PyObject *
write_str(XLPySheet *self, PyObject *args)
{
	int row, col;
	const char *val;
	if(!PyArg_ParseTuple(args, "iis", &row, &col, &val)) {
		return NULL;
	}

	if (!xlSheetWriteStr(self->handler, row, col, val, 0)) {
		Py_RETURN_FALSE;
	}
	Py_RETURN_TRUE;
}

static PyObject *
read_num(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = NULL;
	double num;
	num = xlSheetReadNum(self->handler, row, col, &fmt);

	if(!fmt) return Py_BuildValue("(dO)", num, Py_None);

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return Py_BuildValue("(dO)", num, obj);
}

static PyObject *
write_num(XLPySheet *self, PyObject *args)
{
	int row, col;
	double val;
	if(!PyArg_ParseTuple(args, "iid", &row, &col, &val)) {
		return NULL;
	}

	if (!xlSheetWriteNum(self->handler, row, col, val, 0)) {
		Py_RETURN_FALSE;
	}
	Py_RETURN_TRUE;
}

static PyObject *
read_bool(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = NULL;
	int val = xlSheetReadBool(self->handler, row, col, &fmt);

	if(!fmt) return Py_BuildValue("(OO)",
			(0 == val) ? Py_False : Py_True, Py_None);

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return Py_BuildValue("(OO)",
			(0 == val) ? Py_False : Py_True, obj);
}

static PyObject *
write_bool(XLPySheet *self, PyObject *args)
{
	int row, col;
	PyObject *val;
	if(!PyArg_ParseTuple(args, "iiO!", &row, &col, &PyBool_Type, &val)) {
		return NULL;
	}

	if (!xlSheetWriteBool(self->handler, row, col,
				PyObject_IsTrue(val) ? 1 : 0,
				0)) {
		Py_RETURN_FALSE;
	}
	Py_RETURN_TRUE;
}

static PyObject *
read_blank(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = NULL;
	int val = xlSheetReadBlank(self->handler, row, col, &fmt);

	if(!fmt) return Py_BuildValue("(OO)",
			(0 == val) ? Py_False : Py_True, Py_None);

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return Py_BuildValue("(OO)",
			(0 == val) ? Py_False : Py_True, obj);
}

static PyObject *
write_blank(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	if (!xlSheetWriteBlank(self->handler, row, col, 0)) {
		Py_RETURN_FALSE;
	}
	Py_RETURN_TRUE;
}

static PyObject *
read_formula(XLPySheet *self, PyObject *args)
{
	const int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	FormatHandle fmt = NULL;
	const char *val = xlSheetReadFormula(self->handler, row, col, &fmt);

	if(!fmt) return Py_BuildValue("(sO)", val, Py_None);

	XLPyFormat *obj = PyObject_New(XLPyFormat, &XLPyFormatType);
	obj->handler = fmt;
	return Py_BuildValue("(sO)", val, obj);
}

static PyObject *
write_formula(XLPySheet *self, PyObject *args)
{
	int row, col;
	const char *val;
	if(!PyArg_ParseTuple(args, "iis", &row, &col, &val)) return NULL;

	if (!xlSheetWriteFormula(self->handler, row, col, val, 0)) {
		Py_RETURN_FALSE;
	}
	Py_RETURN_TRUE;
}

static PyObject *
read_comment(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	const char *str = xlSheetReadComment(self->handler, row, col);
	return Py_BuildValue("s", str);
}

static PyObject *
write_comment(XLPySheet *self, PyObject *args)
{
	int row, col;
	const char *value, *author = NULL;
	int width = 0, height = 0;
	if(!PyArg_ParseTuple(args, "iis|sii", &row, &col, &value, &author,
				&width, &height)) return NULL;

	xlSheetWriteComment(self->handler, row, col, value, author, width, height);
	Py_RETURN_NONE;
}

static PyObject *
is_date(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	if(xlSheetIsDate(self->handler, row, col))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
read_error(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	return Py_BuildValue("i",
			xlSheetReadError(self->handler, row, col));
}

static PyObject *
col_width(XLPySheet *self, PyObject *args)
{
	int col;
	if(!PyArg_ParseTuple(args, "i", &col)) return NULL;
	return Py_BuildValue("d", xlSheetColWidth(self->handler, col));
}

static PyObject *
row_height(XLPySheet *self, PyObject *args)
{
	int row;
	if(!PyArg_ParseTuple(args, "i", &row)) return NULL;
	return Py_BuildValue("d", xlSheetRowHeight(self->handler, row));
}

static PyObject *
set_col(XLPySheet *self, PyObject *args, PyObject *kwargs)
{
	int colFirst, colLast;
	double width;
	XLPyFormat *fmt = NULL;
	PyObject *hidden = NULL;
	static char *kwlist [] = {
		"colFirst",
		"colLast",
		"width",
		"format",
		"hidden",
		NULL
	};
	if(!PyArg_ParseTupleAndKeywords(args, kwargs, "iid|O!O!", kwlist,
				&colFirst, &colLast, &width,
				&XLPyFormatType, &fmt,
				&PyBool_Type, &hidden)) {
		return NULL;
	}

	if(!hidden) hidden = Py_False;
	if(xlSheetSetCol(self->handler, colFirst, colLast, width,
				(NULL != fmt) ? fmt->handler : 0,
				PyObject_IsTrue(hidden) ? 1 : 0)) {
		Py_RETURN_TRUE;
	}
	Py_RETURN_FALSE;
}

static PyObject *
set_row(XLPySheet *self, PyObject *args, PyObject *kwargs)
{
	int row;
	double height;
	XLPyFormat *fmt = NULL;
	PyObject *hidden = NULL;
	static char *kwlist [] = {
		"row",
		"height",
		"format",
		"hidden",
		NULL
	};
	if(!PyArg_ParseTupleAndKeywords(args, kwargs, "id|O!O!", kwlist,
				&row, &height,
				&XLPyFormatType, &fmt,
				&PyBool_Type, &hidden)) {
		return NULL;
	}

	if(!hidden) hidden = Py_False;
	if(xlSheetSetRow(self->handler, row, height,
				(NULL != fmt) ? fmt->handler : 0,
				PyObject_IsTrue(hidden) ? 1 : 0)) {
		Py_RETURN_TRUE;
	}
	Py_RETURN_FALSE;
}

static PyObject *
row_hidden(XLPySheet *self, PyObject *args)
{
	int row;
	if(!PyArg_ParseTuple(args, "i", &row)) return NULL;
	if(xlSheetRowHidden(self->handler, row))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
set_row_hidden(XLPySheet *self, PyObject *args)
{
	int row;
	PyObject *hidden = NULL;

	if(!PyArg_ParseTuple(args, "iO!", &row, &PyBool_Type, &hidden)) return NULL;
	if(xlSheetSetRowHidden(self->handler, row,
				PyObject_IsTrue(hidden) ? 1 : 0))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
col_hidden(XLPySheet *self, PyObject *args)
{
	int col;
	if(!PyArg_ParseTuple(args, "i", &col)) return NULL;
	if(xlSheetColHidden(self->handler, col))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
set_col_hidden(XLPySheet *self, PyObject *args)
{
	int col;
	PyObject *hidden = NULL;

	if(!PyArg_ParseTuple(args, "iO!", &col, &PyBool_Type, &hidden)) return NULL;
	if(xlSheetSetColHidden(self->handler, col,
				PyObject_IsTrue(hidden) ? 1 : 0))
		Py_RETURN_TRUE;
	Py_RETURN_FALSE;
}

static PyObject *
get_merge(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	int rowFirst, rowLast, colFirst, colLast;
	if(!xlSheetGetMerge(self->handler, row, col, &rowFirst, &rowLast, &colFirst,
		&colLast)) {
		Py_RETURN_NONE;
	}

	return Py_BuildValue("(iiii)", rowFirst, rowLast, colFirst, colLast);
}

static PyObject *
set_merge(XLPySheet *self, PyObject *args)
{
	int rowFirst, rowLast, colFirst, colLast;
	if(!PyArg_ParseTuple(args, "iiii", &rowFirst, &rowLast, &colFirst, &colLast)) {
		return NULL;
	}

	if(xlSheetSetMerge(self->handler, rowFirst, rowLast, colFirst, colLast)) {
		Py_RETURN_TRUE;
	}
	Py_RETURN_FALSE;
}

static PyObject *
del_merge(XLPySheet *self, PyObject *args)
{
	int row, col;
	if(!PyArg_ParseTuple(args, "ii", &row, &col)) return NULL;

	if(!xlSheetDelMerge(self->handler, row, col))
		Py_RETURN_FALSE;
	Py_RETURN_TRUE;
}

static PyObject *
picture_size(XLPySheet *self)
{
	return Py_BuildValue("i", xlSheetPictureSize(self->handler));
}

static PyObject *
get_picture(XLPySheet *self, PyObject *args)
{
	int index;
	if(!PyArg_ParseTuple(args, "i", &index)) return NULL;

	int rowTop, colLeft, rowBottom, colRight, width, height, offset_x, offset_y;

	if(-1 == xlSheetGetPicture(self->handler, index, &rowTop, &colLeft,
		&rowBottom, &colRight, &width, &height, &offset_x, &offset_y)) {
		Py_RETURN_NONE;
	}

	return Py_BuildValue("((ii)iiiiii)",
		rowTop, colLeft, rowBottom, colRight, width, height, offset_x, offset_y);
}

static PyObject *
set_picture(XLPySheet *self, PyObject *args)
{
    int row, col, pictureId;
    double scale;
    int offset_x, offset_y;
    if(!PyArg_ParseTuple(args, "iiidii",
                &row, &col, &pictureId, &scale, &offset_x, &offset_y)) {
        return NULL;
    }

    xlSheetSetPicture(self->handler, row, col,
            pictureId, scale, offset_x, offset_y);
    Py_RETURN_NONE;
}

static PyObject *
set_picture_2(XLPySheet *self, PyObject *args)
{
    int row, col, pictureId;
    int width, height;
    int offset_x, offset_y;
    if(!PyArg_ParseTuple(args, "iiiiiii", &row, &col, &pictureId,
                &width, &height, &offset_x, &offset_y)) {
        return NULL;
    }

    xlSheetSetPicture2(self->handler, row, col,
            pictureId, width, height, offset_x, offset_y);
    Py_RETURN_NONE;
}

static PyObject *
get_hor_page_break(XLPySheet *self, PyObject *args)
{
	int index;
	if(!PyArg_ParseTuple(args, "i", &index)) return NULL;
	return Py_BuildValue("i", xlSheetGetHorPageBreak(self->handler, index));
}

static PyObject *
get_hor_page_break_size(XLPySheet *self)
{
	return Py_BuildValue("i", xlSheetGetHorPageBreakSize(self->handler));
}

static PyObject *
get_ver_page_break(XLPySheet *self, PyObject *args)
{
	int index;
	if(!PyArg_ParseTuple(args, "i", &index)) return NULL;
	return Py_BuildValue("i", xlSheetGetVerPageBreak(self->handler, index));
}

static PyObject *
get_ver_page_break_size(XLPySheet *self)
{
	return Py_BuildValue("i", xlSheetGetVerPageBreakSize(self->handler));
}

static PyObject *
set_hor_page_break(XLPySheet *self, PyObject *args)
{
	int row, pageBreak;
	if(!PyArg_ParseTuple(args, "ii", &row, &pageBreak)) return NULL;

	if(!xlSheetSetHorPageBreak(self->handler, row, pageBreak))
		Py_RETURN_FALSE;
	Py_RETURN_TRUE;
}

static PyObject *
set_ver_page_break(XLPySheet *self, PyObject *args)
{
	int col, pageBreak;
	if(!PyArg_ParseTuple(args, "ii", &col, &pageBreak)) return NULL;

	if(!xlSheetSetVerPageBreak(self->handler, col, pageBreak))
		Py_RETURN_FALSE;
	Py_RETURN_TRUE;
}

static PyObject *
set_name(XLPySheet *self, PyObject *args)
{
    const char *name;
    if(!PyArg_ParseTuple(args, "s", &name)) {
        return NULL;
    }

    xlSheetSetName(self->handler, name);
    Py_RETURN_NONE;
}

static PyMethodDef methods[] = {
    {"cellType", (PyCFunction) cell_type, METH_VARARGS,
        "Returns cell's type."},
    {"isFormula", (PyCFunction) is_formula, METH_VARARGS,
        "Checks that cell contains a formula."},
	{"cellFormat", (PyCFunction) cell_format, METH_VARARGS,
		"Returns cell's format. It can be changed by user."},
	{"setCellFormat", (PyCFunction) set_cell_format, METH_VARARGS,
		"Sets cell's format."},
	{"readStr", (PyCFunction) read_str, METH_VARARGS,
		"Reads a string and its format from cell. "
		"Returns a (text, format) tuple, "
		"Format will be None if specified cell doesn't contain string or error occurs."},
	{"writeStr", (PyCFunction) write_str, METH_VARARGS,
		"Writes a string into cell with specified format (if present). Returns False if error occurs."},
	{"readNum", (PyCFunction) read_num, METH_VARARGS,
		"Reads a number or date/time and its format from cell. "
		"Use Book::dateUnpack() for extract date/time parts. "
		"Returns a tuple with (num, format)."},
    {"writeNum", (PyCFunction) write_num, METH_VARARGS,
        "Writes a string into cell with specified format. "
        "If format is not present then format is ignored. "
        "Returns False if error occurs."},
	{"readBool", (PyCFunction) read_bool, METH_VARARGS,
		"Reads a bool value and its format from cell. "
		"If format is None then error occurs. "
		"Returns a tuple with (num, format)."},
	{"writeBool", (PyCFunction) write_bool, METH_VARARGS,
		"Writes a bool value into cell with specified format. "
		"If format is None then format is ignored. "
		"Returns False if error occurs."},
	{"readBlank", (PyCFunction) read_blank, METH_VARARGS,
		"Reads format from blank cell. Returns False if specified cell isn't blank or error occurs."},
	{"writeBlank", (PyCFunction) write_blank, METH_VARARGS,
		"Writes blank cell with specified format."},
	{"readFormula", (PyCFunction) read_formula, METH_VARARGS,
		"Reads a formula and its format from cell. "
		"Returns None if specified cell doesn't contain formula or error occurs."},
	{"writeFormula", (PyCFunction) write_formula, METH_VARARGS,
		"Writes a formula into cell with specified format. "
		"If format equals None then format is ignored. "
		"Returns False if error occurs."},
	{"readComment", (PyCFunction) read_comment, METH_VARARGS,
		"Reads a comment from specified cell."},
	{"writeComment", (PyCFunction) write_comment, METH_VARARGS,
		"Writes a comment to the cell. Parameters:\n"
		"(row, col) - cell's position;\n"
		"value - comment string;\n"
		"author - author string;\n"
		"width - width of text box in pixels;\n"
		"height - height of text box in pixels."},
	{"isDate", (PyCFunction) is_date, METH_VARARGS,
		"Checks that cell contains a date or time value."},
	{"readError", (PyCFunction) read_error, METH_VARARGS,
		"Reads error from cell."},
	{"colWidth", (PyCFunction) col_width, METH_VARARGS,
		"Returns column width."},
	{"rowHeight", (PyCFunction) row_height, METH_VARARGS,
		"Returns row height."},
	{"setCol", (PyCFunction) set_col, METH_KEYWORDS,
		"Sets column width and format for all columns from colFirst to colLast. "
		"Column width measured as the number of characters of the maximum digit width of the numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. "
		"If format equals None then format is ignored. "
		"Columns may be hidden. Returns False if error occurs"},
	{"setRow", (PyCFunction) set_row, METH_KEYWORDS,
		"Sets row height and format. Row height measured in point size. "
		"If format equals None then format is ignored. "
		"Row may be hidden. Returns False if error occurs"},
	{"rowHidden", (PyCFunction) row_hidden, METH_VARARGS,
		"Returns whether row is hidden."},
	{"setRowHidden", (PyCFunction) set_row_hidden, METH_VARARGS,
		"Hides row."},
	{"colHidden", (PyCFunction) col_hidden, METH_VARARGS,
		"Returns whether col is hidden."},
	{"setColHidden", (PyCFunction) set_col_hidden, METH_VARARGS,
		"Hides col."},
	{"getMerge", (PyCFunction) get_merge, METH_VARARGS,
		"Gets merged cells for cell at row, col. "
		"Result is a tuple of rowFirst, rowLast, colFirst, colLast. "
		"Returns None if error occurs."},
	{"setMerge", (PyCFunction) set_merge, METH_VARARGS,
		"Sets merged cells for range: rowFirst - rowLast, colFirst - colLast. "
		"Returns False if error occurs."},
	{"delMerge", (PyCFunction) del_merge, METH_VARARGS,
		"Removes merged cells. Returns False if error occurs."},
	{"pictureSize", (PyCFunction) picture_size, METH_NOARGS,
		"Returns a number of pictures in this worksheet."},
	{"getPicture", (PyCFunction) get_picture, METH_VARARGS,
		"Returns a workbook picture index at position index in worksheet. "
		"Returns a tuple with the following values: \n"
		"(rowTop, colLeft) - top left position of picture;\n"
		"(rowBottom, colRight) - bottom right position of picture;\n"
		"width - width of picture in pixels; \n"
		"height - height of picture in pixels; \n"
		"offset_x - horizontal offset of picture in pixels; \n"
		"offset_y - vertical offset of picture in pixels. "
		"Use Book::getPicture() for extracting binary data of picture by workbook picture index. "
		"Returns None if error occurs."},
    {"setPicture", (PyCFunction) set_picture, METH_VARARGS,
        "Sets a picture with pictureId identifier at position row and col with scale factor and offsets in pixels. "
        "Use Book::addPicture() for getting picture identifier."},
    {"setPicture2", (PyCFunction) set_picture_2, METH_VARARGS,
        "Sets a picture with pictureId identifier at position row and col with custom size and offsets in pixels. "
        "Use Book::addPicture() for getting a picture identifier."},
    {"getHorPageBreak", (PyCFunction) get_hor_page_break, METH_VARARGS,
    	"Returns row with horizontal page break at position index."},
    {"getHorPageBreakSize", (PyCFunction) get_hor_page_break_size, METH_NOARGS,
    	"Returns a number of horizontal page breaks in the sheet."},
    {"getVerPageBreak", (PyCFunction) get_ver_page_break, METH_VARARGS,
    	"Returns column with vertical page break at position index."},
    {"getVerPageBreakSize", (PyCFunction) get_ver_page_break_size, METH_NOARGS,
    	"Returns a number of vertical page breaks in the sheet."},
    {"setHorPageBreak", (PyCFunction) set_hor_page_break, METH_VARARGS,
    	"Sets/removes a horizontal page break (sets if True, removes if False). "
    	"Returns False if error occurs."},
    {"setVerPageBreak", (PyCFunction) set_ver_page_break, METH_VARARGS,
    	"Sets/removes a vertical page break (sets if True, removes if False). "
    	"Returns False if error occurs."},
    {"setName", (PyCFunction) set_name, METH_VARARGS,
        "Sets the name of the sheet."},
	{NULL}
};

PyTypeObject XLPySheetType = {
   PyObject_HEAD_INIT(NULL)
   0,                         /* ob_size */
   "XLPySheet",               /* tp_name */
   sizeof(XLPySheet),         /* tp_basicsize */
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
   "XLPy Sheet",                 /* tp_doc */
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
