import unittest
from libxlpy import *

class TestSheet(unittest.TestCase):
    def setUp(self):
        self.book = Book()
        self.sheet = self.book.addSheet('foo')

    def test_cellType(self):
        type = self.sheet.cellType(0, 0)
        self.assertIn(
                type,
                [SHEETTYPE_CHART, SHEETTYPE_SHEET, SHEETTYPE_UNKNOWN])

    def test_isFormula(self):
        self.assertFalse(
                self.sheet.isFormula(0, 0))

    def test_cellFormat(self):
        self.assertIsNone(self.sheet.cellFormat(0,0))

    def test_setCellFormat(self):
        fmt = self.book.addFormat()
        fmt.setNumFormat(NUMFORMAT_ACCOUNT_D2_CUR)

        self.sheet.setCellFormat(3, 3, fmt)
        fmt = self.sheet.cellFormat(3, 3)

        # num format should be equals for the same cell
        self.assertEqual(NUMFORMAT_ACCOUNT_D2_CUR,
                fmt.numFormat())

    def test_readStr(self):
        (text, fmt) = self.sheet.readStr(0, 0)
        self.assertEqual(text,
                "Created by LibXL trial version. Please buy the LibXL full version for removing this message.")
        self.assertIsNone(fmt)

        self.sheet.writeStr(2, 2, "Hello!")
        (text, fmt) = self.sheet.readStr(2, 2)
        self.assertEqual("Hello!", text)
        self.assertEqual('XLPyFormat', type(fmt).__name__)

    def test_writeStr(self):
        # Trial Version return False on 0,0
        self.assertFalse(
                self.sheet.writeStr(0, 0, "Hello World"))

        self.assertTrue(
                self.sheet.writeStr(1, 0, "Hello World"))

    def test_readNum(self):
        self.sheet.writeNum(3, 3, 200)
        (num, fmt) = self.sheet.readNum(3, 3)
        self.assertEqual(200.0, num)
        self.assertEqual('XLPyFormat', type(fmt).__name__)

    def test_writeNum(self):
        self.assertTrue(
                self.sheet.writeNum(1, 1, 20.0))

        self.assertTrue(
                self.sheet.writeNum(1, 1, 20))

        with self.assertRaises(TypeError):
            self.sheet.writeNum(1, 1, "twenty")

    def test_readBool(self):
        self.assertTrue(self.sheet.writeBool(3, 3, False))
        self.assertFalse(self.sheet.readBool(3, 3)[0])

        self.assertTrue(self.sheet.writeBool(3, 3, True))
        self.assertTrue(self.sheet.readBool(3, 3)[0])

    def test_writeBool(self):
        # error with unregistered lib
        self.assertFalse(self.sheet.writeBool(0, 0, True))
        self.assertTrue(self.sheet.writeBool(1, 1, True))
        self.assertTrue(self.sheet.writeBool(1, 1, False))

    def test_readBlank(self):
        (val, fmt) = self.sheet.readBlank(0, 0)
        self.assertFalse(val)
        self.assertIsNone(fmt)

    def test_writeBlank(self):
        self.assertFalse(self.sheet.writeBlank(0, 0))
        self.sheet.writeBlank(3, 3)
        (val, fmt) = self.sheet.readBlank(3, 3)
        self.assertTrue(val)

    def test_readFormula(self):
        self.assertEqual(
                (None, None),
                self.sheet.readFormula(0, 0)
        )

    def test_writeFormula(self):
        # can't write row 0 in trial version
        self.assertFalse(self.sheet.writeFormula(0, 0, '=3+3'))

        # incorrect token in formula
        self.assertFalse(self.sheet.writeFormula(1, 0, 'incorrect_token'))

        self.assertTrue(self.sheet.writeFormula(1, 0, '=3+3'))
        (val, fmt) = self.sheet.readFormula(1, 0)
        self.assertEqual('3+3', val)

    def test_readComment(self):
        self.assertIsNone(self.sheet.readComment(0, 0))

    def test_writeComment(self):
        self.assertIsNone(
                self.sheet.writeComment(0, 0, 'comment', 'author', 10, 10))

        self.assertEqual('comment',
                self.sheet.readComment(0, 0))

        self.assertIsNone( self.sheet.writeComment(0, 0,
            'passing just this param should also work'))

    def test_isDate(self):
        self.assertFalse(self.sheet.isDate(0, 0))

    def test_readError(self):
        self.assertEqual(ERRORTYPE_NOERROR, self.sheet.readError(0, 0));

    def test_colWidth(self):
        val = self.sheet.colWidth(0)
        self.assertTrue( isinstance(val, float) )

    def test_rowHeight(self):
        val = self.sheet.rowHeight(0)
        self.assertTrue( isinstance(val, float) )

    def test_setCol(self):
        self.assertTrue(self.sheet.setCol(0, 10, 20))
        self.assertTrue(self.sheet.setCol(0, 10, 20.0, hidden = True))

    def test_setRow(self):
        self.assertFalse(self.sheet.setRow(0, 20))
        self.assertTrue(self.sheet.setRow(1, 20))
        self.sheet.setRow(0, 20.0, hidden = True)

    def test_rowHidden(self):
        self.assertFalse(self.sheet.rowHidden(0))
        self.assertFalse(self.sheet.rowHidden(1))

    def test_setRowHidden(self):
        self.assertFalse(self.sheet.setRowHidden(0, True)) # trial version error

        self.sheet.setRowHidden(1, True)
        self.assertTrue(self.sheet.rowHidden(1))

        self.sheet.setRowHidden(1, False)
        self.assertFalse(self.sheet.rowHidden(1))

    def test_colHidden(self):
        self.assertFalse(self.sheet.colHidden(0))
        self.assertFalse(self.sheet.colHidden(1))

    def test_setColHidden(self):
        self.sheet.setColHidden(1, True)
        self.assertTrue(self.sheet.colHidden(1))

        self.sheet.setColHidden(1, False)
        self.assertFalse(self.sheet.colHidden(1))

    def test_getMerge(self):
        self.assertEqual(
            (0, 0, 0, 255),
            self.sheet.getMerge(0, 0)
        )
        self.assertIsNone(self.sheet.getMerge(1, 1))

    def test_setMerge(self):
        self.assertTrue(self.sheet.setMerge(3, 3, 3, 3))
        self.assertEqual(
            (3, 3, 3, 3),
            self.sheet.getMerge(3, 3)
        )
        self.assertFalse(self.sheet.setMerge(0, 0, -3, -3))

    def test_delMerge(self):
        self.sheet.setMerge(3, 3, 5, 5)

        self.assertIsNotNone(self.sheet.getMerge(3, 5))
        self.sheet.delMerge(3, 5)
        self.assertIsNone(self.sheet.getMerge(3, 5))

    def test_setName(self):
        self.assertIsNone(self.sheet.setName('foo'))

    def test_setPicture(self):
        img = self.book.addPicture('./logo.png')
        self.assertIsNone(self.sheet.setPicture(0, 0, 1, img, 0, 0))

    def test_setPicture2(self):
        img = self.book.addPicture('./logo.png')
        self.assertIsNone(self.sheet.setPicture2(0, 0, 100, 100, img, 0, 0))

if __name__ == '__main__':
    unittest.main()
