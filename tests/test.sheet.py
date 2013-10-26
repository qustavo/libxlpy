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

    def test_writeStr(self):
        # Trial Version return False on 0,0
        self.assertFalse(
                self.sheet.writeStr(0, 0, "Hello World"))

        self.assertTrue(
                self.sheet.writeStr(1, 0, "Hello World"))

    def test_writeNum(self):
        self.assertTrue(
                self.sheet.writeNum(1, 1, 20.0))

        self.assertTrue(
                self.sheet.writeNum(1, 1, 20))

        with self.assertRaises(TypeError):
            self.sheet.writeNum(1, 1, "twenty")

    def test_setMerge(self):
        self.assertTrue(
                self.sheet.setMerge(0, 0, 3, 3))
        
        self.assertFalse(
                self.sheet.setMerge(0, 0, -3, -3))

    def test_setName(self):
        self.assertIsNone(
                self.sheet.setName('foo'))


    def test_setPicture(self):
        img = self.book.addPicture('./logo.png')
        self.assertIsNone(
                self.sheet.setPicture(0, 0, 1, img, 0, 0))

    def test_setPicture2(self):
        img = self.book.addPicture('./logo.png')
        self.assertIsNone(
                self.sheet.setPicture2(0, 0, 100, 100, img, 0, 0))

if __name__ == '__main__':
    unittest.main()
