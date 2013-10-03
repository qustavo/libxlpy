import unittest
from libxlpy import *

class TestBook(unittest.TestCase):
    def setUp(self):
        self.book = Book()

    def test_load(self):
        self.assertTrue(
            self.book.load('./book.xls')
        )

        self.assertFalse(
            self.book.load('./unexisting_file')
        )

    def test_addSheet(self):
        sheet = self.book.addSheet('foo')
        self.assertEqual('XLPySheet', type(sheet).__name__)

    def test_getSheet(self):
        sheet = self.book.getSheet(0)
        self.assertIsNone(sheet)

        self.book.addSheet('foo')
        sheet = self.book.getSheet(0)
        self.assertEqual('XLPySheet', type(sheet).__name__)

    def test_sheetType(self):
       self.book.addSheet('foo')
       self.assertEqual(
               self.book.sheetType(0),
               SHEETTYPE_SHEET)

       self.assertEqual(
               self.book.sheetType(99),
               SHEETTYPE_UNKNOWN)

    def test_delSheet(self):
        self.book.addSheet('foo')
        self.book.addSheet('bar')
        self.assertIsNotNone(self.book.getSheet(1))
        self.book.delSheet(1)
        self.assertIsNone(self.book.getSheet(1))

    def test_sheetCount(self):
        self.assertEqual(0, self.book.sheetCount())
        self.book.addSheet('foo')
        self.assertEqual(1, self.book.sheetCount())

    def test_format(self):
        fmt = self.book.format(0)
        self.assertEqual('XLPyFormat', type(fmt).__name__)

    def test_addFormat(self):
        fmt = self.book.addFormat()
        self.assertEqual('XLPyFormat', type(fmt).__name__)

    def test_addFont(self):
        fnt = self.book.addFont()
        self.assertEqual('XLPyFont', type(fnt).__name__)

    def test_formatSize(self):
        num = self.book.formatSize()
        self.book.addFormat()
        self.assertEqual(num + 1, self.book.formatSize())

    def test_activeSheet(self):
        index = self.book.activeSheet()
        self.assertEqual(index, 0)

    @unittest.skip("not working on libxl")
    def test_setActiveSheet(self):
        sheet = self.book.addSheet('foo')
        self.book.setActiveSheet(1)
        self.assertEqual(1, self.book.activeSheet())

    def test_addPicture(self):
        index = self.book.addPicture("./logo.png")
        self.assertEqual(0, index)

    def test_pictureSize(self):
        self.assertEqual(0, self.book.pictureSize())
        index = self.book.addPicture("./logo.png")
        self.assertEqual(1, self.book.pictureSize())

    def test_getPicture(self):
        (t, img) = self.book.getPicture(0)
        self.assertEqual(255, t)
        
        index = self.book.addPicture("./logo.png")
        (t, img) = self.book.getPicture(index)
        self.assertEqual(0, t)

    def test_defaultFont(self):
        (name, size) = self.book.defaultFont()
        self.assertIsInstance(name, str)
        self.assertIsInstance(size, int)

    def test_setDefaultFont(self):
        name, size = "Mono", 14
        self.book.setDefaultFont(name, size)
        self.assertEqual(
                self.book.defaultFont(),
                (name, size))

    def test_setKey(self):
        self.assertIsNone( self.book.setKey("foo", "bar") )

    def test_setLocale(self):
        self.assertTrue(self.book.setLocale("UTF-8"))
        self.assertFalse(self.book.setLocale("BadLocale"))

    def test_errorMessage(self):
        self.assertEqual('ok', self.book.errorMessage())

        # perform some bad op
        self.book.load('ThereIsNoSuchFile.xls')
        self.assertNotEqual('ok', self.book.errorMessage())

if __name__ == '__main__':
    unittest.main()
