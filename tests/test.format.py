import unittest
from libxlpy import *

class TestFormat(unittest.TestCase):
    def setUp(self):
        self.book = Book()
        self.format = self.book.addFormat()

    def test_font(self):
        font = self.format.font()
        self.assertEquals(type(font).__name__, 'XLPyFont')

    def test_setFont(self):
        font = self.book.addFont()
        self.assertTrue(self.format.setFont(font))

    def test_numFormat(self):
        self.assertEqual(0, self.format.numFormat())

    def test_setNumFormat(self):
        self.format.setNumFormat(NUMFORMAT_ACCOUNT)
        self.assertEqual(NUMFORMAT_ACCOUNT, self.format.numFormat())

    def test_formatAlignH(self):
        self.assertIsInstance(self.format.alignH(), int)

    def test_setFormatAlignH(self):
        self.format.setAlignH(ALIGNH_DISTRIBUTED)
        self.assertEquals(self.format.alignH(), ALIGNH_DISTRIBUTED)

    def test_formatAlignV(self):
        self.assertIsInstance(self.format.alignV(), int)

    def test_setFormatAlignV(self):
        self.format.setAlignV(ALIGNV_DISTRIBUTED)
        self.assertEquals(self.format.alignV(), ALIGNV_DISTRIBUTED)

if __name__ == '__main__':
    unittest.main()
