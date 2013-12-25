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

    def test_wrap(self):
        self.assertIsInstance(self.format.wrap(), bool)

    def test_setWrap(self):
        self.format.setWrap(True)
        self.assertTrue(self.format.wrap())

        self.format.setWrap(False)
        self.assertFalse(self.format.wrap())


    def test_rotation(self):
        self.assertIsInstance(self.format.rotation(), int)

    def test_setRotation(self):
        self.format.setRotation(0)
        self.assertEquals(0, self.format.rotation())

        self.format.setRotation(91)
        self.assertEquals(91, self.format.rotation())

        self.format.setRotation(255)
        self.assertEquals(255, self.format.rotation())

    def test_indent(self):
        self.assertIsInstance(self.format.indent(), int)

    def test_setIndent(self):
        self.format.setIndent(0)
        self.assertEquals(self.format.indent(), 0)

        self.format.setIndent(15)
        self.assertEquals(self.format.indent(), 15)

        self.format.setIndent(999)
        self.assertLessEqual(self.format.indent(), 15)

    def test_shrinkToFit(self):
        self.assertIsInstance(self.format.shrinkToFit(), bool)

    def test_setShrinkToFit(self):
        self.format.setShrinkToFit(True)
        self.assertTrue(self.format.shrinkToFit())

        self.format.setShrinkToFit(False)
        self.assertFalse(self.format.shrinkToFit())

    def test_borders(self):
        self.format.setBorder(BORDERSTYLE_DASHDOT)
        self.assertEquals(BORDERSTYLE_DASHDOT, self.format.borderLeft())
        self.assertEquals(BORDERSTYLE_DASHDOT, self.format.borderRight())
        self.assertEquals(BORDERSTYLE_DASHDOT, self.format.borderTop())
        self.assertEquals(BORDERSTYLE_DASHDOT, self.format.borderBottom())

        self.format.setBorderColor(COLOR_SEAGREEN);
        self.assertEquals(COLOR_SEAGREEN, self.format.borderLeftColor())
        self.assertEquals(COLOR_SEAGREEN, self.format.borderRightColor())
        self.assertEquals(COLOR_SEAGREEN, self.format.borderTopColor())
        self.assertEquals(COLOR_SEAGREEN, self.format.borderBottomColor())

        self.format.setBorderLeft(BORDERSTYLE_THICK)
        self.format.setBorderRight(BORDERSTYLE_THICK)
        self.format.setBorderTop(BORDERSTYLE_THICK)
        self.format.setBorderBottom(BORDERSTYLE_THICK)
        self.assertEquals(BORDERSTYLE_THICK, self.format.borderLeft())
        self.assertEquals(BORDERSTYLE_THICK, self.format.borderRight())
        self.assertEquals(BORDERSTYLE_THICK, self.format.borderTop())
        self.assertEquals(BORDERSTYLE_THICK, self.format.borderBottom())

        self.format.setBorderLeftColor(COLOR_INDIGO)
        self.format.setBorderRightColor(COLOR_INDIGO)
        self.format.setBorderTopColor(COLOR_INDIGO)
        self.format.setBorderBottomColor(COLOR_INDIGO)
        self.assertEquals(COLOR_INDIGO, self.format.borderLeftColor())
        self.assertEquals(COLOR_INDIGO, self.format.borderRightColor())
        self.assertEquals(COLOR_INDIGO, self.format.borderTopColor())
        self.assertEquals(COLOR_INDIGO, self.format.borderBottomColor())

        self.format.setBorderDiagonal(BORDERDIAGONAL_UP)
        self.format.setBorderDiagonalColor(COLOR_PLUM)
        self.assertEquals(BORDERDIAGONAL_UP, self.format.borderDiagonal())
        self.assertEquals(COLOR_PLUM, self.format.borderDiagonalColor())


if __name__ == '__main__':
    unittest.main()
