import unittest
from libxlpy import *

class TestFont(unittest.TestCase):
    def setUp(self):
        self.book = Book(BOOK_XLS)
        self.font = self.book.addFont()

    def testSize(self):
        self.assertIsInstance(self.font.size(), int)

    def testSetSize(self):
        self.font.setSize(24)
        self.assertEquals(self.font.size(), 24)

    def testItalic(self):
        self.assertIsInstance(self.font.italic(), bool)

    def testSetItalic(self):
        self.font.setItalic(True)
        self.assertTrue(self.font.italic())

        self.font.setItalic(False)
        self.assertFalse(self.font.italic())

    def testStrikeOut(self):
        self.assertIsInstance(self.font.strikeOut(), bool)

    def testSetStrikeOut(self):
        self.font.setStrikeOut(True)
        self.assertTrue(self.font.strikeOut())

        self.font.setStrikeOut(False)
        self.assertFalse(self.font.strikeOut())

    def testColor(self):
        self.assertIsInstance(self.font.color(), int)

    def testSetColor(self):
        self.font.setColor(COLOR_LIGHTORANGE)
        self.assertEquals(self.font.color(), COLOR_LIGHTORANGE)

    def testBold(self):
        self.assertIsInstance(self.font.bold(), bool)

    def testSetBold(self):
        self.font.setBold(True)
        self.assertTrue(self.font.bold())

        self.font.setBold(False)
        self.assertFalse(self.font.bold())

    def testScript(self):
        self.assertIsInstance(self.font.script(), int)

    def testSetScript(self):
        self.font.setScript(SCRIPT_SUPER)
        self.assertEquals(self.font.script(), SCRIPT_SUPER)

    def testUnderline(self):
        self.assertIsInstance(self.font.underline(), int)

    def testSetUnderline(self):
        self.font.setUnderline(UNDERLINE_DOUBLE)
        self.assertEquals(self.font.underline(), UNDERLINE_DOUBLE)

    def testName(self):
        self.assertIsInstance(self.font.name(), str)

    def testSetName(self):
        self.font.setName("Sans")
        self.assertEquals("Sans", self.font.name())

if __name__ == '__main__':
    unittest.main()
