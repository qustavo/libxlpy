import unittest
from libxlpy import *

class TestFormat(unittest.TestCase):
    def setUp(self):
        self.book = Book()
        self.format = self.book.addFormat()

    def test_numFormat(self):
        self.assertEqual(0, self.format.numFormat())

if __name__ == '__main__':
    unittest.main()
