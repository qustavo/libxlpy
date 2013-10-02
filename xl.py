import libxlpy
b = libxlpy.Book()
b.setKey("David Okuniew", "linux-eed1127199a7a21b09012d3748i2d8ne")
b.setLocale("UTF-8")

s = b.addSheet('foo')
b.setActiveSheet(0)
s.writeStr(0, 0, 'gol')
b.addSheet('bar', s)
print b.pictureSize()
print b.save("foo.xls")
