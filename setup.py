from distutils.core import setup, Extension

module = Extension('libxlpy'
        , sources = [
            'libxlpy.c',
            'book.c',
            'sheet.c',
            'format.c',
            'font.c'
            ]
        , libraries = ['xl']
        )
 
setup (name = 'libxlpy'
        , version = '1.0'
        , description = 'libxl python wrapper'
        , ext_modules = [module]
        )
