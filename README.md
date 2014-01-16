# LibXLPy:
A libxl python wrapper

# Building
# (Skip to Installation, if libxl headers and library are directly in your system search paths)
python setup.py build_ext --include-dirs /path/to/libxl/include_c --library-dirs /path/to/libxl/lib64

# Installation:
python setup.py install

# Dependencies:
* libxl

# Usage:
See tests under `./tests` folder.
