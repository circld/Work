from distutils.core import setup
import py2exe
import pandas
import numpy
import matplotlib

# , '_gtkagg', '_tkagg'

setup(console=['raw.py'],
      options={'py2exe': {'includes': ['zmq.backend.cython'],
                          'excludes': ['zmq.libzmq']}},
      data_files=matplotlib.get_py2exe_datafiles())
