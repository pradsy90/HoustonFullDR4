""""
from distutils.core import setup
import py2exe, sys, os, numpy,tkinter,threading,site
sys.argv.append('py2exe')
Mydata_files = [('images', ['c:\\cutover\\cccore1.png'])]
setup(
    options = {
        'py2exe': {
            #'includes': ['tkinter','threading', 'multiprocessing'],
            #'packages': ['numpy','os','sys','site']
            #'excludes': ['numpy.f2py.diagnose','numpy.array_api._typing','numpy._typing._array_like','azure']
        }
    },
    console=[
        {
            "script":'cutoverlaunch.py',
            "dest_base":'ccHouston'
        }],
    data_files = Mydata_files
)

"""
