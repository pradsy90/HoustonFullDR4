import sys
import platform
locate_python = sys.exec_prefix
print (str(locate_python) + " is where it's located")
print (str(platform.python_version())  + " is the version")