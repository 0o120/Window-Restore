import sys
import os
import distutils.dir_util

python_path = os.path.dirname(sys.executable)

tcl_path = os.path.join(python_path, "tcl")

if os.path.exists(tcl_path):
    distutils.dir_util.copy_tree(tcl_path, './env/tcl')
    print(f'Copied: {tcl_path} to ./env/tcl')
