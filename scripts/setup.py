# 好好学习 天天向上
# {2023/11/19} {20:49}
from distutils.core import setup
import py2exe

setup(console=["mainGui.py"],py_modules=["command","sheetplot","os","tkinter","openpyxl"])