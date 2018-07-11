from cx_Freeze import setup, Executable

base = "Win32GUI"

executables = [Executable("ADIStoGWBase.py", base=base,)]

packages = ["idna", "pyqt5", "pandas", "datetime", "xlrd", "XlsxWriter", "numpy", "openpyxl"]
options = {
    'build_exe': {
        'packages':packages,
    },
}

setup(
    name = "ADIStoGWBase",
    options = options,
    version = "0.1",
    description = '<any description>',
    executables = executables
)