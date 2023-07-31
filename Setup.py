from cx_Freeze import setup, Executable

includefiles = ['Functions.py', 'config.ini', 'helper.py', 'ChExcel.py']
includes = []
excludes = []

setup(
    name='Prep ToolKit',
    version='0.4',
    description='A Translation file prepping utility',
    author='Enzo Agosta',
    author_email='agosta.enzowork@gmail.com',
    options={'build_exe': {'includes': includes, 'excludes': excludes,
             'include_files': includefiles}},
    executables=[Executable('Main.py', base="Win32GUI")]
)
