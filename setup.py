from cx_Freeze import setup, Executable
base = None    
executables = [Executable("python.py", base=base)]
packages = ["idna","openpyxl"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}
setup(
    name = "Form Validation",
    options = options,
    version = "1.0",
    description = 'Form Validation',
    executables = executables
)