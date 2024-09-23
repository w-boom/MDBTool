from cx_Freeze import setup, Executable

# 定义可执行文件及其相关的参数
executables = [
    Executable(
        script="MDBTool.py",  
        base="Win32GUI", 
        target_name="MDBTool.exe", 
    )
]

# 设置打包选项
build_options = {
    'packages': ['os', 'tkinter', 'pyodbc', 'subprocess'], 
    'include_files': [], 
    'excludes': [], 
}

setup(
    name="MDBTool", 
    version="1.0.0", 
    description="读取mdb文件",  
    options={"build_exe": build_options},  
    executables=executables  
)
