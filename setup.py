from cx_Freeze import setup, Executable

exe = Executable(
    script="frontend.py",
    base="Win32GUI"
)

setup(
    name="Database",
    version="0.1",
    description="first",
    executables=[exe]
)
