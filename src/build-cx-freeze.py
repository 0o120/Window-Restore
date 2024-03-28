from cx_Freeze import setup, Executable
from config import APP_NAME, APP_VERSION, APP_ICON, APP_ICON_SYS_TRAY
from glob import glob
from zipfile import ZipFile, ZIP_DEFLATED
import os
import shutil


# https://cx-freeze.readthedocs.io/en/stable/setup_script.html


release_path = './build/release'
build_path = './build/cx_freeze'
app_release_name = f"{APP_NAME.lower().replace(' ', '_')}-{APP_VERSION.lower()}"

build_exe_options = {
    'build_exe': build_path,
    # "zip_exclude_packages": [],
    "zip_include_packages": ["collections", "PIL", "pillow", "ctypes", "email", "encodings", "html", "http", "importlib", "json", "logging", "multiprocessing", "packaging", "pyasn1", "rsa", "unittest", "tkinter", "urllib", "xml", "xmlrpc", "pydoc_data", "_ctypes"],
    "packages": ["PySimpleGUI", "psgtray", "pywin32_system32", "win32gui", "win32api", "win32con"],
    "zip_includes": [
        (f'.\\src\\{APP_ICON}', ""),
        (f'.\\src\\{APP_ICON_SYS_TRAY}', "")
    ],
    "excludes": ["test", "asyncio"],
    "include_files": [
        (f'.\\src\\{APP_ICON}'),
        (f'.\\src\\{APP_ICON_SYS_TRAY}'),
        (".\\env\\tcl\\tcl8.6\\", "lib\\tcl8.6\\"),
        (".\\env\\Lib\\site-packages\\PySimpleGUI\\", "lib\\PySimpleGUI\\"),
        (".\\env\\Lib\\site-packages\\\pystray\\", "lib\\\pystray\\"),
        (".\\env\\Lib\\site-packages\\pywin32_system32\\pywintypes310.dll", "lib\\pywintypes310.dll"),
        (".\\env\\Lib\\site-packages\\pywin32_system32\\pythoncom310.dll", "lib\\pythoncom310.dll"),
    ],
    "include_msvcr": False,
}

directory_table = [
    ("ProgramMenuFolder", "TARGETDIR", "."),
    ("MyProgramMenu", "ProgramMenuFolder", "MYPROG~1|My Program"),
]

msi_data = {
    "Directory": directory_table,
    "ProgId": [
        ("Prog.Id", None, None, "This is a description", "IconId", None),
    ],
    "Icon": [
        ("IconId", f'./src/{APP_ICON}'),
    ],
}

bdist_msi_options = {
    "add_to_path": True,
    "data": msi_data,
    "environment_variables": [
        ("E_MYAPP_VAR", "=-*MYAPP_VAR", "1", "TARGETDIR")
    ],
    "upgrade_code": "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}",
}

executables = [
    Executable(
        ".\\src\\main.py",
        target_name=APP_NAME,
        base="Win32GUI",
        icon=f'.\\src\\{APP_ICON}',
        shortcut_name=APP_NAME,
        shortcut_dir="ShortcutDirectory",
    )
]

setup(
    name=APP_NAME,
    version=APP_VERSION,
    description=APP_NAME,
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options
    },
    executables=executables,
)


# Rename the executable to the release name
old_exe = os.path.join(build_path, f"{APP_NAME}.exe")
new_exe = os.path.join(build_path, f"{app_release_name}.exe")
os.rename(old_exe, new_exe)

# Create Release Path
os.makedirs(release_path, exist_ok=True)

zip_filename = f"{app_release_name}-cx_freeze.zip"
zip_file = os.path.join(release_path, zip_filename)
cleanup_files = [zip_file]

# Cleanup
for file in cleanup_files:
    if os.path.exists(file):
        if os.path.isfile(file):
            os.remove(file)
        else:
            shutil.rmtree(file, ignore_errors=True)

# Zip the build directory
with ZipFile(os.path.join(release_path, zip_filename), 'w', compression=ZIP_DEFLATED, compresslevel=4) as zip_object:
    for folder_name, sub_folders, file_names in os.walk(build_path):
        for filename in file_names:
            file_path = os.path.join(folder_name, filename)
            zip_object.write(file_path, os.path.join(APP_NAME, os.path.relpath(file_path, build_path)))
