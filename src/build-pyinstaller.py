import PyInstaller.__main__
from config import APP_ICON, APP_NAME, APP_ICON_SYS_TRAY, APP_VERSION
from zipfile import ZipFile, ZIP_DEFLATED
import os
import shutil


# https://pyinstaller.org/en/stable/usage.html#running-pyinstaller-from-python-code


release_path = './build/release'
build_path = './build/pyinstaller'
app_release_name = f"{APP_NAME.lower().replace(' ', '_')}-{APP_VERSION.lower()}"

PyInstaller.__main__.run([
    './src/main.py',
    f'--distpath={build_path}',
    '--onedir',
    '--windowed',
    f'--add-data=./src/{APP_ICON}:.',
    f'--add-data=./src/{APP_ICON_SYS_TRAY}:.',
    '--noconsole',
    f'--icon=./src/{APP_ICON}',
    f'--name={APP_NAME}',
    '--clean',
    '--noconfirm',
    '--log-level=DEBUG'
])


# Rename the executable to the release name
old_exe = os.path.join(build_path, APP_NAME, f"{APP_NAME}.exe")
new_exe = os.path.join(build_path, APP_NAME, f"{app_release_name}.exe")
os.rename(old_exe, new_exe)

# Create Release Path
os.makedirs(release_path, exist_ok=True)

zip_filename = f"{app_release_name}-pyinstaller.zip"
zip_file = os.path.join(release_path, zip_filename)
spec_file = f"./{APP_NAME}.spec"
temp_dir = os.path.join("./build", APP_NAME)
cleanup_files = [zip_file, spec_file, temp_dir]

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
            zip_object.write(file_path, os.path.relpath(file_path, build_path))