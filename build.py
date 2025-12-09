"""
Build script to create standalone executable
Run this on Windows to create the .exe file
"""

import PyInstaller.__main__
import os
import shutil
import sys

# Get the directory of this script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Clean previous builds
for folder in ['build', 'dist']:
    folder_path = os.path.join(BASE_DIR, folder)
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)

# Determine path separator for --add-data (Windows uses ;, Linux/Mac uses :)
sep = ';' if sys.platform == 'win32' else ':'

# PyInstaller arguments
args = [
    os.path.join(BASE_DIR, 'src', 'main.py'),
    '--name=FormFiller',
    '--onefile',  # Single executable
    '--windowed',  # No console window
    f'--add-data={os.path.join(BASE_DIR, "Forms", "Ajman-University-Logo.png")}{sep}Forms',
    f'--distpath={os.path.join(BASE_DIR, "dist")}',
    f'--workpath={os.path.join(BASE_DIR, "build")}',
    f'--specpath={BASE_DIR}',
    '--clean',
]

# Add icon if available
icon_path = os.path.join(BASE_DIR, 'Forms', 'Ajman-University-Logo.ico')
if os.path.exists(icon_path):
    args.append(f'--icon={icon_path}')

print("Building executable...")
PyInstaller.__main__.run(args)

print("\nBuild complete!")
print(f"Executable created at: {os.path.join(BASE_DIR, 'dist', 'FormFiller.exe')}")
