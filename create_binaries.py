# build.py
import os
from PyInstaller.__main__ import run

# Get the absolute path to the src directory
src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "vba_edit")

# Build Word VBA executable
word_vba_args = [
    os.path.join(src_dir, "word_vba.py"),  # Full path to entry point
    "--onefile",
    "--name=word-vba",
    "--clean",
    "--paths",
    src_dir,  # Add src directory to Python path
]

# Build Excel VBA executable
excel_vba_args = [
    os.path.join(src_dir, "excel_vba.py"),  # Full path to entry point
    "--onefile",
    "--name=excel-vba",
    "--clean",
    "--paths",
    src_dir,  # Add src directory to Python path
]

# Create both executables
print("Building word_vba.exe...")
run(word_vba_args)

print("Building excel_vba.exe...")
run(excel_vba_args)
