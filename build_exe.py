#!/usr/bin/env python
"""
PyInstaller build script to compile the DOCX Translator GUI to a single standalone EXE.
"""
import os
import sys
import shutil
import PyInstaller.__main__

def build_standalone_exe():
    print("Building standalone executable...")
    
    # Define PyInstaller arguments
    args = [
        'run_gui.py',                     # Script to bundle
        '--onefile',                       # Package into a single executable
        '--windowed',                      # Hide console window (GUI-only)
        '--name=DOCX_Translator',         # Output executable name
        '--collect-all=docx',             # Collect all data files, metadata, and binaries for python-docx
        '--collect-all=deep_translator',  # Collect all files for deep_translator
        '--clean',                        # Clean cache before build
    ]
    
    # Run PyInstaller
    PyInstaller.__main__.run(args)
    print("\nBuild complete. Check the 'dist' directory for the executable.")

if __name__ == "__main__":
    build_standalone_exe()
