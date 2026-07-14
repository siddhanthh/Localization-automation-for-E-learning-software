#!/usr/bin/env python
"""
Launcher script for the DOCX Translator GUI.
Enables running the GUI directly from the repository root directory.
"""
import sys
import os

# Add the repository root to the Python path to ensure clean package imports
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from docx_translator.gui import main

if __name__ == "__main__":
    main()
