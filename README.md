# DOCX Document Localizer - Desktop Client

An AI-powered document translation and layout preservation tool designed to translate specific columns in DOCX tables (e.g., translation tables exported from **Articulate Storyline** or **Articulate Rise 360**) while preserving the exact layout, styles, and format of the original document.

This tool introduces a user-friendly **Desktop GUI Client** and a **Standalone Executable (.exe)**, making the tool accessible to non-technical users without requiring Python installations or command-line usage.

---

## Quick Start

To run the application without installing Python or setting up a workspace:

1. Go to the **Releases** section on this GitHub repository.
2. Download the latest version of **`DOCX_Localizer.exe`**.
3. Double-click **`DOCX_Localizer.exe`** to launch the desktop application.

---

## Desktop GUI Features

- **Document Selection**: Simple "Browse" button to select your input `.docx` file.
- **Target Language Dropdown**: Easy selection of standard languages (Thai, Spanish, French, Chinese, Japanese, etc.) or input of custom ISO language codes.
- **Multiple Translation Engines**: Toggle between Google Translate (Free) and Google's Gemini API (Requires API Key).
- **Execution Log Monitor**: An embedded console displaying real-time pipeline progress and diagnostic information.
- **Layout Conservation**: Automatically copies table designs, cell widths, heights, and XML formatting from the original document so the output matches perfectly.
- **Generates Post-Translation CSV Reports**: Outputs a CSV spreadsheet for review and verification of all translated vs skipped text cells.

---

## Developer / Local Execution

If you have Python installed and want to run or edit the codebase:

### 1. Setup Virtual Environment & Install Dependencies
```bash
# Create venv
python -m venv venv

# Activate venv
# On Windows:
venv\Scripts\activate.ps1

# Install requirements
pip install -r requirements.txt
pip install -e .
```

### 2. Launch the GUI
```bash
python run_gui.py
```

### 3. Rebuild the Standalone Executable
If you modify the source files, recompile the standalone `.exe` using:
```bash
python build_exe.py
```
The compiled output will be generated inside the `dist/` folder.

---

## Advanced CLI Usage (Backup)

For CLI environments, you can still run the tool directly:
```bash
# General CLI Command
docx-translate <input_file> <target_language_code> [options]

# Example: Translate a Storyline export to Thai using Google Translate
docx-translate export.docx th --test-mode
```

---

## Project Structure

```
docx-translator/
├── docx_translator/        # Core source package
│   ├── gui.py              # Tkinter-based Desktop GUI Client
│   ├── pipeline.py         # Table parsing and pipeline runner
│   ├── translator.py       # Rate-limited translation client classes
│   ├── layout.py           # Layout and format restoration logic
│   └── config.py           # Configuration parser and settings
├── build_exe.py            # PyInstaller packaging script
├── run_gui.py              # GUI Launcher script
├── run.py                  # CLI Launcher script
├── setup.py                # Package setup script
├── requirements.txt        # Python dependency list
├── config.json             # Backend configuration profiles
└── dist/                   # Directory containing compiled standalone executable (.exe)
```
