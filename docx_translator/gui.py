import os
import sys
import threading
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from docx_translator.config import Config
from docx_translator.pipeline import LocalizationPipeline
from docx_translator.translator import resolve_language_code

# Standard languages mapped to display names
LANGUAGES = [
    ("Thai (th)", "th"),
    ("Spanish (es)", "es"),
    ("French (fr)", "fr"),
    ("German (de)", "de"),
    ("Russian (ru)", "ru"),
    ("Chinese (zh)", "zh"),
    ("Japanese (ja)", "ja"),
    ("Korean (ko)", "ko"),
    ("Arabic (ar)", "ar"),
    ("Hebrew (he)", "he"),
    ("Hindi (hi)", "hi"),
    ("Italian (it)", "it"),
    ("Portuguese (pt)", "pt"),
]

class TextHandler(logging.Handler):
    """Logging handler to direct logs to a Tkinter ScrolledText widget."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append_log():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')
        # Use after to safely call Tkinter from background thread
        self.text_widget.after(0, append_log)


class DocxTranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Translator - Desktop Client")
        self.root.geometry("680x620")
        self.root.minsize(600, 500)
        
        # Configure Grid weight for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Apply standard style/theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Main Frame with padding
        self.main_frame = ttk.Frame(self.root, padding="15")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(2, weight=1)  # Log output takes most space
        
        # 1. File Selection Frame
        self.file_frame = ttk.LabelFrame(self.main_frame, text=" Document Selection ", padding="10")
        self.file_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.file_frame, text="Input File:").grid(row=0, column=0, sticky="w", padx=5)
        
        self.input_file_var = tk.StringVar()
        self.input_entry = ttk.Entry(self.file_frame, textvariable=self.input_file_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=5)
        
        self.browse_btn = ttk.Button(self.file_frame, text="Browse...", command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=5)
        
        # 2. Options Frame
        self.options_frame = ttk.LabelFrame(self.main_frame, text=" Translation Settings ", padding="10")
        self.options_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.options_frame.columnconfigure(1, weight=1)
        self.options_frame.columnconfigure(3, weight=1)
        
        # Target Language Dropdown
        ttk.Label(self.options_frame, text="Target Language:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.lang_display_names = [name for name, code in LANGUAGES]
        self.lang_var = tk.StringVar(value=self.lang_display_names[0])
        self.lang_combo = ttk.Combobox(self.options_frame, textvariable=self.lang_var, values=self.lang_display_names, state="readonly")
        self.lang_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Custom Language input (if they want an ISO code not in the list)
        ttk.Label(self.options_frame, text="Or Custom Code:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.custom_lang_var = tk.StringVar()
        self.custom_lang_entry = ttk.Entry(self.options_frame, textvariable=self.custom_lang_var)
        self.custom_lang_entry.grid(row=0, column=3, sticky="ew", padx=5, pady=5)
        
        # Translation Backend Selector
        ttk.Label(self.options_frame, text="Backend:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.backend_var = tk.StringVar(value="google_translator")
        
        self.google_radio = ttk.Radiobutton(self.options_frame, text="Google Translate (Free)", variable=self.backend_var, value="google_translator")
        self.google_radio.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        self.gemini_radio = ttk.Radiobutton(self.options_frame, text="Gemini (Requires API Key)", variable=self.backend_var, value="gemini")
        self.gemini_radio.grid(row=1, column=2, columnspan=2, sticky="w", padx=5, pady=5)
        
        # Column Header setting
        ttk.Label(self.options_frame, text="Target Column Name:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.column_header_var = tk.StringVar(value="Translation")
        self.column_entry = ttk.Entry(self.options_frame, textvariable=self.column_header_var)
        self.column_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        # Source Language Code setting
        ttk.Label(self.options_frame, text="Source Language:").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.source_lang_var = tk.StringVar(value="en")
        self.source_entry = ttk.Entry(self.options_frame, textvariable=self.source_lang_var)
        self.source_entry.grid(row=2, column=3, sticky="ew", padx=5, pady=5)
        
        # Toggles frame (Test Mode, Review Report, Verbose)
        self.toggles_frame = ttk.Frame(self.options_frame)
        self.toggles_frame.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(5, 0))
        
        self.test_mode_var = tk.BooleanVar(value=False)
        self.test_chk = ttk.Checkbutton(self.toggles_frame, text="Test Mode (Limit 5 tables/50 cells)", variable=self.test_mode_var)
        self.test_chk.pack(side="left", padx=10)
        
        self.review_report_var = tk.BooleanVar(value=True)
        self.review_chk = ttk.Checkbutton(self.toggles_frame, text="Generate Review CSV Report", variable=self.review_report_var)
        self.review_chk.pack(side="left", padx=10)
        
        self.verbose_var = tk.BooleanVar(value=False)
        self.verbose_chk = ttk.Checkbutton(self.toggles_frame, text="Verbose Logs", variable=self.verbose_var)
        self.verbose_chk.pack(side="left", padx=10)

        # 3. Log / Output Terminal Frame
        self.log_frame = ttk.LabelFrame(self.main_frame, text=" Execution Output / Logs ", padding="10")
        self.log_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        self.log_text = ScrolledText(self.log_frame, state='disabled', height=12, bg="black", fg="#00FF00", font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # Setup redirection of logging to the UI text box
        self.log_handler = TextHandler(self.log_text)
        self.log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)
        
        # 4. Action Buttons Frame
        self.action_frame = ttk.Frame(self.main_frame)
        self.action_frame.grid(row=3, column=0, sticky="ew")
        
        self.translate_btn = ttk.Button(self.action_frame, text="Start Translation", command=self.start_translation, style="Accent.TButton")
        self.translate_btn.pack(side="right", padx=5)
        
        self.open_dir_btn = ttk.Button(self.action_frame, text="Open Output Directory", command=self.open_output_dir, state="disabled")
        self.open_dir_btn.pack(side="right", padx=5)
        
        self.status_var = tk.StringVar(value="Ready to translate.")
        self.status_label = ttk.Label(self.action_frame, textvariable=self.status_var, font=("TkDefaultFont", 9, "italic"))
        self.status_label.pack(side="left", padx=5)
        
        # Style definition for Accent Button
        self.style.configure("Accent.TButton", font=("TkDefaultFont", 10, "bold"))
        
        # Track generated output path
        self.last_output_path = None

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx")]
        )
        if filename:
            self.input_file_var.set(os.path.normpath(filename))
            
    def get_selected_target_lang(self):
        # Check custom language entry first
        custom = self.custom_lang_var.get().strip()
        if custom:
            return custom
            
        display_name = self.lang_var.get()
        for name, code in LANGUAGES:
            if name == display_name:
                return code
        return "th"

    def start_translation(self):
        input_path = self.input_file_var.get().strip()
        if not input_path:
            messagebox.showerror("Error", "Please select an input Word document (.docx) first.")
            return
            
        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"Input file not found at: {input_path}")
            return
            
        # Get options
        target_lang = self.get_selected_target_lang()
        source_lang = self.source_lang_var.get().strip() or "en"
        backend = self.backend_var.get()
        column_header = self.column_header_var.get().strip() or "Translation"
        test_mode = self.test_mode_var.get()
        verbose = self.verbose_var.get()
        
        # Auto-calculate default output path
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_{target_lang}{ext}"
        self.last_output_path = output_path
        
        # Setup review report path
        review_report = f"{base}_review_report.csv" if self.review_report_var.get() else None
        
        # Enable verbose logger if selected
        log_level = logging.DEBUG if verbose else logging.INFO
        logging.getLogger().setLevel(log_level)
        for handler in logging.getLogger().handlers:
            handler.setLevel(log_level)
            
        # Run in separate thread so GUI doesn't hang
        self.set_ui_state(running=True)
        threading.Thread(
            target=self.run_pipeline_thread,
            args=(input_path, output_path, target_lang, source_lang, backend, column_header, test_mode, review_report),
            daemon=True
        ).start()

    def set_ui_state(self, running):
        if running:
            self.translate_btn.configure(state="disabled")
            self.browse_btn.configure(state="disabled")
            self.input_entry.configure(state="disabled")
            self.lang_combo.configure(state="disabled")
            self.custom_lang_entry.configure(state="disabled")
            self.column_entry.configure(state="disabled")
            self.source_entry.configure(state="disabled")
            self.google_radio.configure(state="disabled")
            self.gemini_radio.configure(state="disabled")
            self.test_chk.configure(state="disabled")
            self.review_chk.configure(state="disabled")
            self.open_dir_btn.configure(state="disabled")
            self.status_var.set("Translation in progress...")
        else:
            self.translate_btn.configure(state="normal")
            self.browse_btn.configure(state="normal")
            self.input_entry.configure(state="normal")
            self.lang_combo.configure(state="readonly")
            self.custom_lang_entry.configure(state="normal")
            self.column_entry.configure(state="normal")
            self.source_entry.configure(state="normal")
            self.google_radio.configure(state="normal")
            self.gemini_radio.configure(state="normal")
            self.test_chk.configure(state="normal")
            self.review_chk.configure(state="normal")
            if self.last_output_path and os.path.exists(self.last_output_path):
                self.open_dir_btn.configure(state="normal")
            self.status_var.set("Finished.")

    def run_pipeline_thread(self, input_path, output_path, target_lang, source_lang, backend, column_header, test_mode, review_report):
        try:
            # Initialize configuration
            # Check if local config.json exists
            config_path = "config.json"
            if not os.path.exists(config_path):
                config_path = None
                
            config = Config(config_path)
            config.settings["translation_backend"] = backend
            
            # Setup pipeline
            pipeline = LocalizationPipeline(
                config=config,
                target_lang=target_lang,
                source_lang=source_lang,
                tm_path="translation_memory.json",
                glossary_path="glossary.json"
            )
            
            limit_tables = 5 if test_mode else None
            limit_cells = 50 if test_mode else None
            
            logging.info(f"Starting localization pipeline:")
            logging.info(f"  Input: {input_path}")
            logging.info(f"  Output: {output_path}")
            logging.info(f"  Language: {source_lang} -> {target_lang}")
            logging.info(f"  Backend: {backend}")
            
            success = pipeline.translate_document(
                input_path=input_path,
                output_path=output_path,
                column_header=column_header,
                limit_tables=limit_tables,
                limit_cells=limit_cells,
                review_report_path=review_report
            )
            
            if success:
                logging.info("Translation completed successfully!")
                messagebox.showinfo("Success", f"Translation finished successfully!\n\nSaved to:\n{output_path}")
            else:
                logging.error("Translation pipeline failed.")
                messagebox.showerror("Error", "Translation pipeline failed. Check the execution logs for details.")
                
        except Exception as e:
            logging.exception(f"Unexpected error during translation: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}")
        finally:
            self.root.after(0, lambda: self.set_ui_state(running=False))

    def open_output_dir(self):
        if self.last_output_path:
            dir_path = os.path.dirname(os.path.abspath(self.last_output_path))
            try:
                os.startfile(dir_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open directory: {e}")


def main():
    root = tk.Tk()
    app = DocxTranslatorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
