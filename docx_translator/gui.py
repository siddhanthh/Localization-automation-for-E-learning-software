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

# Color Palette (Modern Dark/Slate Theme)
BG_COLOR = "#0f172a"          # Slate 900
CARD_BG = "#1e293b"           # Slate 800
ACCENT_COLOR = "#0ea5e9"      # Sky 500
ACCENT_HOVER = "#38bdf8"      # Sky 400
TEXT_MAIN = "#f8fafc"         # Slate 50
TEXT_MUTED = "#94a3b8"        # Slate 400
BORDER_COLOR = "#334155"      # Slate 700
SUCCESS_COLOR = "#10b981"     # Emerald 500

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
        self.text_widget.after(0, append_log)


class DocxTranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Document Localizer")
        self.root.geometry("700x700")
        self.root.minsize(620, 650)
        self.root.configure(bg=BG_COLOR)
        
        # Configure Grid weight for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Apply standard styles
        self.setup_styles()
        
        # Main Scrollable/Padded Container
        self.main_frame = tk.Frame(self.root, bg=BG_COLOR, padx=20, pady=20)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(3, weight=1)  # Log output takes most space
        
        # Header Area
        self.header_label = tk.Label(
            self.main_frame, 
            text="DOCX DOCUMENT LOCALIZER", 
            font=("Segoe UI", 16, "bold"), 
            fg=TEXT_MAIN, 
            bg=BG_COLOR,
            anchor="w"
        )
        self.header_label.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        # 1. File Selection Frame
        self.file_frame = tk.LabelFrame(
            self.main_frame, 
            text=" 1. Select Document ", 
            font=("Segoe UI", 10, "bold"),
            fg=ACCENT_COLOR, 
            bg=CARD_BG, 
            bd=1, 
            relief="solid",
            padx=10, 
            pady=10
        )
        self.file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self.file_frame.columnconfigure(1, weight=1)
        
        tk.Label(self.file_frame, text="Word Document:", font=("Segoe UI", 9), fg=TEXT_MAIN, bg=CARD_BG).grid(row=0, column=0, sticky="w", padx=5)
        
        self.input_file_var = tk.StringVar()
        self.input_entry = tk.Entry(
            self.file_frame, 
            textvariable=self.input_file_var,
            font=("Segoe UI", 9),
            bg=BG_COLOR,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            bd=1,
            relief="solid",
            highlightthickness=0
        )
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=10, ipady=4)
        
        self.browse_btn = tk.Button(
            self.file_frame, 
            text="Browse...", 
            font=("Segoe UI", 9, "bold"),
            bg=ACCENT_COLOR,
            fg=BG_COLOR,
            activebackground=ACCENT_HOVER,
            activeforeground=BG_COLOR,
            bd=0,
            cursor="hand2",
            command=self.browse_file,
            padx=15
        )
        self.browse_btn.grid(row=0, column=2, padx=5, ipady=2)
        
        # 2. Options Frame
        self.options_frame = tk.LabelFrame(
            self.main_frame, 
            text=" 2. Translation Configuration ", 
            font=("Segoe UI", 10, "bold"),
            fg=ACCENT_COLOR, 
            bg=CARD_BG, 
            bd=1, 
            relief="solid",
            padx=15, 
            pady=12
        )
        self.options_frame.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        self.options_frame.columnconfigure(1, weight=1)
        self.options_frame.columnconfigure(3, weight=1)
        
        # Target Language Dropdown
        tk.Label(self.options_frame, text="Target Language:", font=("Segoe UI", 9), fg=TEXT_MAIN, bg=CARD_BG).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.lang_display_names = [name for name, code in LANGUAGES]
        self.lang_var = tk.StringVar(value=self.lang_display_names[0])
        
        # Custom TTK Combobox styled via custom theme
        self.lang_combo = ttk.Combobox(
            self.options_frame, 
            textvariable=self.lang_var, 
            values=self.lang_display_names, 
            state="readonly",
            font=("Segoe UI", 9)
        )
        self.lang_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Custom Language ISO Code
        tk.Label(self.options_frame, text="Or Custom ISO:", font=("Segoe UI", 9), fg=TEXT_MAIN, bg=CARD_BG).grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.custom_lang_var = tk.StringVar()
        self.custom_lang_entry = tk.Entry(
            self.options_frame, 
            textvariable=self.custom_lang_var,
            font=("Segoe UI", 9),
            bg=BG_COLOR,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            bd=1,
            relief="solid"
        )
        self.custom_lang_entry.grid(row=0, column=3, sticky="ew", padx=5, pady=5, ipady=3)
        
        # Source Language Code
        tk.Label(self.options_frame, text="Source Language:", font=("Segoe UI", 9), fg=TEXT_MAIN, bg=CARD_BG).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.source_lang_var = tk.StringVar(value="en")
        self.source_entry = tk.Entry(
            self.options_frame, 
            textvariable=self.source_lang_var,
            font=("Segoe UI", 9),
            bg=BG_COLOR,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            bd=1,
            relief="solid"
        )
        self.source_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5, ipady=3)
        
        # Column Header setting
        tk.Label(self.options_frame, text="Column Header:", font=("Segoe UI", 9), fg=TEXT_MAIN, bg=CARD_BG).grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.column_header_var = tk.StringVar(value="Translation")
        self.column_entry = tk.Entry(
            self.options_frame, 
            textvariable=self.column_header_var,
            font=("Segoe UI", 9),
            bg=BG_COLOR,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            bd=1,
            relief="solid"
        )
        self.column_entry.grid(row=1, column=3, sticky="ew", padx=5, pady=5, ipady=3)

        # Separator line
        self.separator = tk.Frame(self.options_frame, height=1, bg=BORDER_COLOR)
        self.separator.grid(row=2, column=0, columnspan=4, sticky="ew", pady=10)

        # Translation Backend Selector
        tk.Label(self.options_frame, text="Translation Engine:", font=("Segoe UI", 9, "bold"), fg=TEXT_MAIN, bg=CARD_BG).grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.backend_var = tk.StringVar(value="google_translator")
        
        self.backend_choices_frame = tk.Frame(self.options_frame, bg=CARD_BG)
        self.backend_choices_frame.grid(row=3, column=1, columnspan=3, sticky="w", pady=5)
        
        self.google_radio = tk.Radiobutton(
            self.backend_choices_frame, 
            text="Google Translate (Free)", 
            variable=self.backend_var, 
            value="google_translator",
            font=("Segoe UI", 9),
            fg=TEXT_MAIN,
            bg=CARD_BG,
            activeforeground=ACCENT_COLOR,
            activebackground=CARD_BG,
            selectcolor=BG_COLOR,
            command=self.toggle_backend_options
        )
        self.google_radio.pack(side="left", padx=(5, 20))
        
        self.gemini_radio = tk.Radiobutton(
            self.backend_choices_frame, 
            text="Gemini AI (High Quality)", 
            variable=self.backend_var, 
            value="gemini",
            font=("Segoe UI", 9),
            fg=TEXT_MAIN,
            bg=CARD_BG,
            activeforeground=ACCENT_COLOR,
            activebackground=CARD_BG,
            selectcolor=BG_COLOR,
            command=self.toggle_backend_options
        )
        self.gemini_radio.pack(side="left", padx=5)

        # Gemini API Key dynamic frame
        self.api_key_frame = tk.Frame(self.options_frame, bg=CARD_BG)
        self.api_key_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(5, 5))
        self.api_key_frame.columnconfigure(1, weight=1)
        
        self.api_label = tk.Label(self.api_key_frame, text="Gemini API Key:", font=("Segoe UI", 9, "bold"), fg=ACCENT_COLOR, bg=CARD_BG)
        self.api_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.api_key_var = tk.StringVar()
        self.api_entry = tk.Entry(
            self.api_key_frame, 
            textvariable=self.api_key_var,
            font=("Segoe UI", 9),
            bg=BG_COLOR,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            bd=1,
            relief="solid",
            show="*"
        )
        self.api_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=5, ipady=3)
        
        self.show_key_var = tk.BooleanVar(value=False)
        self.show_key_btn = tk.Checkbutton(
            self.api_key_frame,
            text="Show",
            variable=self.show_key_var,
            command=self.toggle_key_visibility,
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=CARD_BG,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor=BG_COLOR
        )
        self.show_key_btn.grid(row=0, column=2, padx=5)

        # Toggles Frame
        self.toggles_frame = tk.Frame(self.options_frame, bg=CARD_BG)
        self.toggles_frame.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        
        self.test_mode_var = tk.BooleanVar(value=False)
        self.test_chk = tk.Checkbutton(
            self.toggles_frame, 
            text="Test Mode (5 Tables/50 Cells limit)", 
            variable=self.test_mode_var,
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=CARD_BG,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor=BG_COLOR
        )
        self.test_chk.pack(side="left", padx=(5, 15))
        
        self.review_report_var = tk.BooleanVar(value=True)
        self.review_chk = tk.Checkbutton(
            self.toggles_frame, 
            text="Generate CSV Review Report", 
            variable=self.review_report_var,
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=CARD_BG,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor=BG_COLOR
        )
        self.review_chk.pack(side="left", padx=15)
        
        self.verbose_var = tk.BooleanVar(value=False)
        self.verbose_chk = tk.Checkbutton(
            self.toggles_frame, 
            text="Verbose Logs", 
            variable=self.verbose_var,
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=CARD_BG,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor=BG_COLOR
        )
        self.verbose_chk.pack(side="left", padx=15)

        # 3. Log / Output Terminal Frame
        self.log_frame = tk.LabelFrame(
            self.main_frame, 
            text=" Live Output Logs ", 
            font=("Segoe UI", 10, "bold"),
            fg=ACCENT_COLOR, 
            bg=BG_COLOR, 
            bd=1, 
            relief="solid",
            padx=5, 
            pady=5
        )
        self.log_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 15))
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        self.log_text = ScrolledText(
            self.log_frame, 
            state='disabled', 
            height=10, 
            bg="#030712", 
            fg="#10b981", 
            insertbackground="#10b981", 
            font=("Consolas", 9),
            bd=0,
            highlightthickness=0
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # Setup redirection of logging to the UI text box
        self.log_handler = TextHandler(self.log_text)
        self.log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)
        
        # 4. Action Buttons Frame
        self.action_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        self.action_frame.grid(row=4, column=0, sticky="ew")
        
        self.translate_btn = tk.Button(
            self.action_frame, 
            text="Start Translation", 
            font=("Segoe UI", 10, "bold"),
            bg=SUCCESS_COLOR,
            fg=BG_COLOR,
            activebackground="#34d399",
            activeforeground=BG_COLOR,
            bd=0,
            cursor="hand2",
            command=self.start_translation,
            padx=20,
            pady=5
        )
        self.translate_btn.pack(side="right", padx=5)
        
        self.open_dir_btn = tk.Button(
            self.action_frame, 
            text="Open Output Folder", 
            font=("Segoe UI", 10, "bold"),
            bg=CARD_BG,
            fg=TEXT_MAIN,
            activebackground=BORDER_COLOR,
            activeforeground=TEXT_MAIN,
            bd=1,
            relief="solid",
            cursor="hand2",
            command=self.open_output_dir,
            state="disabled",
            padx=15,
            pady=4
        )
        self.open_dir_btn.pack(side="right", padx=5)
        
        self.status_var = tk.StringVar(value="System Ready.")
        self.status_label = tk.Label(self.action_frame, textvariable=self.status_var, font=("Segoe UI", 9, "italic"), fg=TEXT_MUTED, bg=BG_COLOR)
        self.status_label.pack(side="left", padx=5)
        
        # Initialize visibility state
        self.toggle_backend_options()
        self.last_output_path = None

    def setup_styles(self):
        # Configure standard styling for Combobox dropdown
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure(
            "TCombobox", 
            fieldbackground=BG_COLOR, 
            background=CARD_BG, 
            foreground=TEXT_MAIN,
            bordercolor=BORDER_COLOR,
            arrowcolor=TEXT_MAIN,
            darkcolor=CARD_BG,
            lightcolor=CARD_BG
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", BG_COLOR)],
            foreground=[("readonly", TEXT_MAIN)],
            arrowcolor=[("readonly", TEXT_MAIN)]
        )
        self.root.option_add("*TCombobox*Listbox.background", BG_COLOR)
        self.root.option_add("*TCombobox*Listbox.foreground", TEXT_MAIN)
        self.root.option_add("*TCombobox*Listbox.selectBackground", ACCENT_COLOR)
        self.root.option_add("*TCombobox*Listbox.selectForeground", BG_COLOR)
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 9))

    def toggle_backend_options(self):
        backend = self.backend_var.get()
        if backend == "gemini":
            # Show Gemini API configuration components
            self.api_key_frame.grid()
        else:
            # Hide Gemini API key configuration components
            self.api_key_frame.grid_remove()

    def toggle_key_visibility(self):
        if self.show_key_var.get():
            self.api_entry.configure(show="")
        else:
            self.api_entry.configure(show="*")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx")]
        )
        if filename:
            self.input_file_var.set(os.path.normpath(filename))
            
    def get_selected_target_lang(self):
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
            
        backend = self.backend_var.get()
        api_key = self.api_key_var.get().strip()
        
        if backend == "gemini" and not api_key:
            messagebox.showerror("Error", "Please enter your Gemini API Key to use the Gemini engine.")
            return
            
        target_lang = self.get_selected_target_lang()
        source_lang = self.source_lang_var.get().strip() or "en"
        column_header = self.column_header_var.get().strip() or "Translation"
        test_mode = self.test_mode_var.get()
        verbose = self.verbose_var.get()
        
        # Calculate output path
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_{target_lang}{ext}"
        self.last_output_path = output_path
        
        review_report = f"{base}_review_report.csv" if self.review_report_var.get() else None
        
        # Adjust logging level
        log_level = logging.DEBUG if verbose else logging.INFO
        logging.getLogger().setLevel(log_level)
        for handler in logging.getLogger().handlers:
            handler.setLevel(log_level)
            
        self.set_ui_state(running=True)
        threading.Thread(
            target=self.run_pipeline_thread,
            args=(input_path, output_path, target_lang, source_lang, backend, api_key, column_header, test_mode, review_report),
            daemon=True
        ).start()

    def set_ui_state(self, running):
        state_str = "disabled" if running else "normal"
        self.translate_btn.configure(state=state_str)
        self.browse_btn.configure(state=state_str)
        self.input_entry.configure(state=state_str)
        self.lang_combo.configure(state="disabled" if running else "readonly")
        self.custom_lang_entry.configure(state=state_str)
        self.column_entry.configure(state=state_str)
        self.source_entry.configure(state=state_str)
        self.google_radio.configure(state=state_str)
        self.gemini_radio.configure(state=state_str)
        self.api_entry.configure(state=state_str)
        self.show_key_btn.configure(state=state_str)
        self.test_chk.configure(state=state_str)
        self.review_chk.configure(state=state_str)
        self.verbose_chk.configure(state=state_str)
        
        if running:
            self.translate_btn.configure(bg="#6b7280")  # Grayed out
            self.status_var.set("Running translation pipeline...")
            self.open_dir_btn.configure(state="disabled")
        else:
            self.translate_btn.configure(bg=SUCCESS_COLOR)
            self.status_var.set("Process finished.")
            if self.last_output_path and os.path.exists(self.last_output_path):
                self.open_dir_btn.configure(state="normal")

    def run_pipeline_thread(self, input_path, output_path, target_lang, source_lang, backend, api_key, column_header, test_mode, review_report):
        try:
            config_path = "config.json"
            if not os.path.exists(config_path):
                config_path = None
                
            config = Config(config_path)
            config.settings["translation_backend"] = backend
            
            # Inject API key dynamically from GUI input
            if backend == "gemini" and api_key:
                config.settings["gemini"]["api_key"] = api_key
            
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
                messagebox.showerror("Error", "Translation pipeline failed. Check the logs for details.")
                
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
