#!/usr/bin/env python3
"""
Docu Split - PDF Document Splitter
Windows 10 Flat Design - PDF Only
- Split PDF by ID (groups same ID pages together)
- Split PDF by custom page ranges (multiple ranges supported)
- Multi-criteria file naming (up to 4 fields)
- Light theme only
- Flat scrollbars (no burger lines)
"""

import re
import csv
import PyPDF2
import fitz
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import threading
import os
import sys
import json

# Color scheme - Light theme only
COLORS = {
    'bg': '#f0f0f0',
    'card': '#ffffff',
    'text': '#000000',
    'text_light': '#000000',
    'border': '#d0d0d0',
    'header_bg': '#e8e8e8',
    'button_bg': '#e0e0e0',
    'button_fg': '#333333',
    'accent': '#5A9EFF',
    'accent_fg': '#ffffff',
    'success': '#4CAF50',
    'warning': '#f39c12',
    'error': '#e74c3c',
    'info': '#17a2b8',
    'info_bg': '#d1ecf1',
    'info_text': '#0c5460',
    'log_bg': '#f8f8f8',
    'tab_selected': '#ffffff',
    'tab_unselected': '#e8e8e8',
    'scrollbar_bg': '#f0f0f0',
    'scrollbar_trough': '#f0f0f0',
    'scrollbar_active': '#e0e0e0'
}

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Docu Split - PDF Document Splitter")
        self.root.geometry("1280x950")
        self.root.minsize(1100, 800)
        
        # Set window icon
        try:
            self.root.iconbitmap('docsplit.ico')
        except:
            try:
                img = tk.PhotoImage(file='docsplit.ico')
                self.root.iconphoto(True, img)
            except:
                pass
        
        # Set default cursor
        self.root.config(cursor="arrow")
        
        # Settings file path
        self.settings_file = Path.home() / ".doc_split_settings.json"
        
        # Criteria - Separate for Tab 1 and Tab 2
        self.criteria_tab1 = []  # For documents with IDs
        self.criteria_tab2 = []  # For documents without IDs
        
        self.tab1_naming_selections = []
        self.tab2_naming_selections = []
        self.naming_separator = tk.StringVar(value="_")
        self.filename_suffix = tk.StringVar(value="")
        
        # Page ranges for Tab 3
        self.page_ranges = tk.StringVar(value="")
        
        # CSV settings
        self.csv_output_folder = tk.StringVar(value="")
        self.csv_filename = tk.StringVar(value="extracted_data")
        self.csv_filename_tab2 = tk.StringVar(value="extracted_data_tab2")
        self.csv_filename_tab3 = tk.StringVar(value="page_extract")
        
        # Button text variables
        self.tab1_button_text = tk.StringVar(value="SPLIT BY ID")
        self.tab2_button_text = tk.StringVar(value="SPLIT BY NAME")
        self.tab3_button_text = tk.StringVar(value="EXTRACT PAGES")
        
        # Tab name overrides
        self.tab1_button_override = False
        self.tab2_button_override = False
        self.tab1_name = "Documents with IDs"
        self.tab2_name = "Documents without IDs"
        
        # Tip visibility settings
        self.show_tip_tab1 = tk.BooleanVar(value=True)
        self.show_tip_tab2 = tk.BooleanVar(value=True)
        self.tip_frame_tab1 = None
        self.tip_frame_tab2 = None
        
        # Load saved settings
        self.load_settings()
        
        # Create menu bar
        self.create_menu()
        
        # Configure ttk styles
        self.configure_styles()
        
        # Create main container with scrollbar
        self.create_scrollable_main()
        
        # Create header
        self.create_header()
        
        # Create notebook (tabs)
        self.create_notebook()
        
        # Initialize all tabs
        self.init_tab1()
        self.init_tab2()
        self.init_pagesplit_tab()
        
        # Sync buttons
        self.sync_tab1_button_to_name()
        self.sync_tab2_button_to_name()
    
    def configure_styles(self):
        """Configure ttk styles for light theme"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Notebook tabs
        style.configure('TNotebook', background=COLORS['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', background=COLORS['tab_unselected'], 
                       padding=[12, 5], borderwidth=0, font=('Segoe UI', 9))
        style.map('TNotebook.Tab',
                  background=[('selected', COLORS['tab_selected'])],
                  foreground=[('selected', COLORS['text'])],
                  padding=[('selected', [12, 5]), ('!selected', [10, 3])])
        
        # Entry
        style.configure('TEntry', fieldbackground=COLORS['card'], 
                       borderwidth=1, relief='solid', highlightthickness=0,
                       foreground=COLORS['text'])
        style.map('TEntry', fieldbackground=[('focus', COLORS['card'])])
        
        # Button
        style.configure('TButton', background=COLORS['button_bg'], 
                       borderwidth=1, relief='solid', padding=6,
                       foreground=COLORS['button_fg'])
        style.map('TButton', background=[('active', COLORS['border'])])
        
        # Progressbar
        style.configure('TProgressbar', background=COLORS['accent'], 
                       troughcolor=COLORS['border'], borderwidth=0)
        
        # Combobox
        style.configure('TCombobox', fieldbackground=COLORS['card'], 
                       borderwidth=1, relief='solid', highlightthickness=0,
                       foreground=COLORS['text'])
        
        # LabelFrame
        style.configure('TLabelframe', background=COLORS['card'], 
                       borderwidth=1, relief='solid')
        style.configure('TLabelframe.Label', background=COLORS['card'], 
                       foreground=COLORS['text_light'])
        
        # Frame
        style.configure('TFrame', background=COLORS['card'])
    
    def apply_theme(self):
        """Apply light theme to all widgets"""
        self.root.configure(bg=COLORS['bg'])
        
        if hasattr(self, 'main_canvas'):
            self.main_canvas.configure(bg=COLORS['bg'])
        if hasattr(self, 'scrollable_frame'):
            self.scrollable_frame.configure(bg=COLORS['bg'])
        
        if hasattr(self, 'header_frame'):
            for widget in self.header_frame.winfo_children():
                if isinstance(widget, tk.Label):
                    widget.configure(bg=COLORS['card'])
                elif isinstance(widget, tk.Frame):
                    widget.configure(bg=COLORS['border'])
        
        if hasattr(self, 'status_preview'):
            self.status_preview.configure(bg=COLORS['card'], fg=COLORS['text_light'])
        if hasattr(self, 'preview_label'):
            self.preview_label.configure(bg=COLORS['card'])
        if hasattr(self, 'preview_label_tab2'):
            self.preview_label_tab2.configure(bg=COLORS['card'])
        if hasattr(self, 'page_preview_label'):
            self.page_preview_label.configure(bg=COLORS['card'])
        
        for log in ['log_text', 'log_text2', 'log_text3']:
            if hasattr(self, log):
                getattr(self, log).configure(bg=COLORS['log_bg'], fg=COLORS['text'])
        
        self.update_criteria_display_tab1()
        self.update_criteria_display_tab2()
    
    def create_scrollable_main(self):
        """Create a scrollable main container with flat Windows 10 scrollbar"""
        self.main_canvas = tk.Canvas(self.root, highlightthickness=0, borderwidth=0, bg=COLORS['bg'])
        self.main_canvas.pack(side='left', fill='both', expand=True)
        
        self.scrollbar = tk.Scrollbar(self.root, orient='vertical', command=self.main_canvas.yview,
                                     bg=COLORS['scrollbar_bg'], troughcolor=COLORS['scrollbar_trough'],
                                     activebackground=COLORS['scrollbar_active'], relief='flat', borderwidth=0,
                                     highlightthickness=0)
        self.scrollbar.pack(side='right', fill='y')
        
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollable_frame = tk.Frame(self.main_canvas, bg=COLORS['bg'])
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
        
        self.main_canvas.bind('<Configure>', self._on_canvas_configure)
        self._bind_mousewheel()
    
    def _on_frame_configure(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox('all'))
    
    def _on_canvas_configure(self, event):
        self.main_canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _bind_mousewheel(self):
        def _on_mousewheel(event):
            self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_frame(event):
            self.scrollable_frame.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_frame(event):
            self.scrollable_frame.unbind_all("<MouseWheel>")
        
        self.scrollable_frame.bind("<Enter>", _bind_to_frame)
        self.scrollable_frame.bind("<Leave>", _unbind_from_frame)
    
    def create_menu(self):
        menubar = tk.Menu(self.root, tearoff=0)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        
        # Criteria Menu - Separate for Tab 1 and Tab 2
        criteria_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Criteria", menu=criteria_menu)
        criteria_menu.add_command(label="Tab 1: Add Criterion", command=lambda: self.add_criterion_dialog(tab=1))
        criteria_menu.add_command(label="Tab 1: Remove Last", command=lambda: self.remove_last_criterion(tab=1))
        criteria_menu.add_separator()
        criteria_menu.add_command(label="Tab 2: Add Criterion", command=lambda: self.add_criterion_dialog(tab=2))
        criteria_menu.add_command(label="Tab 2: Remove Last", command=lambda: self.remove_last_criterion(tab=2))
        criteria_menu.add_separator()
        criteria_menu.add_command(label="Reset Tab 1 to Defaults", command=lambda: self.reset_criteria(tab=1))
        criteria_menu.add_command(label="Reset Tab 2 to Defaults", command=lambda: self.reset_criteria(tab=2))
        
        # View menu with Restore Tip options
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Restore Tab 1 Tip", command=self.restore_tip_tab1)
        view_menu.add_command(label="Restore Tab 2 Tip", command=self.restore_tip_tab2)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Edit Tab 1 Button", command=self.edit_tab1_button)
        settings_menu.add_command(label="Edit Tab 2 Button", command=self.edit_tab2_button)
        settings_menu.add_separator()
        settings_menu.add_command(label="Sync Buttons to Tab Names", command=self.sync_buttons_to_tab_names)
        settings_menu.add_separator()
        settings_menu.add_command(label="Rename Tab 1", command=self.rename_tab1)
        settings_menu.add_command(label="Rename Tab 2", command=self.rename_tab2)
        settings_menu.add_separator()
        settings_menu.add_command(label="Reset Tab Names", command=self.reset_tab_names)
        settings_menu.add_command(label="Save Settings", command=self.save_settings)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Quick Start Guide", command=self.show_quick_start)
        help_menu.add_command(label="About", command=self.show_about)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Quick Start Guide", command=self.show_quick_start)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_separator()  # Add separator
        help_menu.add_command(label="☕ Support the Project", command=self.open_donation_page)  # Add this line
    
    def create_header(self):
        self.header_frame = tk.Frame(self.scrollable_frame, bg=COLORS['card'])
        self.header_frame.pack(fill='x', pady=(0, 10))
        
        # Left side - title and subtitle
        title_label = tk.Label(self.header_frame, text="Docu Split", font=('Segoe UI', 18, 'bold'), 
                            bg=COLORS['card'], fg=COLORS['text'])
        title_label.pack(anchor='w', padx=15, pady=(10, 0))
        
        subtitle_label = tk.Label(self.header_frame, text="Split and organize PDF documents with precision", 
                                bg=COLORS['card'], fg=COLORS['text_light'], font=('Segoe UI', 9))
        subtitle_label.pack(anchor='w', padx=15, pady=(0, 10))
        
        # Right side - donation button (NEW)
        donation_btn = tk.Button(self.header_frame, text="☕ Support", command=self.open_donation_page,
                                bg=COLORS['accent'], fg='white', font=('Segoe UI', 9, 'bold'),
                                relief='flat', padx=15, pady=5, cursor="hand2")
        donation_btn.place(relx=0.95, rely=0.15, anchor='ne')
        
        separator = tk.Frame(self.header_frame, height=1, bg=COLORS['border'])
        separator.pack(fill='x', padx=15, pady=(0, 10))
    
    def create_notebook(self):
        self.notebook = ttk.Notebook(self.scrollable_frame)
        self.notebook.pack(fill='both', expand=True, pady=(10, 0))
        
        self.tab1 = tk.Frame(self.notebook, bg=COLORS['card'])
        self.tab2 = tk.Frame(self.notebook, bg=COLORS['card'])
        self.tab3 = tk.Frame(self.notebook, bg=COLORS['card'])
        
        self.notebook.add(self.tab1, text=f"📄 {self.tab1_name}")
        self.notebook.add(self.tab2, text=f"🔍 {self.tab2_name}")
        self.notebook.add(self.tab3, text=f"✂️ Page Split")
    
    # ==================== TIP MANAGEMENT ====================
    def get_default_csv_path(self):
        """Get the default CSV path from Tab 1 output folder and filename"""
        output_folder = self.csv_output_folder.get()
        csv_filename = self.csv_filename.get()
        
        if output_folder and csv_filename:
            # Ensure .csv extension
            if not csv_filename.endswith('.csv'):
                csv_filename += '.csv'
            default_path = Path(output_folder) / csv_filename
            if default_path.exists():
                return str(default_path)
        
        # Also try the output folder from Tab 1 when a PDF was processed
        if hasattr(self, 'output_folder') and self.output_folder.get():
            default_path = Path(self.output_folder.get()) / "extracted_data.csv"
            if default_path.exists():
                return str(default_path)
        
        return None
    
    def toggle_tip_tab1(self):
        """Toggle tip visibility for Tab 1"""
        if self.show_tip_tab1.get():
            self.show_tip_tab1.set(False)
            if self.tip_frame_tab1:
                self.tip_frame_tab1.destroy()
                self.tip_frame_tab1 = None
        else:
            self.show_tip_tab1.set(True)
            self.create_tip_tab1()
        self.save_settings()

    def create_tip_tab1(self):
        """Create and show tip for Tab 1"""
        if not self.show_tip_tab1.get() or not hasattr(self, 'tab1_header_frame'):
            return
        
        # Remove existing tip if any
        if self.tip_frame_tab1:
            try:
                self.tip_frame_tab1.destroy()
            except:
                pass
        
        # Create new tip frame
        self.tip_frame_tab1 = tk.Frame(self.tab1_header_frame, bg=COLORS['info_bg'], relief='flat', bd=1, highlightthickness=0)
        self.tip_frame_tab1.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        # Tip label (no X button)
        tip_label = tk.Label(self.tip_frame_tab1, text="💡 Tip: When adding/editing a criterion, you can optionally enable 'Stop Text' to stop reading at a specific phrase (useful for multi-line values)", 
                             bg=COLORS['info_bg'], fg=COLORS['info_text'], font=('Segoe UI', 9), wraplength=650, justify='left')
        tip_label.pack(side='left', fill='x', expand=True, padx=5, pady=5)

    def create_tip_tab2(self):
        """Create and show tip for Tab 2"""
        if not self.show_tip_tab2.get() or not hasattr(self, 'tab2_header_frame'):
            return
        
        # Remove existing tip if any
        if self.tip_frame_tab2:
            try:
                self.tip_frame_tab2.destroy()
            except:
                pass
        
        # Create new tip frame
        self.tip_frame_tab2 = tk.Frame(self.tab2_header_frame, bg=COLORS['info_bg'], relief='flat', bd=1, highlightthickness=0)
        self.tip_frame_tab2.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        # Tip label (no X button)
        tip_label = tk.Label(self.tip_frame_tab2, text="💡 Tip: Enable 'Stop Text' when adding/editing a criterion to stop reading at a specific phrase. Example: For 'Description', set Stop Text to 'on the day of' to capture only the description without the following text.", 
                             bg=COLORS['info_bg'], fg=COLORS['info_text'], font=('Segoe UI', 9), wraplength=650, justify='left')
        tip_label.pack(side='left', fill='x', expand=True, padx=5, pady=5)

    def toggle_tip_tab2(self):
        """Toggle tip visibility for Tab 2"""
        if self.show_tip_tab2.get():
            self.show_tip_tab2.set(False)
            if self.tip_frame_tab2:
                self.tip_frame_tab2.destroy()
                self.tip_frame_tab2 = None
        else:
            self.show_tip_tab2.set(True)
            self.create_tip_tab2()
        self.save_settings()

    def restore_tip_tab1(self):
        """Restore tip for Tab 1 (called from menu)"""
        if not self.show_tip_tab1.get():
            self.toggle_tip_tab1()

    def restore_tip_tab2(self):
        """Restore tip for Tab 2 (called from menu)"""
        if not self.show_tip_tab2.get():
            self.toggle_tip_tab2()
    
    # ==================== SETTINGS ====================
    def load_settings(self):
        defaults_tab1 = [
            {"name": "Document ID", "prefix": "ID Number:", "type": "id", "stop_text": ""},
            {"name": "Name", "prefix": "Name:", "type": "text", "stop_text": ""},
            {"name": "Description", "prefix": "Description:", "type": "text", "stop_text": ""},
            {"name": "Date", "prefix": "Date:", "type": "text", "stop_text": ""}
        ]
        defaults_tab2 = [
            {"name": "Document ID", "prefix": "ID:", "type": "id", "stop_text": ""},
            {"name": "Name", "prefix": "Name:", "type": "text", "stop_text": ""},
            {"name": "Description", "prefix": "Description:", "type": "text", "stop_text": ""},
            {"name": "Date", "prefix": "Date:", "type": "text", "stop_text": ""}
        ]
        
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r') as f:
                    s = json.load(f)
                    self.tab1_name = s.get("tab1_name", "Documents with IDs")
                    self.tab2_name = s.get("tab2_name", "Documents without IDs")
                    self.tab1_button_override = s.get("tab1_button_override", False)
                    self.tab2_button_override = s.get("tab2_button_override", False)
                    self.tab1_button_text.set(s.get("tab1_button_text", "SPLIT DOCUMENTS"))
                    self.tab2_button_text.set(s.get("tab2_button_text", "SPLIT DOCUMENTS"))
                    self.tab1_naming_selections = s.get("tab1_naming_selections", [0])
                    self.tab2_naming_selections = s.get("tab2_naming_selections", [0])
                    self.naming_separator.set(s.get("naming_separator", "_"))
                    self.filename_suffix.set(s.get("filename_suffix", ""))
                    self.show_tip_tab1.set(s.get("show_tip_tab1", True))
                    self.show_tip_tab2.set(s.get("show_tip_tab2", True))
                    
                    loaded_tab1 = s.get("criteria_tab1", None)
                    loaded_tab2 = s.get("criteria_tab2", None)
                    
                    if loaded_tab1 and len(loaded_tab1) > 0:
                        self.criteria_tab1 = loaded_tab1
                    else:
                        self.criteria_tab1 = defaults_tab1.copy()
                        
                    if loaded_tab2 and len(loaded_tab2) > 0:
                        self.criteria_tab2 = loaded_tab2
                    else:
                        self.criteria_tab2 = defaults_tab2.copy()
                        
                    self.csv_output_folder.set(s.get("csv_output_folder", ""))
                    self.csv_filename.set(s.get("csv_filename", "extracted_data"))
                    self.csv_filename_tab2.set(s.get("csv_filename_tab2", "extracted_data_tab2"))
                    self.csv_filename_tab3.set(s.get("csv_filename_tab3", "page_extract"))
                    self.page_ranges.set(s.get("page_ranges", ""))
            except Exception as e:
                print(f"Error loading settings: {e}")
                self.criteria_tab1 = defaults_tab1.copy()
                self.criteria_tab2 = defaults_tab2.copy()
        else:
            self.criteria_tab1 = defaults_tab1.copy()
            self.criteria_tab2 = defaults_tab2.copy()
            self.tab1_naming_selections = [0]
            self.tab2_naming_selections = [0]
            self.tab1_name = "Documents with IDs"
            self.tab2_name = "Documents without IDs"
    
    def save_settings(self):
        settings = {
            "tab1_name": self.tab1_name,
            "tab2_name": self.tab2_name,
            "tab1_button_override": getattr(self, 'tab1_button_override', False),
            "tab2_button_override": getattr(self, 'tab2_button_override', False),
            "tab1_button_text": self.tab1_button_text.get(),
            "tab2_button_text": self.tab2_button_text.get(),
            "tab1_naming_selections": self.tab1_naming_selections,
            "tab2_naming_selections": self.tab2_naming_selections,
            "naming_separator": self.naming_separator.get(),
            "filename_suffix": self.filename_suffix.get(),
            "criteria_tab1": self.criteria_tab1,
            "criteria_tab2": self.criteria_tab2,
            "csv_output_folder": self.csv_output_folder.get(),
            "csv_filename": self.csv_filename.get(),
            "csv_filename_tab2": self.csv_filename_tab2.get(),
            "csv_filename_tab3": self.csv_filename_tab3.get(),
            "page_ranges": self.page_ranges.get(),
            "show_tip_tab1": self.show_tip_tab1.get(),
            "show_tip_tab2": self.show_tip_tab2.get()
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except:
            pass
    
    # ==================== SYNC FUNCTIONS ====================
    def sync_tab1_button_to_name(self):
        if not hasattr(self, 'tab1_button_override') or not self.tab1_button_override:
            default_button = f"SPLIT {self.tab1_name.upper()}"
            self.tab1_button_text.set(default_button)
            if hasattr(self, 'process_btn'):
                self.process_btn.config(text=self.tab1_button_text.get())
    
    def sync_tab2_button_to_name(self):
        if not hasattr(self, 'tab2_button_override') or not self.tab2_button_override:
            default_button = f"SPLIT {self.tab2_name.upper()}"
            self.tab2_button_text.set(default_button)
            if hasattr(self, 'process_btn2'):
                self.process_btn2.config(text=self.tab2_button_text.get())
    
    def sync_buttons_to_tab_names(self):
        self.tab1_button_override = False
        self.tab2_button_override = False
        self.sync_tab1_button_to_name()
        self.sync_tab2_button_to_name()
        self.save_settings()
        messagebox.showinfo("Success", "Buttons synced to tab names!")
    
    def edit_tab1_button(self):
        new_text = simpledialog.askstring("Edit Tab 1 Button", 
            f"Current: {self.tab1_button_text.get()}\n\nEnter new text (leave empty to sync with tab name):",
            initialvalue=self.tab1_button_text.get())
        if new_text is not None:
            if new_text.strip():
                self.tab1_button_text.set(new_text.strip().upper())
                self.tab1_button_override = True
            else:
                self.tab1_button_override = False
                self.sync_tab1_button_to_name()
            if hasattr(self, 'process_btn'):
                self.process_btn.config(text=self.tab1_button_text.get())
            self.save_settings()
    
    def edit_tab2_button(self):
        new_text = simpledialog.askstring("Edit Tab 2 Button",
            f"Current: {self.tab2_button_text.get()}\n\nEnter new text (leave empty to sync with tab name):",
            initialvalue=self.tab2_button_text.get())
        if new_text is not None:
            if new_text.strip():
                self.tab2_button_text.set(new_text.strip().upper())
                self.tab2_button_override = True
            else:
                self.tab2_button_override = False
                self.sync_tab2_button_to_name()
            if hasattr(self, 'process_btn2'):
                self.process_btn2.config(text=self.tab2_button_text.get())
            self.save_settings()
    
    def rename_tab1(self):
        new = simpledialog.askstring("Rename Tab 1", "Enter new name:", initialvalue=self.tab1_name)
        if new and new.strip():
            self.tab1_name = new.strip()
            self.notebook.tab(self.tab1, text=f"📄 {self.tab1_name}")
            self.sync_tab1_button_to_name()
            self.save_settings()
    
    def rename_tab2(self):
        new = simpledialog.askstring("Rename Tab 2", "Enter new name:", initialvalue=self.tab2_name)
        if new and new.strip():
            self.tab2_name = new.strip()
            self.notebook.tab(self.tab2, text=f"🔍 {self.tab2_name}")
            self.sync_tab2_button_to_name()
            self.save_settings()
    
    def reset_tab_names(self):
        if messagebox.askyesno("Reset Names", "Reset tab names to defaults?"):
            self.tab1_name = "Documents with IDs"
            self.tab2_name = "Documents without IDs"
            self.notebook.tab(self.tab1, text=f"📄 {self.tab1_name}")
            self.notebook.tab(self.tab2, text=f"🔍 {self.tab2_name}")
            self.sync_tab1_button_to_name()
            self.sync_tab2_button_to_name()
            self.save_settings()
    
    # ==================== CRITERIA MANAGEMENT ====================
    def add_criterion_dialog(self, tab=1):
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Add New Criterion - Tab {tab}")
        dialog.geometry("600x500")
        dialog.resizable(False, False)
        dialog.config(cursor="", bg=COLORS['bg'])
        
        try:
            dialog.iconbitmap('docsplit.ico')
        except:
            pass
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"+{x}+{y}")
        
        main = tk.Frame(dialog, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=25, pady=25)
        
        criteria_list = self.criteria_tab1 if tab == 1 else self.criteria_tab2
        tk.Label(main, text=f"Add New Extraction Criterion (Tab {tab})", font=('Segoe UI', 14, 'bold'), 
                bg=COLORS['card'], fg=COLORS['text']).pack(pady=(0, 20))
        
        tk.Label(main, text="Criterion Display Name:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        name_entry = ttk.Entry(main, font=('Segoe UI', 10), width=50)
        name_entry.pack(fill='x', pady=(0, 15))
        name_entry.insert(0, f"Criteria {len(criteria_list) + 1}")
        
        tk.Label(main, text="Text Prefix (appears BEFORE the value):", anchor='w', bg=COLORS['card'],
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        prefix_entry = ttk.Entry(main, font=('Segoe UI', 10), width=50)
        prefix_entry.pack(fill='x', pady=(0, 15))
        prefix_entry.insert(0, "Field name:")
        
        # Stop Text with checkbox
        stop_frame = tk.Frame(main, bg=COLORS['card'])
        stop_frame.pack(fill='x', pady=(0, 15))
        
        use_stop_var = tk.BooleanVar(value=False)
        stop_check = tk.Checkbutton(stop_frame, text="Enable Stop Text (stop reading when this text appears)", 
                                    variable=use_stop_var, bg=COLORS['card'], fg=COLORS['text'],
                                    activebackground=COLORS['card'], selectcolor=COLORS['card'])
        stop_check.pack(anchor='w')
        
        stop_entry_frame = tk.Frame(stop_frame, bg=COLORS['card'])
        stop_entry_frame.pack(fill='x', pady=(5, 0))
        
        tk.Label(stop_entry_frame, text="Stop Text:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        stop_entry = ttk.Entry(stop_entry_frame, font=('Segoe UI', 10), width=40)
        stop_entry.pack(side='left', fill='x', expand=True, padx=(5, 0))
        stop_entry.config(state='disabled')
        
        tk.Label(stop_frame, text="Example: Enter 'on the day of' to stop reading at that line (keeps multi-line values together)", 
                bg=COLORS['card'], fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(anchor='w', pady=(5, 0))
        
        def toggle_stop_entry():
            if use_stop_var.get():
                stop_entry.config(state='normal')
            else:
                stop_entry.config(state='disabled')
                stop_entry.delete(0, tk.END)
        
        stop_check.config(command=toggle_stop_entry)
        
        tk.Label(main, text="Data Type:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        type_var = tk.StringVar(value="text")
        type_frame = tk.Frame(main, bg=COLORS['card'])
        type_frame.pack(fill='x', pady=(0, 20))
        tk.Radiobutton(type_frame, text="Text (letters, spaces, punctuation)", variable=type_var, 
                    value="text", bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        tk.Radiobutton(type_frame, text="ID (numbers only - for grouping pages)", variable=type_var, 
                    value="id", bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        
        btn_frame = tk.Frame(main, bg=COLORS['card'])
        btn_frame.pack(fill='x', pady=(10, 0))
        
        def save():
            name = name_entry.get().strip()
            prefix = prefix_entry.get().strip()
            stop_text = stop_entry.get().strip() if use_stop_var.get() else ""
            data_type = type_var.get()
            if not name or not prefix:
                messagebox.showerror("Error", "Please fill in all fields")
                return
            criterion = {"name": name, "prefix": prefix, "stop_text": stop_text, "type": data_type, "active": True}
            if tab == 1:
                self.criteria_tab1.append(criterion)
                self.update_criteria_display_tab1()
            else:
                self.criteria_tab2.append(criterion)
                self.update_criteria_display_tab2()
            self.save_settings()
            messagebox.showinfo("Success", f"Added criterion: {name}")
            dialog.destroy()
        
        tk.Button(btn_frame, text="Save", command=save, bg=COLORS['success'], fg='white', 
                font=('Segoe UI', 10), width=15, relief='flat').pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy, bg=COLORS['button_bg'], 
                fg=COLORS['button_fg'], font=('Segoe UI', 10), width=15, relief='flat').pack(side='left', padx=5)
    
    def remove_last_criterion(self, tab=1):
        if tab == 1 and self.criteria_tab1:
            removed = self.criteria_tab1.pop()
            self.update_criteria_display_tab1()
            self.save_settings()
            messagebox.showinfo("Success", f"Removed: {removed['name']}")
        elif tab == 2 and self.criteria_tab2:
            removed = self.criteria_tab2.pop()
            self.update_criteria_display_tab2()
            self.save_settings()
            messagebox.showinfo("Success", f"Removed: {removed['name']}")
        else:
            messagebox.showwarning("Warning", f"No criteria to remove in Tab {tab}")
    
    def reset_criteria(self, tab=1):
        if messagebox.askyesno("Reset Criteria", f"Reset Tab {tab} criteria to defaults?"):
            if tab == 1:
                self.criteria_tab1 = [
                    {"name": "Document ID", "prefix": "ID Number:", "type": "id", "stop_text": ""},
                    {"name": "Name", "prefix": "Name:", "type": "text", "stop_text": ""},
                    {"name": "Description", "prefix": "Description:", "type": "text", "stop_text": ""},
                    {"name": "Date", "prefix": "Date:", "type": "text", "stop_text": ""}
                ]
                self.tab1_naming_selections = [0]
                self.update_criteria_display_tab1()
            else:
                self.criteria_tab2 = [
                    {"name": "Document ID", "prefix": "ID:", "type": "id", "stop_text": ""},
                    {"name": "Name", "prefix": "Name:", "type": "text", "stop_text": ""},
                    {"name": "Description", "prefix": "Description:", "type": "text", "stop_text": ""},
                    {"name": "Date", "prefix": "Date:", "type": "text", "stop_text": ""}
                ]
                self.tab2_naming_selections = [0]
                self.update_criteria_display_tab2()
            self.save_settings()
            messagebox.showinfo("Success", f"Tab {tab} criteria reset to defaults!")
    
    def update_criteria_display_tab1(self):
        """Update Tab 1 criteria display"""
        if not hasattr(self, 'criteria_container_tab1'):
            return
        
        for widget in self.criteria_container_tab1.winfo_children():
            widget.destroy()
        
        if not self.criteria_tab1:
            tk.Label(self.criteria_container_tab1, text="No criteria defined. Click 'Add Criterion' to begin.",
                    bg=COLORS['card'], fg=COLORS['warning']).pack(pady=30)
            return
        
        header = tk.Frame(self.criteria_container_tab1, bg=COLORS['header_bg'])
        header.pack(fill='x', pady=(0, 2))
        
        tk.Label(header, text="Criterion Name", width=22, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8, pady=8)
        tk.Label(header, text="Text Prefix", width=39, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="Type", width=10, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="Include in\nFilename", width=14, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="", width=12, anchor='w', bg=COLORS['header_bg']).pack(side='left')
        
        tk.Frame(self.criteria_container_tab1, height=1, bg=COLORS['border']).pack(fill='x', pady=5)
        
        for idx, crit in enumerate(self.criteria_tab1):
            row = tk.Frame(self.criteria_container_tab1, bg=COLORS['card'])
            row.pack(fill='x', pady=2)
            
            name_label = tk.Label(row, text=crit["name"], width=25, anchor='w', 
                                 bg=COLORS['log_bg'], fg=COLORS['text'], relief='flat')
            name_label.pack(side='left', padx=5, pady=5)
            
            prefix_var = tk.StringVar(value=crit["prefix"])
            prefix_entry = ttk.Entry(row, textvariable=prefix_var, width=48)
            prefix_entry.pack(side='left', padx=5, pady=5)
            prefix_entry.bind('<FocusOut>', lambda e, i=idx, v=prefix_var: self.update_criterion_prefix_tab1(i, v.get()))
            
            type_text = "ID" if crit["type"] == "id" else "Text"
            type_label = tk.Label(row, text=type_text, width=10, anchor='w', 
                                 bg=COLORS['card'], fg=COLORS['text'])
            type_label.pack(side='left', padx=5, pady=5)
            
            naming_var = tk.BooleanVar(value=(idx in self.tab1_naming_selections))
            naming_cb = tk.Checkbutton(row, variable=naming_var, bg=COLORS['card'], 
                                      fg=COLORS['text'], activebackground=COLORS['card'], 
                                      selectcolor=COLORS['card'],
                                      command=lambda i=idx, v=naming_var: self.toggle_tab1_naming(i, v))
            naming_cb.pack(side='left', padx=5, pady=5)
            
            btn_frame = tk.Frame(row, bg=COLORS['card'])
            btn_frame.pack(side='left', padx=5, pady=5)
            tk.Button(btn_frame, text="Edit", bg=COLORS['warning'], fg='white', relief='flat', width=6,
                     command=lambda i=idx, n=crit["name"], p=crit["prefix"], t=crit["type"]: 
                     self.edit_criterion_tab1(i, n, p, t)).pack(side='left', padx=1)
            tk.Button(btn_frame, text="Delete", bg=COLORS['error'], fg='white', relief='flat', width=6,
                     command=lambda i=idx: self.delete_criterion_tab1(i)).pack(side='left', padx=1)
        
        control_frame = tk.Frame(self.criteria_container_tab1, bg=COLORS['card'])
        control_frame.pack(fill='x', pady=15)
        
        tk.Label(control_frame, text="Separator:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=5)
        sep_combo = ttk.Combobox(control_frame, textvariable=self.naming_separator,
                                values=["_", "-", " ", ".", "__"], width=8, state="readonly")
        sep_combo.pack(side='left', padx=5)
        
        tk.Label(control_frame, text="Suffix:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=(15, 5))
        suffix_entry = ttk.Entry(control_frame, textvariable=self.filename_suffix, width=15)
        suffix_entry.pack(side='left', padx=5)
        tk.Label(control_frame, text="(appended with separator)", bg=COLORS['card'], 
                fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(side='left', padx=2)
        
        tk.Label(control_frame, text="Preview:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=(15, 5))
        self.preview_label = tk.Label(control_frame, text="", bg=COLORS['card'], fg=COLORS['success'], font=('Segoe UI', 9, 'bold'))
        self.preview_label.pack(side='left')
        
        action_frame = tk.Frame(self.criteria_container_tab1, bg=COLORS['card'])
        action_frame.pack(fill='x', pady=10)
        tk.Button(action_frame, text="+ Add Criterion", command=lambda: self.add_criterion_dialog(tab=1),
                 bg=COLORS['success'], fg='white', font=('Segoe UI', 9), relief='flat', padx=10).pack(side='left', padx=5)
        tk.Button(action_frame, text="Reset to Defaults", command=lambda: self.reset_criteria(tab=1),
                 bg=COLORS['button_bg'], fg=COLORS['button_fg'], font=('Segoe UI', 9), relief='flat', padx=10).pack(side='left', padx=5)
        
        refresh_row = tk.Frame(self.criteria_container_tab1, bg=COLORS['card'])
        refresh_row.pack(fill='x', pady=5)
        
        refresh_btn = tk.Button(refresh_row, text="⟳ Refresh Preview", command=self.refresh_preview,
                               bg=COLORS['button_bg'], fg=COLORS['button_fg'], relief='flat', padx=10)
        refresh_btn.pack(side='left', padx=5)
        
        self.status_preview = tk.Label(refresh_row, text="", bg=COLORS['card'], fg=COLORS['text_light'], font=('Segoe UI', 8))
        self.status_preview.pack(side='left', padx=5)
        
        self.update_naming_preview()
    
    def update_criteria_display_tab2(self):
        """Update Tab 2 criteria display - Independent from Tab 1"""
        if not hasattr(self, 'criteria_container_tab2'):
            return
        
        for widget in self.criteria_container_tab2.winfo_children():
            widget.destroy()
        
        if not self.criteria_tab2:
            tk.Label(self.criteria_container_tab2, text="No criteria defined. Click 'Add Criterion' to begin.",
                    bg=COLORS['card'], fg=COLORS['warning']).pack(pady=30)
            return
        
        header = tk.Frame(self.criteria_container_tab2, bg=COLORS['header_bg'])
        header.pack(fill='x', pady=(0, 2))
        
        tk.Label(header, text="Criterion Name", width=18, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8, pady=8)
        tk.Label(header, text="Text Prefix", width=25, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="Stop Text", width=20, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="Type", width=8, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="Include in\nFilename", width=12, anchor='w', font=('Segoe UI', 9, 'bold'), 
                bg=COLORS['header_bg'], fg=COLORS['text']).pack(side='left', padx=8)
        tk.Label(header, text="", width=10, anchor='w', bg=COLORS['header_bg']).pack(side='left')
        
        tk.Frame(self.criteria_container_tab2, height=1, bg=COLORS['border']).pack(fill='x', pady=5)
        
        for idx, crit in enumerate(self.criteria_tab2):
            row = tk.Frame(self.criteria_container_tab2, bg=COLORS['card'])
            row.pack(fill='x', pady=2)
            
            name_label = tk.Label(row, text=crit["name"], width=20, anchor='w', 
                                bg=COLORS['log_bg'], fg=COLORS['text'], relief='flat')
            name_label.pack(side='left', padx=5, pady=5)
            
            prefix_var = tk.StringVar(value=crit["prefix"])
            prefix_entry = ttk.Entry(row, textvariable=prefix_var, width=30)
            prefix_entry.pack(side='left', padx=5, pady=5)
            prefix_entry.bind('<FocusOut>', lambda e, i=idx, v=prefix_var: self.update_criterion_prefix_tab2(i, v.get()))
            
            stop_var = tk.StringVar(value=crit.get("stop_text", ""))
            stop_entry = ttk.Entry(row, textvariable=stop_var, width=25)
            stop_entry.pack(side='left', padx=5, pady=5)
            stop_entry.bind('<FocusOut>', lambda e, i=idx, v=stop_var: self.update_criterion_stop_text_tab2(i, v.get()))
            
            type_text = "ID" if crit["type"] == "id" else "Text"
            type_label = tk.Label(row, text=type_text, width=10, anchor='w', 
                                bg=COLORS['card'], fg=COLORS['text'])
            type_label.pack(side='left', padx=5, pady=5)
            
            naming_var = tk.BooleanVar(value=(idx in self.tab2_naming_selections))
            naming_cb = tk.Checkbutton(row, variable=naming_var, bg=COLORS['card'], 
                                    fg=COLORS['text'], activebackground=COLORS['card'], 
                                    selectcolor=COLORS['card'],
                                    command=lambda i=idx, v=naming_var: self.toggle_tab2_naming(i, v))
            naming_cb.pack(side='left', padx=5, pady=5)
            
            btn_frame = tk.Frame(row, bg=COLORS['card'])
            btn_frame.pack(side='left', padx=5, pady=5)
            tk.Button(btn_frame, text="Edit", bg=COLORS['warning'], fg='white', relief='flat', width=6,
                    command=lambda i=idx, n=crit["name"], p=crit["prefix"], s=crit.get("stop_text",""), t=crit["type"]: 
                    self.edit_criterion_tab2(i, n, p, t, s)).pack(side='left', padx=1)
            tk.Button(btn_frame, text="Delete", bg=COLORS['error'], fg='white', relief='flat', width=6,
                    command=lambda i=idx: self.delete_criterion_tab2(i)).pack(side='left', padx=1)
        
        control_frame = tk.Frame(self.criteria_container_tab2, bg=COLORS['card'])
        control_frame.pack(fill='x', pady=15)
        
        tk.Label(control_frame, text="Separator:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=5)
        sep_combo = ttk.Combobox(control_frame, textvariable=self.naming_separator,
                                values=["_", "-", " ", ".", "__"], width=8, state="readonly")
        sep_combo.pack(side='left', padx=5)
        
        tk.Label(control_frame, text="Suffix:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=(15, 5))
        suffix_entry = ttk.Entry(control_frame, textvariable=self.filename_suffix, width=15)
        suffix_entry.pack(side='left', padx=5)
        
        tk.Label(control_frame, text="Preview:", bg=COLORS['card'], fg=COLORS['text_light']).pack(side='left', padx=(15, 5))
        self.preview_label_tab2 = tk.Label(control_frame, text="", bg=COLORS['card'], fg=COLORS['success'], font=('Segoe UI', 9, 'bold'))
        self.preview_label_tab2.pack(side='left')
        
        action_frame = tk.Frame(self.criteria_container_tab2, bg=COLORS['card'])
        action_frame.pack(fill='x', pady=10)
        tk.Button(action_frame, text="+ Add Criterion", command=lambda: self.add_criterion_dialog(tab=2),
                bg=COLORS['success'], fg='white', font=('Segoe UI', 9), relief='flat', padx=10).pack(side='left', padx=5)
        tk.Button(action_frame, text="Reset to Defaults", command=lambda: self.reset_criteria(tab=2),
                bg=COLORS['button_bg'], fg=COLORS['button_fg'], font=('Segoe UI', 9), relief='flat', padx=10).pack(side='left', padx=5)
        
        refresh_row = tk.Frame(self.criteria_container_tab2, bg=COLORS['card'])
        refresh_row.pack(fill='x', pady=5)
        
        refresh_btn2 = tk.Button(refresh_row, text="⟳ Refresh Preview", command=self.refresh_preview,
                                bg=COLORS['button_bg'], fg=COLORS['button_fg'], relief='flat', padx=10)
        refresh_btn2.pack(side='left', padx=5)
        
        self.update_naming_preview_tab2()
    
    def update_criterion_stop_text_tab2(self, idx, new_stop_text):
        if 0 <= idx < len(self.criteria_tab2):
            self.criteria_tab2[idx]["stop_text"] = new_stop_text
            self.save_settings()

    def edit_criterion_tab1(self, idx, name, prefix, type_val):
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Criterion - Tab 1")
        dialog.geometry("600x500")
        dialog.resizable(False, False)
        dialog.configure(bg=COLORS['bg'])
        
        try:
            dialog.iconbitmap('docsplit.ico')
        except:
            pass
        
        main = tk.Frame(dialog, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=25, pady=25)
        
        # Get existing stop text
        existing_stop_text = self.criteria_tab1[idx].get("stop_text", "")
        has_stop_text = bool(existing_stop_text)
        
        tk.Label(main, text="Edit Criterion (Tab 1)", font=('Segoe UI', 14, 'bold'), 
                bg=COLORS['card'], fg=COLORS['text']).pack(pady=(0, 20))
        
        tk.Label(main, text="Criterion Name:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        name_entry = ttk.Entry(main, font=('Segoe UI', 10), width=40)
        name_entry.pack(fill='x', pady=(0, 15))
        name_entry.insert(0, name)
        
        tk.Label(main, text="Text Prefix:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        prefix_entry = ttk.Entry(main, font=('Segoe UI', 10), width=40)
        prefix_entry.pack(fill='x', pady=(0, 15))
        prefix_entry.insert(0, prefix)
        
        # Stop Text with checkbox
        stop_frame = tk.Frame(main, bg=COLORS['card'])
        stop_frame.pack(fill='x', pady=(0, 15))
        
        use_stop_var = tk.BooleanVar(value=has_stop_text)
        stop_check = tk.Checkbutton(stop_frame, text="Enable Stop Text (stop reading when this text appears)", 
                                    variable=use_stop_var, bg=COLORS['card'], fg=COLORS['text'],
                                    activebackground=COLORS['card'], selectcolor=COLORS['card'])
        stop_check.pack(anchor='w')
        
        stop_entry_frame = tk.Frame(stop_frame, bg=COLORS['card'])
        stop_entry_frame.pack(fill='x', pady=(5, 0))
        
        tk.Label(stop_entry_frame, text="Stop Text:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        stop_entry = ttk.Entry(stop_entry_frame, font=('Segoe UI', 10), width=40)
        stop_entry.pack(side='left', fill='x', expand=True, padx=(5, 0))
        stop_entry.insert(0, existing_stop_text)
        if not has_stop_text:
            stop_entry.config(state='disabled')
        
        tk.Label(stop_frame, text="Example: Enter 'on the day of' to stop reading at that line (keeps multi-line values together)", 
                bg=COLORS['card'], fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(anchor='w', pady=(5, 0))
        
        def toggle_stop_entry():
            if use_stop_var.get():
                stop_entry.config(state='normal')
            else:
                stop_entry.config(state='disabled')
                stop_entry.delete(0, tk.END)
        
        stop_check.config(command=toggle_stop_entry)
        
        tk.Label(main, text="Data Type:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        type_var = tk.StringVar(value=type_val)
        tk.Radiobutton(main, text="Text", variable=type_var, value="text", 
                    bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        tk.Radiobutton(main, text="ID (numbers only)", variable=type_var, value="id", 
                    bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        
        btn_frame = tk.Frame(main, bg=COLORS['card'])
        btn_frame.pack(fill='x', pady=(20, 0))
        
        def save():
            self.criteria_tab1[idx]["name"] = name_entry.get().strip()
            self.criteria_tab1[idx]["prefix"] = prefix_entry.get().strip()
            self.criteria_tab1[idx]["stop_text"] = stop_entry.get().strip() if use_stop_var.get() else ""
            self.criteria_tab1[idx]["type"] = type_var.get()
            self.update_criteria_display_tab1()
            self.save_settings()
            dialog.destroy()
        
        tk.Button(btn_frame, text="Save", command=save, bg=COLORS['success'], fg='white', 
                relief='flat', width=15).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy, bg=COLORS['button_bg'], 
                fg=COLORS['button_fg'], relief='flat', width=15).pack(side='left', padx=5)
    
    def edit_criterion_tab2(self, idx, name, prefix, type_val, stop_text=""):
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Criterion - Tab 2")
        dialog.geometry("600x500")
        dialog.resizable(False, False)
        dialog.configure(bg=COLORS['bg'])
        
        try:
            dialog.iconbitmap('docsplit.ico')
        except:
            pass
        
        main = tk.Frame(dialog, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=25, pady=25)
        
        has_stop_text = bool(stop_text)
        
        tk.Label(main, text="Edit Criterion (Tab 2)", font=('Segoe UI', 14, 'bold'), 
                bg=COLORS['card'], fg=COLORS['text']).pack(pady=(0, 20))
        
        tk.Label(main, text="Criterion Name:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        name_entry = ttk.Entry(main, font=('Segoe UI', 10), width=40)
        name_entry.pack(fill='x', pady=(0, 15))
        name_entry.insert(0, name)
        
        tk.Label(main, text="Text Prefix:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        prefix_entry = ttk.Entry(main, font=('Segoe UI', 10), width=40)
        prefix_entry.pack(fill='x', pady=(0, 15))
        prefix_entry.insert(0, prefix)
        
        # Stop Text with checkbox
        stop_frame = tk.Frame(main, bg=COLORS['card'])
        stop_frame.pack(fill='x', pady=(0, 15))
        
        use_stop_var = tk.BooleanVar(value=has_stop_text)
        stop_check = tk.Checkbutton(stop_frame, text="Enable Stop Text (stop reading when this text appears)", 
                                    variable=use_stop_var, bg=COLORS['card'], fg=COLORS['text'],
                                    activebackground=COLORS['card'], selectcolor=COLORS['card'])
        stop_check.pack(anchor='w')
        
        stop_entry_frame = tk.Frame(stop_frame, bg=COLORS['card'])
        stop_entry_frame.pack(fill='x', pady=(5, 0))
        
        tk.Label(stop_entry_frame, text="Stop Text:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        stop_entry = ttk.Entry(stop_entry_frame, font=('Segoe UI', 10), width=40)
        stop_entry.pack(side='left', fill='x', expand=True, padx=(5, 0))
        stop_entry.insert(0, stop_text)
        if not has_stop_text:
            stop_entry.config(state='disabled')
        
        tk.Label(stop_frame, text="Example: Enter 'on the day of' to stop reading at that line (keeps multi-line values together)", 
                bg=COLORS['card'], fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(anchor='w', pady=(5, 0))
        
        def toggle_stop_entry():
            if use_stop_var.get():
                stop_entry.config(state='normal')
            else:
                stop_entry.config(state='disabled')
                stop_entry.delete(0, tk.END)
        
        stop_check.config(command=toggle_stop_entry)
        
        tk.Label(main, text="Data Type:", anchor='w', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(fill='x', pady=(0, 5))
        type_var = tk.StringVar(value=type_val)
        tk.Radiobutton(main, text="Text", variable=type_var, value="text", 
                    bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        tk.Radiobutton(main, text="ID (numbers only)", variable=type_var, value="id", 
                    bg=COLORS['card'], fg=COLORS['text'], selectcolor=COLORS['card']).pack(anchor='w')
        
        btn_frame = tk.Frame(main, bg=COLORS['card'])
        btn_frame.pack(fill='x', pady=(20, 0))
        
        def save():
            self.criteria_tab2[idx]["name"] = name_entry.get().strip()
            self.criteria_tab2[idx]["prefix"] = prefix_entry.get().strip()
            self.criteria_tab2[idx]["stop_text"] = stop_entry.get().strip() if use_stop_var.get() else ""
            self.criteria_tab2[idx]["type"] = type_var.get()
            self.update_criteria_display_tab2()
            self.save_settings()
            dialog.destroy()
        
        tk.Button(btn_frame, text="Save", command=save, bg=COLORS['success'], fg='white', 
                relief='flat', width=15).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy, bg=COLORS['button_bg'], 
                fg=COLORS['button_fg'], relief='flat', width=15).pack(side='left', padx=5)
    
    def delete_criterion_tab1(self, idx):
        if messagebox.askyesno("Confirm Delete", f"Delete '{self.criteria_tab1[idx]['name']}'?"):
            del self.criteria_tab1[idx]
            self.update_criteria_display_tab1()
            self.save_settings()
    
    def delete_criterion_tab2(self, idx):
        if messagebox.askyesno("Confirm Delete", f"Delete '{self.criteria_tab2[idx]['name']}'?"):
            del self.criteria_tab2[idx]
            self.update_criteria_display_tab2()
            self.save_settings()
    
    def toggle_tab1_naming(self, idx, var):
        if var.get():
            if idx not in self.tab1_naming_selections and len(self.tab1_naming_selections) < 4:
                self.tab1_naming_selections.append(idx)
            elif len(self.tab1_naming_selections) >= 4:
                var.set(False)
                messagebox.showwarning("Limit Reached", "You can select up to 4 criteria for file naming")
        else:
            if idx in self.tab1_naming_selections:
                self.tab1_naming_selections.remove(idx)
        self.tab1_naming_selections.sort()
        self.save_settings()
        self.update_naming_preview()
    
    def toggle_tab2_naming(self, idx, var):
        if var.get():
            if idx not in self.tab2_naming_selections and len(self.tab2_naming_selections) < 4:
                self.tab2_naming_selections.append(idx)
            elif len(self.tab2_naming_selections) >= 4:
                var.set(False)
                messagebox.showwarning("Limit Reached", "You can select up to 4 criteria for file naming")
        else:
            if idx in self.tab2_naming_selections:
                self.tab2_naming_selections.remove(idx)
        self.tab2_naming_selections.sort()
        self.save_settings()
        self.update_naming_preview_tab2()
    
    def update_naming_preview(self):
        if hasattr(self, 'preview_label'):
            if self.tab1_naming_selections:
                parts = [self.criteria_tab1[i]["name"] for i in self.tab1_naming_selections if i < len(self.criteria_tab1)]
                sep = self.naming_separator.get()
                suffix = self.filename_suffix.get().strip()
                if suffix:
                    preview = f"{sep.join(parts)}{sep}{suffix}.pdf"
                else:
                    preview = f"{sep.join(parts)}.pdf"
                self.preview_label.config(text=preview)
            else:
                self.preview_label.config(text="Select criteria")
    
    def update_naming_preview_tab2(self):
        if hasattr(self, 'preview_label_tab2'):
            if self.tab2_naming_selections:
                parts = [self.criteria_tab2[i]["name"] for i in self.tab2_naming_selections if i < len(self.criteria_tab2)]
                sep = self.naming_separator.get()
                suffix = self.filename_suffix.get().strip()
                if suffix:
                    preview = f"{sep.join(parts)}{sep}{suffix}.pdf"
                else:
                    preview = f"{sep.join(parts)}.pdf"
                self.preview_label_tab2.config(text=preview)
            else:
                self.preview_label_tab2.config(text="Select criteria")
    
    def update_criterion_prefix_tab1(self, idx, new_prefix):
        if 0 <= idx < len(self.criteria_tab1):
            self.criteria_tab1[idx]["prefix"] = new_prefix
            self.save_settings()
            self.update_naming_preview()
    
    def update_criterion_prefix_tab2(self, idx, new_prefix):
        if 0 <= idx < len(self.criteria_tab2):
            self.criteria_tab2[idx]["prefix"] = new_prefix
            self.save_settings()
            self.update_naming_preview_tab2()
    
    def build_regex_for_criterion(self, criterion):
        prefix = re.escape(criterion["prefix"].strip())
        # Use DOTALL to match across multiple lines
        return prefix + r'\s*[:]?\s*(.+?)(?=\n\s*(?:' + re.escape(criterion.get("stop_text", "")) + r')|\Z)'
    
    def build_filename(self, values, indices, criteria_list, separator, suffix, for_tab2=False):
        parts = []
        for idx in indices:
            if idx < len(criteria_list):
                name = criteria_list[idx]["name"]
                val = values.get(name, "unknown")
                if val == "Not Found" or val == "Unknown":
                    val = "unknown"
                clean = re.sub(r'[^\w\s-]', '', str(val)).strip()
                clean = re.sub(r'\s+', '_', clean) if clean else "unknown"
                parts.append(clean)
        
        if for_tab2:
            parts = [part for part in parts if not (part.isdigit() and len(part) < 5)]
        
        base = separator.join(parts) if parts else "unknown"
        
        if suffix:
            return f"{base}{separator}{suffix}.pdf"
        else:
            return f"{base}.pdf"
    
    # ==================== HELP ====================
    def show_quick_start(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Quick Start Guide")
        help_window.geometry("700x650")
        help_window.resizable(False, False)
        help_window.configure(bg=COLORS['bg'])
        
        try:
            help_window.iconbitmap('docsplit.ico')
        except:
            pass
        
        text = tk.Text(help_window, wrap=tk.WORD, padx=15, pady=15, font=('Segoe UI', 10), 
                      bg=COLORS['card'], fg=COLORS['text'], relief='flat')
        text.pack(fill='both', expand=True, padx=10, pady=10)
        
        content = """
DOCU SPLIT - QUICK START GUIDE

================================================================================

TAB 1: SPLIT BY ID
- For documents that contain ID numbers
- Groups all pages with the same ID into one PDF
- Extracts data from each record into CSV
- Configure your own extraction criteria for this tab

TAB 2: SPLIT BY NAME
- For documents without IDs
- Uses CSV from Tab 1 to match names to IDs
- Configure your own extraction criteria for this tab
- The Name field from Tab 2 is used to match against the CSV

TAB 3: SPLIT BY PAGE RANGE
- Extract specific page ranges from a PDF
- Supports multiple ranges: 1-5, 10-15, 20-25
- Maximum 20 ranges allowed

================================================================================

SETTING UP CRITERIA:

Tab 1 (Documents with IDs):
- Click "Criteria" menu → "Tab 1: Add Criterion"
- Enter the text prefix exactly as it appears in your PDF
- Example: If PDF shows "ID Number: 12345", set prefix to "ID Number:"
- The first criterion should be an ID type for grouping
- Optionally enable "Stop Text" to stop reading at a specific phrase

Tab 2 (Documents without IDs):
- Click "Criteria" menu → "Tab 2: Add Criterion"
- Enter prefixes that match your document
- Enable "Stop Text" to stop reading at a specific phrase for multi-line values
- The second criterion (typically Name) is used for matching to the CSV

================================================================================

TIPS:
- Maximum 4 criteria for filename in Tab 1 & 2
- "Not Found" values become "unknown"
- All settings are saved automatically
- 100% offline - no data leaves your computer
"""
        text.insert('1.0', content)
        text.config(state=tk.DISABLED)
    
    def show_about(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("About Docu Split")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        dialog.configure(bg=COLORS['card'])
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Content
        main = tk.Frame(dialog, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=20, pady=20)
        
        tk.Label(main, text="Docu Split", font=('Segoe UI', 18, 'bold'), 
                bg=COLORS['card'], fg=COLORS['text']).pack(pady=(0, 5))
        
        tk.Label(main, text="Version 4.0", font=('Segoe UI', 10), 
                bg=COLORS['card'], fg=COLORS['text_light']).pack(pady=(0, 15))
        
        tk.Label(main, text="A professional tool for splitting PDF documents\nand extracting data.", 
                justify='center', bg=COLORS['card'], fg=COLORS['text']).pack(pady=(0, 15))
        
        features = """Features:
    • Split PDF files by ID (groups same ID pages)
    • Split PDF files by name matching
    • Split PDF files by custom page ranges
    • Extract unlimited custom fields
    • Multi-criteria file naming (up to 4 fields)
    • 100% offline
    • Dark mode support"""
        
        tk.Label(main, text=features, justify='left', bg=COLORS['card'], 
                fg=COLORS['text'], font=('Segoe UI', 9)).pack(pady=(0, 15))
        
        # Donation button in About dialog
        def donate():
            import webbrowser
            webbrowser.open("https://ko-fi.com/novabananana")
            dialog.destroy()
        
        donation_btn = tk.Button(main, text="☕ Support the Project", command=donate,
                                bg=COLORS['accent'], fg='white', font=('Segoe UI', 10, 'bold'),
                                relief='flat', padx=20, pady=5, cursor="hand2")
        donation_btn.pack(pady=(0, 10))
        
        tk.Button(main, text="Close", command=dialog.destroy,
                bg=COLORS['button_bg'], fg=COLORS['button_fg'], 
                relief='flat', padx=15, pady=3).pack()
    
    def open_donation_page(self):
        """Open Ko-fi donation page"""
        import webbrowser
        webbrowser.open("https://ko-fi.com/novabananana")

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def browse_csv_folder(self):
        f = filedialog.askdirectory(title="Select folder for CSV output")
        if f:
            self.csv_output_folder.set(f)
    
    def refresh_preview(self):
        self.update_naming_preview()
        self.update_naming_preview_tab2()
        if hasattr(self, 'status_preview'):
            self.status_preview.config(text="Preview updated!", foreground=COLORS['success'])
            self.root.after(2000, lambda: self.status_preview.config(text="", foreground=COLORS['text_light']))
    
    def refresh_page_preview(self):
        if hasattr(self, 'page_preview_label'):
            ranges = self.get_page_ranges()
            if ranges:
                total_pages = sum(end - start + 1 for start, end in ranges)
                preview = f"Total pages to extract: {total_pages} | Ranges: "
                range_strs = [f"{s}-{e}" for s, e in ranges]
                preview += ", ".join(range_strs)
                self.page_preview_label.config(text=preview, foreground=COLORS['success'])
            else:
                self.page_preview_label.config(text="No valid ranges entered. Use format: 1-5, 10-15, 20-25", foreground=COLORS['warning'])
    
    def get_page_ranges(self):
        range_text = self.page_ranges.get().strip()
        if not range_text:
            return []
        
        ranges = []
        parts = range_text.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                try:
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    if start > 0 and end >= start:
                        ranges.append((start, end))
                except ValueError:
                    continue
            elif part.isdigit():
                p = int(part)
                ranges.append((p, p))
        
        if len(ranges) > 20:
            ranges = ranges[:20]
            self.root.after(0, lambda: messagebox.showwarning("Limit Reached", "Maximum 20 page ranges allowed. Extra ranges were ignored."))
        
        return ranges
    
    # ==================== HELPER METHODS ====================
    def debug_tab1_criteria(self):
        """Debug method to show Tab 1 criteria"""
        self.log(f"\n=== TAB 1 CRITERIA DEBUG ===")
        self.log(f"Number of criteria: {len(self.criteria_tab1)}")
        for i, crit in enumerate(self.criteria_tab1):
            self.log(f"  {i}: {crit['name']} - '{crit['prefix']}' - type: {crit['type']}")
        self.log(f"Naming selections: {self.tab1_naming_selections}")
    
    # ==================== TAB 1: SPLIT BY ID ====================
    def init_tab1(self):
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        
        main = tk.Frame(self.tab1, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=15, pady=15)
        
        # File selection
        file_frame = ttk.LabelFrame(main, text="📁 1. Select PDF File")
        file_frame.pack(fill='x', pady=(0, 10))
        
        file_row = tk.Frame(file_frame, bg=COLORS['card'])
        file_row.pack(fill='x', padx=10, pady=10)
        
        file_entry = ttk.Entry(file_row, textvariable=self.input_file, font=('Segoe UI', 9))
        file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(file_row, text="Browse", command=self.browse_input, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        tk.Button(file_row, text="Debug Tab 1 Criteria", command=self.debug_tab1_criteria, 
                 bg=COLORS['info'], fg='white', relief='flat', padx=10).pack(side='left', padx=5)
        
        # Output folder
        output_frame = ttk.LabelFrame(main, text="📂 2. Select Output Folder for PDFs")
        output_frame.pack(fill='x', pady=(0, 10))
        
        output_row = tk.Frame(output_frame, bg=COLORS['card'])
        output_row.pack(fill='x', padx=10, pady=10)
        
        output_entry = ttk.Entry(output_row, textvariable=self.output_folder, font=('Segoe UI', 9))
        output_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(output_row, text="Browse", command=self.browse_output, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        # CSV Settings
        csv_frame = ttk.LabelFrame(main, text="📊 3. CSV Output Settings (Optional)")
        csv_frame.pack(fill='x', pady=(0, 10))
        
        csv_folder_row = tk.Frame(csv_frame, bg=COLORS['card'])
        csv_folder_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_folder_row, text="CSV Folder:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=10, anchor='w').pack(side='left')
        tk.Entry(csv_folder_row, textvariable=self.csv_output_folder, font=('Segoe UI', 9), 
                width=60).pack(side='left', fill='x', expand=True, padx=(5, 5))
        tk.Button(csv_folder_row, text="Browse", command=self.browse_csv_folder, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        csv_name_row = tk.Frame(csv_frame, bg=COLORS['card'])
        csv_name_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_name_row, text="CSV Filename:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=10, anchor='w').pack(side='left')
        tk.Entry(csv_name_row, textvariable=self.csv_filename, font=('Segoe UI', 9), 
                width=30).pack(side='left', padx=5)
        tk.Label(csv_name_row, text=".csv will be added automatically", bg=COLORS['card'], 
                fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(side='left', padx=5)
        
        # Criteria card - This will expand with content
        criteria_frame = ttk.LabelFrame(main, text="🔧 4. Extraction Criteria & File Naming")
        criteria_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Header frame for button and tip (top area)
        self.tab1_header_frame = tk.Frame(criteria_frame, bg=COLORS['card'])
        self.tab1_header_frame.pack(fill='x', padx=10, pady=(5, 0))
        
        # Split button on the LEFT side
        self.process_btn = tk.Button(self.tab1_header_frame, text=self.tab1_button_text.get(), command=self.process_tab1,
                                     bg=COLORS['accent'], fg=COLORS['accent_fg'], font=('Segoe UI', 11, 'bold'), 
                                     relief='flat', width=28, pady=8)
        self.process_btn.pack(side='left', padx=(0, 10), pady=5)
        
        # Show/Hide Tip button (toggles tip visibility)
        tip_toggle_btn = tk.Button(self.tab1_header_frame, text="💡 Show/Hide Tip", command=self.toggle_tip_tab1,
                                   bg=COLORS['button_bg'], fg=COLORS['button_fg'], font=('Segoe UI', 9), 
                                   relief='flat', padx=8, pady=4)
        tip_toggle_btn.pack(side='left', padx=(0, 5), pady=5)
        
        # Create tip if visible
        if self.show_tip_tab1.get():
            self.create_tip_tab1()
        
        # Container for criteria that expands
        self.criteria_container_tab1 = tk.Frame(criteria_frame, bg=COLORS['card'])
        self.criteria_container_tab1.pack(fill='both', expand=True, padx=10, pady=10)
        self.update_criteria_display_tab1()
        
        # Progress bar and status
        progress_frame = tk.Frame(main, bg=COLORS['card'])
        progress_frame.pack(fill='x', pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress.pack(fill='x', pady=5)
        
        tk.Label(progress_frame, textvariable=self.status_text, bg=COLORS['card'], fg=COLORS['text_light']).pack(pady=5)

        donation_frame = tk.Frame(main, bg=COLORS['card'])
        donation_frame.pack(fill='x', pady=(5, 0))

        donate_btn = tk.Button(donation_frame, text="☕ Support the Project", 
                            command=self.open_donation_page,
                            bg=COLORS['accent'], fg='white', 
                            font=('Segoe UI', 9, 'bold'),
                            relief='flat', padx=10, pady=3, cursor="hand2")
        donate_btn.pack(side='right')
        
        # Log card
        log_frame = ttk.LabelFrame(main, text="📋 Processing Log")
        log_frame.pack(fill='both', expand=True)
        
        log_inner = tk.Frame(log_frame, bg=COLORS['card'])
        log_inner.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_inner, wrap=tk.WORD, font=('Consolas', 9), 
                               bg=COLORS['log_bg'], fg=COLORS['text'], relief='flat', borderwidth=0, height=12)
        self.log_text.pack(side='left', fill='both', expand=True)
        
        log_scrollbar = tk.Scrollbar(log_inner, command=self.log_text.yview,
                                    bg=COLORS['scrollbar_bg'], troughcolor=COLORS['scrollbar_trough'],
                                    activebackground=COLORS['scrollbar_active'], relief='flat', borderwidth=0,
                                    highlightthickness=0)
        log_scrollbar.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
    
    def browse_input(self):
        f = filedialog.askopenfilename(title="Select PDF File", 
                                       filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if f:
            self.input_file.set(f)
            if not self.output_folder.get():
                self.output_folder.set(str(Path(f).parent / "Output"))
    
    def browse_output(self):
        f = filedialog.askdirectory(title="Select output folder for PDFs")
        if f:
            self.output_folder.set(f)
    
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def process_tab1(self):
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder for PDFs")
            return
        if not self.criteria_tab1:
            messagebox.showerror("Error", "No criteria defined for Tab 1. Add criteria using the Criteria menu.")
            return
        
        id_criterion = None
        for crit in self.criteria_tab1:
            if crit["type"] == "id":
                id_criterion = crit
                break
        
        if not id_criterion:
            messagebox.showerror("Error", "At least one ID type criterion is required for grouping")
            return
        
        if not self.tab1_naming_selections:
            messagebox.showerror("Error", "Please select at least one criterion for file naming")
            return
        
        self.process_btn.config(state=tk.DISABLED)
        self.progress.start()
        self.status_text.set("Processing...")
        self.log_text.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._process_tab1_thread)
        thread.start()
    
    def _process_tab1_thread(self):
        try:
            input_path = Path(self.input_file.get())
            
            if input_path.suffix.lower() != '.pdf':
                self.root.after(0, lambda: messagebox.showerror("Error", "Only PDF files are supported."))
                return
            
            output_path = Path(self.output_folder.get())
            output_path.mkdir(parents=True, exist_ok=True)
            
            id_criterion = None
            for crit in self.criteria_tab1:
                if crit["type"] == "id":
                    id_criterion = crit
                    break
            
            id_pattern = self.build_regex_for_criterion(id_criterion)
            
            doc = fitz.open(input_path)
            page_data = []
            
            for page_num in range(len(doc)):
                text = doc[page_num].get_text()
                id_match = re.search(id_pattern, text)
                if id_match:
                    record_id = id_match.group(1)
                    page_data.append((page_num, record_id, text))
                    self.log(f"Page {page_num + 1}: Found ID {record_id}")
                else:
                    page_data.append((page_num, None, text))
                    self.log(f"Page {page_num + 1}: No ID found")
            
            doc.close()
            
            groups = []
            current_id = None
            current_pages = []
            
            for page_num, record_id, text in page_data:
                if record_id is not None:
                    if current_id is not None and record_id != current_id:
                        groups.append((current_id, current_pages))
                        current_id = record_id
                        current_pages = [page_num]
                    elif current_id is None:
                        current_id = record_id
                        current_pages = [page_num]
                    else:
                        current_pages.append(page_num)
                else:
                    if current_pages is not None:
                        current_pages.append(page_num)
                    else:
                        self.log(f"Warning: Page {page_num + 1} has no ID and appears before first record. Skipping.")
            
            if current_id is not None:
                groups.append((current_id, current_pages))
            
            if not groups:
                self.root.after(0, lambda: messagebox.showerror("Error", "No IDs found. Check your criteria."))
                return
            
            self.log(f"\nFound {len(groups)} record(s):")
            for group_idx, (record_id, pages) in enumerate(groups):
                self.log(f"  Record {group_idx + 1}: ID {record_id} - Pages {[p + 1 for p in pages]}")
            
            csv_rows = []
            naming_indices = self.tab1_naming_selections
            separator = self.naming_separator.get()
            suffix = self.filename_suffix.get().strip()
            
            reader = PyPDF2.PdfReader(str(input_path))
            
            for group_idx, (record_id, pages) in enumerate(groups):
                self.log(f"\nProcessing Record {record_id}:")
                self.log(f"  Pages in this record: {[p + 1 for p in pages]}")
                
                combined_text = ""
                doc = fitz.open(input_path)
                for page_num in pages:
                    page_text = doc[page_num].get_text()
                    combined_text += page_text + "\n"
                doc.close()
                
                self.log(f"  Searching for criteria across {len(pages)} page(s)...")
                
                # Extract ALL criteria from the text with stop text support
                page_values = {}
                for crit in self.criteria_tab1:
                    pattern = self.build_regex_for_criterion(crit)
                    match = re.search(pattern, combined_text, re.IGNORECASE | re.DOTALL)
                    if match:
                        raw_value = match.group(1).strip()
                        page_values[crit["name"]] = raw_value
                        self.log(f"  ✓ Extracted {crit['name']}: {raw_value}")
                    else:
                        page_values[crit["name"]] = "Not Found"
                        self.log(f"  ✗ {crit['name']}: Not Found")
                
                filename = self.build_filename(page_values, naming_indices, self.criteria_tab1, separator, suffix)
                output_file = output_path / filename
                
                if output_file.exists():
                    counter = 1
                    stem = Path(filename).stem
                    while output_file.exists():
                        output_file = output_path / f"{stem}_{counter}.pdf"
                        counter += 1
                    self.log(f"  Duplicate filename, saved as: {output_file.name}")
                
                writer = PyPDF2.PdfWriter()
                for page_num in pages:
                    writer.add_page(reader.pages[page_num])
                with open(output_file, 'wb') as f:
                    writer.write(f)
                
                row = {}
                for crit in self.criteria_tab1:
                    row[crit["name"]] = page_values.get(crit["name"], "Not Found")
                row["filename"] = output_file.name
                csv_rows.append(row)
                
                total_pages = len(pages)
                self.log(f"  ✓ Created: {output_file.name} ({total_pages} page{'s' if total_pages > 1 else ''})")
                self.root.after(0, lambda p=group_idx+1, t=len(groups): self.status_text.set(f"Processed {p}/{len(groups)} records"))
            
            if csv_rows:
                csv_output_path = Path(self.csv_output_folder.get()) if self.csv_output_folder.get() else output_path
                csv_output_path.mkdir(parents=True, exist_ok=True)
                
                csv_name = self.csv_filename.get().strip()
                if not csv_name:
                    csv_name = "extracted_data"
                if not csv_name.endswith('.csv'):
                    csv_name += '.csv'
                
                csv_file = csv_output_path / csv_name
                fieldnames = [crit["name"] for crit in self.criteria_tab1] + ["filename"]
                
                with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
                    writer.writeheader()
                    writer.writerows(csv_rows)
                
                self.log(f"\n[SUCCESS] COMPLETE! Processed {len(csv_rows)} records")
                self.log(f"   Output folder: {output_path}")
                self.log(f"   CSV saved: {csv_file}")
                self.log(f"   Columns: {', '.join(fieldnames)}")
                
                self.log(f"\n   Extraction Summary:")
                for crit in self.criteria_tab1:
                    found_count = sum(1 for row in csv_rows if row[crit["name"]] != "Not Found")
                    self.log(f"     {crit['name']}: {found_count}/{len(csv_rows)} found")
                
                self.root.after(0, lambda: messagebox.showinfo("Success", 
                    f"Successfully processed {len(csv_rows)} records!\n\nPDFs saved to: {output_path}\nCSV saved to: {csv_file}"))
            else:
                self.log("No data extracted")
            
        except Exception as e:
            self.log(f"\n[ERROR] {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self._processing_done)
    
    def _processing_done(self):
        self.progress.stop()
        self.process_btn.config(state=tk.NORMAL)
        self.status_text.set("Ready")
    
    # ==================== TAB 2: SPLIT BY NAME ====================
    def init_tab2(self):
        self.tab2_file = tk.StringVar()
        self.tab2_output = tk.StringVar()
        self.status_text2 = tk.StringVar(value="Ready")
        self.name_map = {}
        
        main = tk.Frame(self.tab2, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=15, pady=15)
        
        info_frame = ttk.LabelFrame(main, text="ℹ️ How It Works")
        info_frame.pack(fill='x', pady=(0, 10))
        
        info_text = """1. First process Tab 1 to create a CSV file
    2. The CSV from Tab 1 is automatically loaded as the default mapping
    3. Configure Tab 2 extraction criteria (use prefixes that match your document)
    4. Select naming rules
    5. Click 'SPLIT BY NAME' to process

    The Name field from Tab 2 will be used to match against the CSV's Name column."""
        tk.Label(info_frame, text=info_text, justify='left', bg=COLORS['card'], 
                fg=COLORS['text_light']).pack(anchor='w', padx=10, pady=10)
        
        file_frame = ttk.LabelFrame(main, text="📁 1. Select PDF File")
        file_frame.pack(fill='x', pady=(0, 10))
        
        file_row = tk.Frame(file_frame, bg=COLORS['card'])
        file_row.pack(fill='x', padx=10, pady=10)
        
        file_entry = ttk.Entry(file_row, textvariable=self.tab2_file, font=('Segoe UI', 9))
        file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(file_row, text="Browse", command=self.browse_tab2, bg=COLORS['button_bg'], 
                fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        # Add Debug PDF Text button
        tk.Button(file_row, text="Debug PDF Text", command=self.debug_pdf_text, 
                bg=COLORS['info'], fg='white', relief='flat', padx=10).pack(side='left', padx=5)
        
        output_frame = ttk.LabelFrame(main, text="📂 2. Select Output Folder for PDFs")
        output_frame.pack(fill='x', pady=(0, 10))
        
        output_row = tk.Frame(output_frame, bg=COLORS['card'])
        output_row.pack(fill='x', padx=10, pady=10)
        
        output_entry = ttk.Entry(output_row, textvariable=self.tab2_output, font=('Segoe UI', 9))
        output_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(output_row, text="Browse", command=self.browse_output_tab2, bg=COLORS['button_bg'], 
                fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        # CSV Settings for mapping
        csv_frame = ttk.LabelFrame(main, text="📊 3. CSV File for Name Mapping")
        csv_frame.pack(fill='x', pady=(0, 10))

        csv_file_row = tk.Frame(csv_frame, bg=COLORS['card'])
        csv_file_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_file_row, text="CSV File:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=10, anchor='w').pack(side='left')
        self.csv_file_path = tk.StringVar()
        csv_entry = ttk.Entry(csv_file_row, textvariable=self.csv_file_path, font=('Segoe UI', 9), width=50)
        csv_entry.pack(side='left', fill='x', expand=True, padx=(5, 5))
        tk.Button(csv_file_row, text="Browse CSV", command=self.browse_csv_file, 
                bg=COLORS['button_bg'], fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        tk.Button(csv_file_row, text="Debug CSV", command=self.debug_csv_structure, 
                bg=COLORS['warning'], fg='white', relief='flat', padx=10).pack(side='left', padx=5)
        
        tk.Button(csv_file_row, text="Refresh Mapping", command=self.refresh_mapping, 
                bg=COLORS['success'], fg='white', relief='flat', padx=10).pack(side='left', padx=5)

        # CSV output settings
        csv_output_frame = ttk.LabelFrame(main, text="📊 4. CSV Output Settings (Optional)")
        csv_output_frame.pack(fill='x', pady=(0, 10))

        csv_folder_row = tk.Frame(csv_output_frame, bg=COLORS['card'])
        csv_folder_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_folder_row, text="Output Folder:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        tk.Entry(csv_folder_row, textvariable=self.csv_output_folder, font=('Segoe UI', 9), 
                width=55).pack(side='left', fill='x', expand=True, padx=(5, 5))
        tk.Button(csv_folder_row, text="Browse", command=self.browse_csv_folder, 
                bg=COLORS['button_bg'], fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')

        csv_name_row = tk.Frame(csv_output_frame, bg=COLORS['card'])
        csv_name_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_name_row, text="CSV Filename:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        tk.Entry(csv_name_row, textvariable=self.csv_filename_tab2, font=('Segoe UI', 9), 
                width=30).pack(side='left', padx=5)
        tk.Label(csv_name_row, text=".csv will be added automatically", bg=COLORS['card'], 
                fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(side='left', padx=5)
        
        # Criteria card for Tab 2
        criteria_frame = ttk.LabelFrame(main, text="🔧 5. Extraction Criteria for Tab 2")
        criteria_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Header frame for button and tip
        self.tab2_header_frame = tk.Frame(criteria_frame, bg=COLORS['card'])
        self.tab2_header_frame.pack(fill='x', padx=10, pady=(5, 0))
        
        # Split button on the LEFT side
        self.process_btn2 = tk.Button(self.tab2_header_frame, text=self.tab2_button_text.get(), command=self.process_tab2,
                                    bg=COLORS['accent'], fg=COLORS['accent_fg'], font=('Segoe UI', 11, 'bold'), 
                                    relief='flat', width=28, pady=8)
        self.process_btn2.pack(side='left', padx=(0, 10), pady=5)
        
        # Show/Hide Tip button (toggles tip visibility)
        tip_toggle_btn = tk.Button(self.tab2_header_frame, text="💡 Show/Hide Tip", command=self.toggle_tip_tab2,
                                bg=COLORS['button_bg'], fg=COLORS['button_fg'], font=('Segoe UI', 9), 
                                relief='flat', padx=8, pady=4)
        tip_toggle_btn.pack(side='left', padx=(0, 5), pady=5)
        
        # Create tip if visible
        if self.show_tip_tab2.get():
            self.create_tip_tab2()
        
        self.criteria_container_tab2 = tk.Frame(criteria_frame, bg=COLORS['card'])
        self.criteria_container_tab2.pack(fill='both', expand=True, padx=10, pady=10)
        self.update_criteria_display_tab2()
        
        mapping_frame = ttk.LabelFrame(main, text="🔄 6. Mapping Status")
        mapping_frame.pack(fill='x', pady=(0, 10))
        
        self.mapping_label = tk.Label(mapping_frame, text="No CSV loaded. Select a CSV file.", 
                                    bg=COLORS['card'], fg=COLORS['warning'])
        self.mapping_label.pack(anchor='w', padx=10, pady=10)
        
        # Progress bar and status
        progress_frame = tk.Frame(main, bg=COLORS['card'])
        progress_frame.pack(fill='x', pady=(0, 5))
        
        self.progress2 = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress2.pack(fill='x', pady=5)
        
        tk.Label(progress_frame, textvariable=self.status_text2, bg=COLORS['card'], fg=COLORS['text_light']).pack(pady=5)
        
        # Log card (create this BEFORE auto-loading messages)
        log_frame = ttk.LabelFrame(main, text="📋 Processing Log")
        log_frame.pack(fill='both', expand=True)
        
        log_inner = tk.Frame(log_frame, bg=COLORS['card'])
        log_inner.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text2 = tk.Text(log_inner, wrap=tk.WORD, font=('Consolas', 9), 
                                bg=COLORS['log_bg'], fg=COLORS['text'], relief='flat', borderwidth=0, height=12)
        self.log_text2.pack(side='left', fill='both', expand=True)
        
        log_scrollbar2 = tk.Scrollbar(log_inner, command=self.log_text2.yview,
                                    bg=COLORS['scrollbar_bg'], troughcolor=COLORS['scrollbar_trough'],
                                    activebackground=COLORS['scrollbar_active'], relief='flat', borderwidth=0,
                                    highlightthickness=0)
        log_scrollbar2.pack(side='right', fill='y')
        self.log_text2.configure(yscrollcommand=log_scrollbar2.set)
        
        # NOW auto-load default CSV from Tab 1 (after log widget is created)
        default_csv = self.get_default_csv_path()
        if default_csv:
            self.csv_file_path.set(default_csv)
            self.load_custom_mapping(default_csv)
            self.log2(f"Auto-loaded default CSV: {Path(default_csv).name}")
        else:
            self.log2("No default CSV found. Please process Tab 1 first or browse for a CSV file.")
    
    def debug_pdf_text(self):
        """Debug method to show text from first page of PDF"""
        pdf_path = self.tab2_file.get()
        if not pdf_path:
            self.log2("No PDF file selected")
            return
        
        try:
            doc = fitz.open(pdf_path)
            self.log2(f"\n=== PDF TEXT DEBUG ===")
            self.log2(f"File: {Path(pdf_path).name}")
            self.log2(f"Total pages: {len(doc)}")
            
            for page_num in range(min(3, len(doc))):
                text = doc[page_num].get_text()
                self.log2(f"\n--- Page {page_num + 1} Text (first 1500 chars) ---")
                self.log2(text[:1500])
                self.log2(f"\n--- End of Page {page_num + 1} ---")
            
            doc.close()
            
            self.log2(f"\nYour current Tab 2 criteria prefixes:")
            for crit in self.criteria_tab2:
                self.log2(f"  '{crit['prefix']}' -> looking for '{crit['name']}'")
                
        except Exception as e:
            self.log2(f"Error debugging PDF: {e}")
    
    def refresh_mapping(self):
        """Reload the current CSV mapping"""
        if self.csv_file_path.get():
            self.load_custom_mapping(self.csv_file_path.get())
        else:
            # Try to load default CSV
            default_csv = self.get_default_csv_path()
            if default_csv:
                self.csv_file_path.set(default_csv)
                self.load_custom_mapping(default_csv)
                self.log2(f"Auto-loaded default CSV: {Path(default_csv).name}")
            else:
                self.log2("No CSV file loaded and no default CSV found")
    
    def browse_csv_file(self):
        f = filedialog.askopenfilename(
            title="Select CSV File for Name Mapping", 
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if f:
            self.csv_file_path.set(f)
            self.load_custom_mapping(f)
    
    def load_custom_mapping(self, csv_path=None):
        if not csv_path:
            csv_path = self.csv_file_path.get()
        
        if not csv_path or not Path(csv_path).exists():
            self.mapping_label.config(text="No CSV file selected", fg=COLORS['warning'])
            return
        
        try:
            self.name_map = {}
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                r = csv.DictReader(f)
                self.log2(f"\n=== LOADING CSV FILE ===")
                self.log2(f"File: {Path(csv_path).name}")
                self.log2(f"Columns found: {r.fieldnames}")
                
                rows_list = list(r)
                self.log2(f"Total rows in CSV: {len(rows_list)}")
                self.log2(f"First 3 rows of CSV:")
                for i, row in enumerate(rows_list[:3]):
                    self.log2(f"  Row {i+1}: {row}")
                
                # Find name column and ID column
                name_col = None
                id_col = None
                
                for col in r.fieldnames:
                    col_lower = col.lower().strip()
                    if not name_col and ('name' in col_lower or 'name' in col_lower):
                        name_col = col
                        self.log2(f"  Name column identified: '{col}'")
                    if not id_col and ('id' in col_lower or 'document id' in col_lower):
                        id_col = col
                        self.log2(f"  ID column identified: '{col}'")
                
                if not name_col and len(r.fieldnames) > 1:
                    name_col = r.fieldnames[1]
                    self.log2(f"Using fallback name column (2nd column): '{name_col}'")
                if not id_col and len(r.fieldnames) > 0:
                    id_col = r.fieldnames[0]
                    self.log2(f"Using fallback ID column (1st column): '{id_col}'")
                
                if not name_col or not id_col:
                    self.log2(f"ERROR: Could not identify name/ID columns")
                    self.mapping_label.config(text="Could not identify name/ID columns", fg=COLORS['error'])
                    return
                
                mapping_count = 0
                for row in rows_list:
                    name = row[name_col].strip() if name_col in row else "Unknown"
                    rid = row.get(id_col, 'NA') if id_col in row else 'NA'
                    
                    if name != "Unknown" and name and rid != 'NA':
                        clean_name = ' '.join(name.split())
                        self.name_map[clean_name] = rid
                        self.name_map[clean_name.lower()] = rid
                        mapping_count += 1
                        if mapping_count <= 5:
                            self.log2(f"  Added mapping: '{clean_name}' -> {rid}")
                
                self.log2(f"\nLoaded {mapping_count} unique mappings")
                self.mapping_label.config(text=f"Loaded {mapping_count} mappings from {Path(csv_path).name}", fg=COLORS['success'])
                self.csv_file_path.set(csv_path)
                
        except Exception as e:
            self.log2(f"Error loading CSV: {e}")
            import traceback
            self.log2(traceback.format_exc())
            self.mapping_label.config(text=f"Error loading CSV: {e}", fg=COLORS['error'])
    
    def browse_tab2(self):
        f = filedialog.askopenfilename(title="Select PDF File", 
                                   filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if f:
            self.tab2_file.set(f)
            if not self.tab2_output.get():
                self.tab2_output.set(str(Path(f).parent / "Output"))
            
            # If no CSV is loaded yet, try to load default CSV
            if not self.name_map:
                default_csv = self.get_default_csv_path()
                if default_csv:
                    self.csv_file_path.set(default_csv)
                    self.load_custom_mapping(default_csv)
                    self.log2(f"Auto-loaded default CSV: {Path(default_csv).name}")
    
    def browse_output_tab2(self):
        f = filedialog.askdirectory(title="Select output folder for PDFs")
        if f:
            self.tab2_output.set(f)
    
    def log2(self, msg):
        self.log_text2.insert(tk.END, msg + "\n")
        self.log_text2.see(tk.END)
        self.root.update_idletasks()
    
    def process_tab2(self):
        if not self.tab2_file.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        if not self.tab2_output.get():
            messagebox.showerror("Error", "Please select an output folder for PDFs")
            return
        if not self.tab2_naming_selections:
            messagebox.showerror("Error", "Please select at least one criterion for file naming")
            return
        if not self.name_map:
            messagebox.showerror("Error", "Please load a CSV file for name mapping first")
            return
        if not self.criteria_tab2:
            messagebox.showerror("Error", "No criteria defined for Tab 2. Add criteria using the Criteria menu.")
            return
        
        self.process_btn2.config(state=tk.DISABLED)
        self.progress2.start()
        self.status_text2.set("Processing...")
        self.log_text2.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._process_tab2_thread)
        thread.start()
    
    def _process_tab2_thread(self):
        try:
            input_path = Path(self.tab2_file.get())
            
            if input_path.suffix.lower() != '.pdf':
                self.root.after(0, lambda: messagebox.showerror("Error", "Only PDF files are supported."))
                return
            
            output_path = Path(self.tab2_output.get())
            output_path.mkdir(parents=True, exist_ok=True)
            
            naming_indices = self.tab2_naming_selections
            separator = self.naming_separator.get()
            suffix = self.filename_suffix.get().strip()
            
            self.log2(f"\n=== STARTING PROCESSING ===")
            self.log2(f"PDF: {input_path.name}")
            self.log2(f"Output: {output_path}")
            self.log2(f"Tab 2 Criteria: {[c['name'] for c in self.criteria_tab2]}")
            self.log2(f"Mappings available: {len(self.name_map)}")
            
            # Find ID and Name criteria from Tab 2
            id_criterion = None
            name_criterion = None
            
            for crit in self.criteria_tab2:
                if crit["type"] == "id":
                    id_criterion = crit
                    self.log2(f"ID criterion: '{id_criterion['name']}'")
                elif 'name' in crit["name"].lower():
                    name_criterion = crit
                    self.log2(f"Name criterion: '{name_criterion['name']}'")
            
            if not id_criterion and len(self.criteria_tab2) >= 1:
                id_criterion = self.criteria_tab2[0]
                self.log2(f"Using first criterion as ID: '{id_criterion['name']}'")
            
            if not name_criterion and len(self.criteria_tab2) >= 2:
                name_criterion = self.criteria_tab2[1]
                self.log2(f"Using second criterion as Name: '{name_criterion['name']}'")
            
            if not name_criterion:
                self.log2("ERROR: No name criterion found for matching!")
                return
            
            doc = fitz.open(input_path)
            sections = []
            for page_num in range(len(doc)):
                text = doc[page_num].get_text()
                sections.append((page_num, text))
            doc.close()
            
            total_pages = len(sections)
            self.log2(f"Total pages: {total_pages}")
            
            csv_rows = []
            matched_count = 0
            
            for section_num, (page_num, text) in enumerate(sections):
                self.log2(f"\n--- Page {page_num + 1} ---")
                
                # Extract ALL criteria from the text with proper multi-line support
                page_values = {}
                for crit in self.criteria_tab2:
                    prefix = crit["prefix"].strip()
                    stop_text = crit.get("stop_text", "")
                    
                    # Find the position of the prefix in the text
                    prefix_pos = text.find(prefix)
                    
                    if prefix_pos != -1:
                        # Start after the prefix
                        start_pos = prefix_pos + len(prefix)
                        
                        # Find where the value ends
                        if stop_text:
                            # Look for stop text after the prefix
                            stop_pos = text.find(stop_text, start_pos)
                            if stop_pos != -1:
                                raw_value = text[start_pos:stop_pos].strip()
                            else:
                                raw_value = text[start_pos:].strip()
                        else:
                            # No stop text, take until next newline or end
                            end_of_line = text.find('\n', start_pos)
                            if end_of_line != -1:
                                raw_value = text[start_pos:end_of_line].strip()
                            else:
                                raw_value = text[start_pos:].strip()
                        
                        # Clean up the value: replace newlines and multiple spaces with a single space
                        raw_value = re.sub(r'\s+', ' ', raw_value).strip()
                        page_values[crit["name"]] = raw_value
                        self.log2(f"  Extracted {crit['name']}: {raw_value[:100]}...")
                        if stop_text:
                            self.log2(f"    (Stopped at: '{stop_text}')")
                    else:
                        page_values[crit["name"]] = "Not Found"
                        self.log2(f"  {crit['name']}: Not Found (prefix '{prefix}' not found)")
                
                match_name = page_values.get(name_criterion["name"], "Unknown")
                self.log2(f"  Name to match: '{match_name}'")
                
                matched_id = "Not Found"
                if match_name != "Not Found" and match_name != "Unknown" and len(match_name) > 2:
                    clean_name = ' '.join(match_name.split())
                    clean_name_lower = clean_name.lower()
                    
                    if clean_name in self.name_map:
                        matched_id = self.name_map[clean_name]
                        self.log2(f"  ✓ Direct match: '{clean_name}' -> {matched_id}")
                        matched_count += 1
                    elif clean_name_lower in self.name_map:
                        matched_id = self.name_map[clean_name_lower]
                        self.log2(f"  ✓ Case-insensitive match -> {matched_id}")
                        matched_count += 1
                    else:
                        for csv_name, sid in self.name_map.items():
                            if len(csv_name) < 4:
                                continue
                            if clean_name_lower in csv_name.lower() or csv_name.lower() in clean_name_lower:
                                matched_id = sid
                                self.log2(f"  ✓ Partial match: '{csv_name}' -> {matched_id}")
                                matched_count += 1
                                break
                        
                        if matched_id == "Not Found":
                            self.log2(f"  ✗ No match found for: '{match_name}'")
                else:
                    self.log2(f"  Cannot match: name is invalid")
                
                page_values[id_criterion["name"]] = matched_id
                self.log2(f"  Set ID to: {matched_id}")
                
                filename = self.build_filename(page_values, naming_indices, self.criteria_tab2, separator, suffix, for_tab2=True)
                output_file = output_path / filename
                
                if output_file.exists():
                    counter = 1
                    stem = Path(filename).stem
                    while output_file.exists():
                        output_file = output_path / f"{stem}_{counter}.pdf"
                        counter += 1
                    self.log2(f"  Duplicate filename, saved as: {output_file.name}")
                
                reader = PyPDF2.PdfReader(str(input_path))
                writer = PyPDF2.PdfWriter()
                writer.add_page(reader.pages[page_num])
                with open(output_file, 'wb') as f:
                    writer.write(f)
                
                row = {}
                for crit in self.criteria_tab2:
                    row[crit["name"]] = page_values.get(crit["name"], "Not Found")
                row["filename"] = output_file.name
                csv_rows.append(row)
                
                status = "MATCHED" if matched_id != "Not Found" else "NOT FOUND"
                self.log2(f"  {status}: {output_file.name}")
                self.root.after(0, lambda p=section_num+1, t=total_pages: self.status_text2.set(f"Processed {p}/{total_pages} pages - {matched_count} matched"))
            
            if csv_rows:
                csv_output_path = Path(self.csv_output_folder.get()) if self.csv_output_folder.get() else output_path
                csv_output_path.mkdir(parents=True, exist_ok=True)
                
                csv_name = self.csv_filename_tab2.get().strip()
                if not csv_name:
                    csv_name = "extracted_data_tab2"
                if not csv_name.endswith('.csv'):
                    csv_name += '.csv'
                
                csv_file = csv_output_path / csv_name
                fieldnames = [crit["name"] for crit in self.criteria_tab2] + ["filename"]
                
                with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
                    writer.writeheader()
                    writer.writerows(csv_rows)
                
                matched = sum(1 for row in csv_rows if row.get(id_criterion["name"], "Not Found") != "Not Found")
                not_matched = len(csv_rows) - matched
                
                self.log2(f"\n=== COMPLETE ===")
                self.log2(f"Processed: {len(csv_rows)} pages")
                self.log2(f"Matched: {matched}")
                self.log2(f"Not matched: {not_matched}")
                self.log2(f"Output folder: {output_path}")
                self.log2(f"CSV saved: {csv_file}")
                self.log2(f"Column order: {', '.join(fieldnames)}")
                
                self.root.after(0, lambda: messagebox.showinfo("Success", 
                    f"Processed {len(csv_rows)} pages!\n\nMatched: {matched}\nUnmatched: {not_matched}\n\nOutput: {output_path}\nCSV: {csv_file}"))
            else:
                self.log2("No data extracted")
            
        except Exception as e:
            self.log2(f"\n[ERROR] {str(e)}")
            import traceback
            self.log2(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self._processing_done_tab2)
    
    def debug_csv_structure(self):
        csv_path = self.csv_file_path.get()
        if not csv_path:
            self.log2("No CSV file selected. Please browse for a CSV file first.")
            return
        
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                r = csv.DictReader(f)
                self.log2(f"\n=== CSV STRUCTURE ===")
                self.log2(f"File: {Path(csv_path).name}")
                self.log2(f"Columns: {', '.join(r.fieldnames)}")
                
                rows = list(r)
                self.log2(f"Number of rows: {len(rows)}")
                
                if rows:
                    self.log2(f"\nFirst row sample:")
                    for key, value in rows[0].items():
                        self.log2(f"  {key}: {value}")
                    
                    if len(rows) > 1:
                        self.log2(f"\nLast row sample:")
                        for key, value in rows[-1].items():
                            self.log2(f"  {key}: {value}")
                    
                    name_col = None
                    id_col = None
                    for col in r.fieldnames:
                        if 'name' in col.lower():
                            name_col = col
                        if 'id' in col.lower() or 'document' in col.lower():
                            id_col = col
                    
                    self.log2(f"\nSuggested mapping:")
                    self.log2(f"  Use Name column: '{name_col or '(second column)'}'")
                    self.log2(f"  Use ID column: '{id_col or '(first column)'}'")
                    
        except Exception as e:
            self.log2(f"Error debugging CSV: {e}")

    def _processing_done_tab2(self):
        self.progress2.stop()
        self.process_btn2.config(state=tk.NORMAL)
        self.status_text2.set("Ready")
    
    # ==================== TAB 3: SPLIT BY PAGE RANGE ====================
    def init_pagesplit_tab(self):
        self.page_input_file = tk.StringVar()
        self.page_output_folder = tk.StringVar()
        self.status_text3 = tk.StringVar(value="Ready")
        
        main = tk.Frame(self.tab3, bg=COLORS['card'])
        main.pack(fill='both', expand=True, padx=15, pady=15)
        
        file_frame = ttk.LabelFrame(main, text="📁 1. Select PDF File")
        file_frame.pack(fill='x', pady=(0, 10))
        
        file_row = tk.Frame(file_frame, bg=COLORS['card'])
        file_row.pack(fill='x', padx=10, pady=10)
        
        file_entry = ttk.Entry(file_row, textvariable=self.page_input_file, font=('Segoe UI', 9))
        file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(file_row, text="Browse", command=self.browse_page_input, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        output_frame = ttk.LabelFrame(main, text="📂 2. Select Output Folder")
        output_frame.pack(fill='x', pady=(0, 10))
        
        output_row = tk.Frame(output_frame, bg=COLORS['card'])
        output_row.pack(fill='x', padx=10, pady=10)
        
        output_entry = ttk.Entry(output_row, textvariable=self.page_output_folder, font=('Segoe UI', 9))
        output_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        tk.Button(output_row, text="Browse", command=self.browse_page_output, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        ranges_frame = ttk.LabelFrame(main, text="📄 3. Page Ranges")
        ranges_frame.pack(fill='x', pady=(0, 10))
        
        ranges_row = tk.Frame(ranges_frame, bg=COLORS['card'])
        ranges_row.pack(fill='x', padx=10, pady=10)
        
        tk.Label(ranges_row, text="Page Ranges:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=12, anchor='w').pack(side='left')
        ranges_entry = ttk.Entry(ranges_row, textvariable=self.page_ranges, font=('Segoe UI', 9), width=50)
        ranges_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        examples_row = tk.Frame(ranges_frame, bg=COLORS['card'])
        examples_row.pack(fill='x', padx=10, pady=5)
        
        examples_text = "Examples: '5-10' (pages 5-10), '1,3,5' (pages 1,3,5), '1-5,10-15,20' (multiple ranges, up to 20)"
        tk.Label(examples_row, text=examples_text, bg=COLORS['card'], fg=COLORS['text_light'], 
                font=('Segoe UI', 8)).pack(anchor='w')
        
        preview_row = tk.Frame(ranges_frame, bg=COLORS['card'])
        preview_row.pack(fill='x', padx=10, pady=5)
        
        refresh_btn = tk.Button(preview_row, text="⟳ Refresh Page Preview", command=self.refresh_page_preview,
                               bg=COLORS['button_bg'], fg=COLORS['button_fg'], relief='flat', padx=10)
        refresh_btn.pack(side='left', padx=5)
        
        self.page_preview_label = tk.Label(preview_row, text="No ranges entered yet.", bg=COLORS['card'], 
                                          fg=COLORS['warning'], font=('Segoe UI', 9))
        self.page_preview_label.pack(side='left', padx=10)
        
        csv_frame = ttk.LabelFrame(main, text="📊 4. CSV Output Settings (Optional)")
        csv_frame.pack(fill='x', pady=(0, 10))
        
        csv_folder_row = tk.Frame(csv_frame, bg=COLORS['card'])
        csv_folder_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_folder_row, text="CSV Folder:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=10, anchor='w').pack(side='left')
        tk.Entry(csv_folder_row, textvariable=self.csv_output_folder, font=('Segoe UI', 9), 
                width=60).pack(side='left', fill='x', expand=True, padx=(5, 5))
        tk.Button(csv_folder_row, text="Browse", command=self.browse_csv_folder, bg=COLORS['button_bg'], 
                 fg=COLORS['button_fg'], relief='flat', padx=10).pack(side='left')
        
        csv_name_row = tk.Frame(csv_frame, bg=COLORS['card'])
        csv_name_row.pack(fill='x', padx=10, pady=5)
        tk.Label(csv_name_row, text="CSV Filename:", bg=COLORS['card'], fg=COLORS['text_light'], 
                width=10, anchor='w').pack(side='left')
        tk.Entry(csv_name_row, textvariable=self.csv_filename_tab3, font=('Segoe UI', 9), 
                width=30).pack(side='left', padx=5)
        tk.Label(csv_name_row, text=".csv will be added automatically", bg=COLORS['card'], 
                fg=COLORS['text_light'], font=('Segoe UI', 8)).pack(side='left', padx=5)
        
        self.process_btn3 = tk.Button(main, text=self.tab3_button_text.get(), command=self.process_page_split,
                                      bg=COLORS['accent'], fg=COLORS['accent_fg'], font=('Segoe UI', 10, 'bold'), 
                                      relief='flat', width=30, pady=5)
        self.process_btn3.pack(pady=10)
        
        self.progress3 = ttk.Progressbar(main, mode="indeterminate")
        self.progress3.pack(fill='x', pady=5)
        
        tk.Label(main, textvariable=self.status_text3, bg=COLORS['card'], fg=COLORS['text_light']).pack(pady=5)
        
        log_frame = ttk.LabelFrame(main, text="📋 Processing Log")
        log_frame.pack(fill='both', expand=True)
        
        log_inner = tk.Frame(log_frame, bg=COLORS['card'])
        log_inner.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text3 = tk.Text(log_inner, wrap=tk.WORD, font=('Consolas', 9), 
                                bg=COLORS['log_bg'], fg=COLORS['text'], relief='flat', borderwidth=0, height=12)
        self.log_text3.pack(side='left', fill='both', expand=True)
        
        log_scrollbar3 = tk.Scrollbar(log_inner, command=self.log_text3.yview,
                                     bg=COLORS['scrollbar_bg'], troughcolor=COLORS['scrollbar_trough'],
                                     activebackground=COLORS['scrollbar_active'], relief='flat', borderwidth=0,
                                     highlightthickness=0)
        log_scrollbar3.pack(side='right', fill='y')
        self.log_text3.configure(yscrollcommand=log_scrollbar3.set)
    
    def browse_page_input(self):
        f = filedialog.askopenfilename(title="Select PDF File", 
                                       filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if f:
            self.page_input_file.set(f)
            if not self.page_output_folder.get():
                self.page_output_folder.set(str(Path(f).parent / "Page_Extract"))
    
    def browse_page_output(self):
        f = filedialog.askdirectory(title="Select output folder")
        if f:
            self.page_output_folder.set(f)
    
    def log3(self, msg):
        self.log_text3.insert(tk.END, msg + "\n")
        self.log_text3.see(tk.END)
        self.root.update_idletasks()
    
    def process_page_split(self):
        if not self.page_input_file.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        if not self.page_output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        ranges = self.get_page_ranges()
        if not ranges:
            messagebox.showerror("Error", "Please enter valid page ranges.\n\nExamples: '5-10', '1,3,5', '1-5,10-15'")
            return
        
        self.process_btn3.config(state=tk.DISABLED)
        self.progress3.start()
        self.status_text3.set("Processing...")
        self.log_text3.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._process_page_split_thread)
        thread.start()
    
    def _process_page_split_thread(self):
        try:
            input_path = Path(self.page_input_file.get())
            
            if input_path.suffix.lower() != '.pdf':
                self.root.after(0, lambda: messagebox.showerror("Error", "Only PDF files are supported."))
                return
            
            output_path = Path(self.page_output_folder.get())
            output_path.mkdir(parents=True, exist_ok=True)
            
            ranges = self.get_page_ranges()
            
            self.log3(f"Processing: {input_path.name}")
            self.log3(f"Output folder: {output_path}")
            self.log3(f"Page ranges: {ranges}")
            
            reader = PyPDF2.PdfReader(str(input_path))
            total_pages = len(reader.pages)
            self.log3(f"Total pages in document: {total_pages}")
            
            valid_ranges = []
            for start, end in ranges:
                if start > total_pages or end > total_pages:
                    self.log3(f"Warning: Range {start}-{end} exceeds document pages ({total_pages}). Skipping.")
                else:
                    valid_ranges.append((start, end))
            
            if not valid_ranges:
                self.root.after(0, lambda: messagebox.showerror("Error", "No valid page ranges found."))
                return
            
            csv_rows = []
            range_idx = 1
            
            for start, end in valid_ranges:
                self.log3(f"\nExtracting pages {start}-{end}:")
                
                writer = PyPDF2.PdfWriter()
                pages_extracted = 0
                for page_num in range(start - 1, end):
                    writer.add_page(reader.pages[page_num])
                    pages_extracted += 1
                
                if len(valid_ranges) == 1:
                    filename = f"pages_{start}_to_{end}.pdf"
                else:
                    filename = f"extract_{range_idx:02d}_pages_{start}_to_{end}.pdf"
                
                output_file = output_path / filename
                
                with open(output_file, 'wb') as f:
                    writer.write(f)
                
                self.log3(f"  ✓ Created: {output_file.name} ({pages_extracted} pages)")
                
                csv_rows.append({
                    'range_number': range_idx,
                    'page_range': f"{start}-{end}",
                    'start_page': start,
                    'end_page': end,
                    'pages_extracted': pages_extracted,
                    'filename': filename
                })
                
                range_idx += 1
            
            if csv_rows:
                csv_output_path = Path(self.csv_output_folder.get()) if self.csv_output_folder.get() else output_path
                csv_output_path.mkdir(parents=True, exist_ok=True)
                
                csv_name = self.csv_filename_tab3.get().strip()
                if not csv_name:
                    csv_name = "page_extract"
                if not csv_name.endswith('.csv'):
                    csv_name += '.csv'
                
                csv_file = csv_output_path / csv_name
                
                with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=['range_number', 'page_range', 'start_page', 'end_page', 'pages_extracted', 'filename'],
                                           quoting=csv.QUOTE_ALL)
                    writer.writeheader()
                    writer.writerows(csv_rows)
                
                self.log3(f"\n[SUCCESS] COMPLETE! Extracted {len(valid_ranges)} range(s)")
                self.log3(f"   Output folder: {output_path}")
                self.log3(f"   CSV saved: {csv_file}")
                
                self.root.after(0, lambda: messagebox.showinfo("Success", 
                    f"Successfully extracted {len(valid_ranges)} page range(s)!\n\nOutput saved to: {output_path}\nCSV saved to: {csv_file}"))
            else:
                self.log3("No data extracted")
            
        except Exception as e:
            self.log3(f"\n[ERROR] {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self._processing_done_page)
    
    def _processing_done_page(self):
        self.progress3.stop()
        self.process_btn3.config(state=tk.NORMAL)
        self.status_text3.set("Ready")

def main():
    root = tk.Tk()
    root.config(cursor="arrow")
    
    app = PDFSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()