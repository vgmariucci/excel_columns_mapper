"""
Excel Column Mapper & Data Transfer Tool
A modern, user-friendly application for mapping and transferring data between Excel files.

Author: Open Source Community
License: MIT
Version: 2.0.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from datetime import datetime
from pathlib import Path
import json
from typing import Dict, List, Optional, Tuple

# Add the project root to the path for imports
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

class Config:
    """Configuration settings for the application"""
    
    # Application settings
    APP_NAME = "Excel Column Mapper"
    APP_VERSION = "1.0"
    WINDOW_TITLE = f"{APP_NAME} v{APP_VERSION}"
    WINDOW_SIZE = "1240x600"
    WINDOW_MIN_SIZE = (1000, 600)
    
    # Directories
    ASSETS_DIR = PROJECT_ROOT / "assets"
    LOG_DIR = PROJECT_ROOT / "log"
    SOURCE_FILES_DIR = PROJECT_ROOT / "data" / "source_files"
    THEMES_DIR = PROJECT_ROOT / "themes"
    
    # Files
    HISTORY_FILE = LOG_DIR / "mapping_history.csv"
    CONFIG_FILE = PROJECT_ROOT / "config.json"
    LOGO_FILE = ASSETS_DIR / "logo.png"
    FAVICON_FILE = ASSETS_DIR / "favicon.ico"
    
    # Theme settings
    THEME_NAME = "azure"  # Using Azure theme - a popular open-source theme
    
    @classmethod
    def ensure_directories(cls):
        """Ensure all required directories exist"""
        for directory in [cls.ASSETS_DIR, cls.LOG_DIR, cls.SOURCE_FILES_DIR, cls.THEMES_DIR]:
            directory.mkdir(exist_ok=True)

class ThemeManager:
    """Manages application themes and styling"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.style = ttk.Style()
        self.current_theme = "light"  # Track current theme
        
    def setup_azure_theme(self, theme_mode="light"):
        """Setup the Azure theme - a modern open-source theme"""
        azure_tcl_path = Config.THEMES_DIR / "azure.tcl"
        print(f"Azure.tcl file path: {azure_tcl_path}")
        
        try:
            # Check if azure.tcl exists
            if not azure_tcl_path.exists():
                print(f"Warning: azure.tcl not found at {azure_tcl_path}")
                self.setup_custom_azure_theme()
                return
            
            # Source the azure theme file
            self.root.tk.call("source", str(azure_tcl_path))
            
            # Set the theme based on mode
            if theme_mode == "dark":
                self.style.theme_use("azure-dark")
                self.current_theme = "dark"
            else:
                self.style.theme_use("azure-light")
                self.current_theme = "light"
                
            print(f"Azure theme loaded successfully: {theme_mode} mode")
            
            # Apply additional custom configurations
            self.apply_custom_styles()
            
        except (ImportError, tk.TclError) as e:
            print(f"Failed to load Azure theme: {e}")
            # Fallback to custom Azure-inspired theme
            self.setup_custom_azure_theme()
    
    def switch_theme(self, mode="light"):
        """Switch between light and dark themes"""
        try:
            if mode == "dark":
                self.root.tk.call("set_theme", "dark")
                self.current_theme = "dark"
            else:
                self.root.tk.call("set_theme", "light")
                self.current_theme = "light"
            
            # Reapply custom styles
            self.apply_custom_styles()
            
            # Update any text widgets that need theme-specific colors
            self.update_text_widget_colors()
            
            print(f"Theme switched to {mode} mode")
            
        except tk.TclError as e:
            print(f"Failed to switch theme: {e}")
            # Fallback to custom theme switching
            self.setup_custom_azure_theme()
    
    def apply_custom_styles(self):
        """Apply additional custom styles for specific widgets"""
        # Custom button styles that work with both light and dark themes
        if self.current_theme == "dark":
            colors = {
                'primary': '#0078d4',
                'primary_hover': '#106ebe',
                'success': '#107c10',
                'danger': '#d13438',
                'secondary': '#6c757d',
                'white': '#ffffff',
                'dark_bg': '#2b2b2b',
                'light_bg': '#404040'
            }
        else:
            colors = {
                'primary': '#0078d4',
                'primary_hover': '#106ebe',
                'success': '#107c10',
                'danger': '#d13438',
                'secondary': '#6c757d',
                'white': '#ffffff',
                'dark_bg': '#ffffff',
                'light_bg': '#f5f5f5'
            }
        
        # Configure custom button styles
        self.configure_custom_button_styles(colors)
        self.configure_custom_frame_styles(colors)
        self.configure_custom_label_styles(colors)
        self.configure_custom_entry_styles(colors)
        self.configure_custom_treeview_styles(colors)
        self.configure_custom_notebook_styles(colors)
    
    def get_theme_colors(self):
        """Get current theme colors for use by other components"""
        if self.current_theme == "dark":
            return {
                'bg': '#2b2b2b',
                'fg': '#ffffff',
                'select_bg': '#0078d4',
                'select_fg': '#ffffff',
                'field_bg': '#3c3c3c',
                'border': '#555555'
            }
        else:
            return {
                'bg': '#ffffff',
                'fg': '#000000',
                'select_bg': '#0078d4',
                'select_fg': '#ffffff',
                'field_bg': '#ffffff',
                'border': '#d1d1d1'
            }

    def update_text_widget_colors(self):
        """Update text widget colors based on current theme"""
        try:
            # This method can be called from the main app to update text widgets
            # We'll implement this in the main app class
            pass
        except Exception as e:
            print(f"Error updating text widget colors: {e}")

    def configure_custom_button_styles(self, colors: Dict[str, str]):
        """Configure custom button styles"""
        # Primary button
        self.style.configure(
            'Primary.TButton',
            background=colors['primary'],
            foreground=colors['white'],
            borderwidth=0,
            focuscolor='none',
            font=('Segoe UI', 10),
            padding=(8, 4, 8, 4)
        )
        self.style.map(
            'Primary.TButton',
            background=[('active', colors['primary_hover']), ('pressed', colors['primary_hover'])],
            relief=[('pressed', 'flat'), ('!pressed', 'flat')]
        )
        
        # Success button
        self.style.configure(
            'Success.TButton',
            background=colors['success'],
            foreground=colors['white'],
            borderwidth=0,
            focuscolor='none',
            font=('Segoe UI', 10),
            padding=(8, 4, 8, 4)
        )
        self.style.map(
            'Success.TButton',
            background=[('active', '#0e6b0e'), ('pressed', '#0e6b0e')],
            relief=[('pressed', 'flat'), ('!pressed', 'flat')]
        )
        
        # Danger button
        self.style.configure(
            'Danger.TButton',
            background=colors['danger'],
            foreground=colors['white'],
            borderwidth=0,
            focuscolor='none',
            font=('Segoe UI', 10),
            padding=(8, 4, 8, 4)
        )
        self.style.map(
            'Danger.TButton',
            background=[('active', '#b02a2e'), ('pressed', '#b02a2e')],
            relief=[('pressed', 'flat'), ('!pressed', 'flat')]
        )
        
        # Secondary button
        self.style.configure(
            'Secondary.TButton',
            background=colors['secondary'],
            foreground=colors['white'],
            borderwidth=0,
            focuscolor='none',
            font=('Segoe UI', 10),
            padding=(8, 4, 8, 4)
        )
        self.style.map(
            'Secondary.TButton',
            background=[('active', '#5a6268'), ('pressed', '#5a6268')],
            relief=[('pressed', 'flat'), ('!pressed', 'flat')]
        )
    
    def configure_custom_frame_styles(self, colors: Dict[str, str]):
        """Configure custom frame styles"""
        self.style.configure(
            'Card.TFrame',
            background=colors['dark_bg'],
            relief='flat',
            borderwidth=1
        )
        
        self.style.configure(
            'Sidebar.TFrame',
            background=colors['light_bg'],
            relief='flat',
            borderwidth=0
        )
    
    def configure_custom_label_styles(self, colors: Dict[str, str]):
        """Configure custom label styles"""
        label_bg = colors['dark_bg']
        label_fg = colors['white'] if self.current_theme == "dark" else '#000000'
        
        self.style.configure(
            'Heading.TLabel',
            background=label_bg,
            foreground=label_fg,
            font=('Segoe UI', 16, 'bold')
        )
        
        self.style.configure(
            'Subheading.TLabel',
            background=label_bg,
            foreground=label_fg,
            font=('Segoe UI', 12, 'bold')
        )
        
        self.style.configure(
            'Body.TLabel',
            background=label_bg,
            foreground=label_fg,
            font=('Segoe UI', 10)
        )
        
        self.style.configure(
            'Caption.TLabel',
            background=label_bg,
            foreground='#737373',
            font=('Segoe UI', 9)
        )
    
    def configure_custom_entry_styles(self, colors: Dict[str, str]):
        """Configure custom entry styles"""
        self.style.configure(
            'Modern.TEntry',
            fieldbackground=colors['dark_bg'] if self.current_theme == "light" else '#3c3c3c',
            borderwidth=1,
            relief='solid',
            insertcolor='#000000' if self.current_theme == "light" else '#ffffff'
        )
    
    def configure_custom_treeview_styles(self, colors: Dict[str, str]):
        """Configure custom treeview styles"""
        bg_color = colors['dark_bg'] if self.current_theme == "light" else '#3c3c3c'
        fg_color = '#000000' if self.current_theme == "light" else '#ffffff'
        
        self.style.configure(
            'Modern.Treeview',
            background=bg_color,
            foreground=fg_color,
            fieldbackground=bg_color,
            borderwidth=1,
            relief='solid'
        )
        
        self.style.configure(
            'Modern.Treeview.Heading',
            background=colors['primary'],
            foreground=colors['white'],
            font=('Segoe UI', 10, 'bold'),
            relief='flat'
        )
    
    def configure_custom_notebook_styles(self, colors: Dict[str, str]):
        """Configure custom notebook styles"""
        self.style.configure(
            'Modern.TNotebook',
            background=colors['light_bg'],
            borderwidth=0
        )
        
        tab_fg = '#000000' if self.current_theme == "light" else '#ffffff'
        
        self.style.configure(
            'Modern.TNotebook.Tab',
            background=colors['light_bg'],
            foreground=tab_fg,
            padding=[20, 10],
            font=('Segoe UI', 10)
        )
        
        self.style.map(
            'Modern.TNotebook.Tab',
            background=[('selected', colors['dark_bg']), ('active', colors['light_bg'])],
            foreground=[('selected', colors['primary']), ('active', tab_fg)]
        )
    
    def setup_custom_azure_theme(self):
        """Setup a custom Azure-inspired theme as fallback"""
        print("Using fallback custom Azure theme")
        
        # Configure base theme
        self.style.theme_use('clam')
        
        # Apply custom styles
        self.apply_custom_styles()

class StatisticsManager:
    """Manages user statistics and achievements"""
    
    def __init__(self):
        self.stats = {
            'mappings_created': 0,
            'files_processed': 0,
            'total_columns_mapped': 0,
            'sessions_completed': 0,
            'last_activity': None
        }
        self.load_stats()
    
    def load_stats(self):
        """Load statistics from file"""
        try:
            if Config.CONFIG_FILE.exists():
                with open(Config.CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    self.stats.update(data.get('statistics', {}))
        except Exception as e:
            print(f"Warning: Could not load statistics: {e}")
    
    def save_stats(self):
        """Save statistics to file"""
        try:
            config_data = {}
            if Config.CONFIG_FILE.exists():
                with open(Config.CONFIG_FILE, 'r') as f:
                    config_data = json.load(f)
            
            config_data['statistics'] = self.stats
            config_data['last_updated'] = datetime.now().isoformat()
            
            with open(Config.CONFIG_FILE, 'w') as f:
                json.dump(config_data, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save statistics: {e}")
    
    def update_mapping_created(self):
        """Update mapping created count"""
        self.stats['mappings_created'] += 1
        self.stats['last_activity'] = datetime.now().isoformat()
        self.save_stats()
    
    def update_file_processed(self, columns_count: int):
        """Update file processed count"""
        self.stats['files_processed'] += 1
        self.stats['total_columns_mapped'] += columns_count
        self.stats['sessions_completed'] += 1
        self.stats['last_activity'] = datetime.now().isoformat()
        self.save_stats()

class ExcelColumnMapper:
    """Main application class for Excel Column Mapping"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.setup_window()
        
        # Initialize managers
        self.theme_manager = ThemeManager(root)
        self.stats_manager = StatisticsManager()
        
        # Application state
        self.source_file_path = tk.StringVar()
        self.destination_file_path = tk.StringVar()
        self.source_headers: List[str] = []
        self.destination_headers: List[str] = []
        self.source_df: Optional[pd.DataFrame] = None
        self.destination_df: Optional[pd.DataFrame] = None
        self.column_mappings: Dict[str, str] = {}
        self.mapping_combos: Dict[str, ttk.Combobox] = {}
        self.combo_values: List[str] = []
        
        # Setup theme and create UI
        self.theme_manager.setup_azure_theme("light")
        self.create_widgets()
        
        # Ensure directories exist
        Config.ensure_directories()
    
    def setup_window(self):
        """Setup main window properties"""
        self.root.title(Config.WINDOW_TITLE)
        self.root.geometry(Config.WINDOW_SIZE)
        self.root.minsize(*Config.WINDOW_MIN_SIZE)
        
        # Set window icon if available
        try:
            if Config.FAVICON_FILE.exists():
                self.root.iconbitmap(str(Config.FAVICON_FILE))
        except Exception:
            pass
        
        # Center window on screen
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def toggle_theme(self):
        """Toggle between light and dark themes"""
        try:
            # Switch to opposite theme
            new_theme = "light" if self.theme_manager.current_theme == "dark" else "dark"
            
            # Apply the theme
            self.theme_manager.switch_theme(new_theme)
            
            # Update button text
            self.theme_button.config(text=self.get_theme_button_text())
            
            # Update text widgets that don't automatically inherit theme colors
            self.update_text_widgets_theme()
            
            # Update status
            self.update_status(f"Switched to {new_theme} mode")
            
            # Force a refresh of the interface to apply theme changes
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"Error toggling theme: {e}")
            messagebox.showerror("Theme Error", f"Failed to switch theme: {str(e)}")

    def get_theme_button_text(self):
        """Get the appropriate text for the theme button"""
        if self.theme_manager.current_theme == "dark":
            return "☀ Light Mode"
        else:
            return "🌙 Dark Mode"

    def update_text_widgets_theme(self):
        """Update text widgets to match current theme"""
        try:
            colors = self.theme_manager.get_theme_colors()
            
            # Update preview text widget if it exists
            if hasattr(self, 'preview_text'):
                self.preview_text.config(
                    bg=colors['field_bg'],
                    fg=colors['fg'],
                    insertbackground=colors['fg'],
                    selectbackground=colors['select_bg'],
                    selectforeground=colors['select_fg']
                )
            
            # Update any other text widgets or canvas elements
            if hasattr(self, 'mapping_canvas'):
                self.mapping_canvas.config(bg=colors['bg'])
                
        except Exception as e:
            print(f"Error updating text widget colors: {e}")

    def create_theme_toggle_button(self, parent):
        """Create theme toggle button"""
        theme_frame = ttk.Frame(parent)
        theme_frame.pack(anchor=tk.E)
        
        # Create the theme toggle button
        self.theme_button = ttk.Button(
            theme_frame,
            text=self.get_theme_button_text(),
            command=self.toggle_theme,
            style='Secondary.TButton',
            width=14
        )
        self.theme_button.pack()
        
        # Store reference for updates
        self.theme_toggle_button = self.theme_button

    def create_widgets(self):
        """Create and layout all widgets"""
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)
        
        # Header section
        self.create_header_section(main_container)
        
        # Control panel
        self.create_control_panel(main_container)
        
        # Main content area with notebook
        self.create_main_content(main_container)
        
        # Status bar
        self.create_status_bar(main_container)
    
    def create_header_section(self, parent):
        """Create the header section with title, statistics, and theme toggle"""
        header_frame = ttk.Frame(parent, style='Card.TFrame', padding="20")
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(0, weight=1)
        
        # Title and subtitle container
        title_container = ttk.Frame(header_frame)
        title_container.grid(row=0, column=0, sticky=(tk.W, tk.E))
        title_container.columnconfigure(0, weight=1)
        
        # Title and subtitle
        title_label = ttk.Label(
            title_container, 
            text=Config.APP_NAME, 
            style='Heading.TLabel'
        )
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        subtitle_label = ttk.Label(
            title_container,
            text="Map columns and transfer data between Excel files with ease",
            style='Body.TLabel'
        )
        subtitle_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        
        # Right side container for theme button and statistics
        right_container = ttk.Frame(header_frame)
        right_container.grid(row=0, column=1, sticky=(tk.N, tk.E))
        
        # Theme toggle button at the top right
        self.create_theme_toggle_button(right_container)
        
        # Statistics panel below the theme button
        stats_container = ttk.Frame(right_container)
        stats_container.pack(pady=(10, 0))
        self.create_statistics_panel(stats_container)

    def get_theme_button_text(self):
        """Get the appropriate text for the theme button"""
        if self.theme_manager.current_theme == "dark":
            return "☀ Light Mode"
        else:
            return "🌙 Dark Mode"

    def toggle_theme(self):
        """Toggle between light and dark themes"""
        try:
            # Switch to opposite theme
            new_theme = "light" if self.theme_manager.current_theme == "dark" else "dark"
            
            # Apply the theme
            self.theme_manager.switch_theme(new_theme)
            
            # Update button text
            self.theme_button.config(text=self.get_theme_button_text())
            
            # Update status
            self.update_status(f"Switched to {new_theme} mode")
            
            # Force a refresh of the interface to apply theme changes
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"Error toggling theme: {e}")
            messagebox.showerror("Theme Error", f"Failed to switch theme: {str(e)}")

    def create_statistics_panel(self, parent):
        """Create statistics display panel"""
        stats = self.stats_manager.stats
        
        ttk.Label(parent, text="Statistics", style='Subheading.TLabel').pack(anchor=tk.W)
        
        stats_text = f"""Files Processed: {stats['files_processed']}
Mappings Created: {stats['mappings_created']}
Columns Mapped: {stats['total_columns_mapped']}"""
        
        ttk.Label(parent, text=stats_text, style='Caption.TLabel').pack(anchor=tk.W)
    
    def create_control_panel(self, parent):
        """Create the control panel with file selection and action buttons"""
        control_frame = ttk.Frame(parent, style='Card.TFrame', padding="20")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        control_frame.columnconfigure(1, weight=1)
        
        # File selection section
        self.create_file_selection(control_frame)
        
        # Action buttons
        self.create_action_buttons(control_frame)
    
    def create_file_selection(self, parent):
        """Create file selection widgets"""
        # Source file
        ttk.Label(parent, text="Source File:", style='Body.TLabel').grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10)
        )
        
        source_frame = ttk.Frame(parent)
        source_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        source_frame.columnconfigure(0, weight=1)
        
        self.source_entry = ttk.Entry(
            source_frame, 
            textvariable=self.source_file_path, 
            state="readonly",
            style='Modern.TEntry'
        )
        self.source_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(
            source_frame, 
            text="Browse...", 
            command=self.browse_source_file,
            style='Secondary.TButton'
        ).grid(row=0, column=1)
        
        # Destination file
        ttk.Label(parent, text="Destination File:", style='Body.TLabel').grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10)
        )
        
        dest_frame = ttk.Frame(parent)
        dest_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        dest_frame.columnconfigure(0, weight=1)
        
        self.dest_entry = ttk.Entry(
            dest_frame, 
            textvariable=self.destination_file_path, 
            state="readonly",
            style='Modern.TEntry'
        )
        self.dest_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(
            dest_frame, 
            text="Browse...", 
            command=self.browse_destination_file,
            style='Secondary.TButton'
        ).grid(row=0, column=1)
    
    def create_action_buttons(self, parent):
        """Create action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # Primary actions
        self.load_button = ttk.Button(
            button_frame,
            text="Load Column Headers",
            command=self.load_headers,
            style='Primary.TButton'
        )
        self.load_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.copy_button = ttk.Button(
            button_frame,
            text="Copy Mapped Data",
            command=self.copy_mapped_data,
            style='Success.TButton',
            state=tk.DISABLED
        )
        self.copy_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Secondary actions
        ttk.Button(
            button_frame,
            text="Clear Mappings",
            command=self.clear_mappings,
            style='Danger.TButton'
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.history_button = ttk.Button(
            button_frame,
            text="Load from History",
            command=self.load_from_history,
            style='Secondary.TButton',
            state=tk.DISABLED
        )
        self.history_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            button_frame,
            text="View History",
            command=self.view_mapping_history,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT)
    
    def create_main_content(self, parent):
        """Create main content area with notebook tabs"""
        # Create notebook
        self.notebook = ttk.Notebook(parent, style='Modern.TNotebook')
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # Mapping tab
        mapping_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(mapping_frame, text="Column Mapping")
        
        self.create_mapping_interface(mapping_frame)
        
        # Preview tab
        preview_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(preview_frame, text="Data Preview")
        
        self.create_preview_interface(preview_frame)
    
    def create_mapping_interface(self, parent):
        """Create the column mapping interface"""
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)
        
        # Source columns section
        source_label = ttk.Label(parent, text="Source Columns", style='Subheading.TLabel')
        source_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=(0, 10))
        
        source_frame = ttk.Frame(parent, style='Card.TFrame')
        source_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        source_frame.columnconfigure(0, weight=1)
        source_frame.rowconfigure(0, weight=1)
        
        # Source treeview
        self.source_tree = ttk.Treeview(
            source_frame,
            columns=('sample',),
            show='tree headings',
            style='Modern.Treeview'
        )
        self.source_tree.heading('#0', text='Column Name', anchor=tk.W)
        self.source_tree.heading('sample', text='Sample Data', anchor=tk.W)
        self.source_tree.column('#0', width=250, minwidth=200)
        self.source_tree.column('sample', width=200, minwidth=150)
        
        source_scrollbar = ttk.Scrollbar(source_frame, orient=tk.VERTICAL, command=self.source_tree.yview)
        self.source_tree.configure(yscrollcommand=source_scrollbar.set)
        
        self.source_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        source_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=10)
        
        # Mapping section
        mapping_label = ttk.Label(parent, text="Column Mappings", style='Subheading.TLabel')
        mapping_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=(0, 10))
        
        mapping_container = ttk.Frame(parent, style='Card.TFrame')
        mapping_container.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        mapping_container.columnconfigure(0, weight=1)
        mapping_container.rowconfigure(0, weight=1)
        
        # Scrollable mapping frame
        self.mapping_canvas = tk.Canvas(mapping_container, highlightthickness=0)
        mapping_scrollbar = ttk.Scrollbar(mapping_container, orient=tk.VERTICAL, command=self.mapping_canvas.yview)
        self.mapping_frame = ttk.Frame(self.mapping_canvas)
        
        self.mapping_frame.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
        
        self.mapping_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        mapping_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=10)
        
        # Bind canvas resize
        self.mapping_canvas.bind('<Configure>', self.on_canvas_configure)
    
    def create_preview_interface(self, parent):
        """Create data preview interface"""
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        # Preview text widget
        preview_label = ttk.Label(parent, text="Data Preview", style='Subheading.TLabel')
        preview_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.preview_text = tk.Text(
            parent,
            wrap=tk.WORD,
            font=('Consolas', 10),
            state=tk.DISABLED
        )
        preview_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        preview_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
    
    def create_status_bar(self, parent):
        """Create status bar"""
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        
        status_frame = ttk.Frame(parent, style='Card.TFrame')
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style='Caption.TLabel',
            padding="10"
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Progress bar (initially hidden)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            mode='determinate'
        )
    
    def on_canvas_configure(self, event):
        """Handle canvas configuration"""
        self.mapping_canvas.itemconfig(self.mapping_canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw"), width=event.width)
    
    def show_progress(self, show: bool = True):
        """Show or hide progress bar"""
        if show:
            self.progress_bar.pack(side=tk.RIGHT, padx=(10, 10))
        else:
            self.progress_bar.pack_forget()
    
    def update_status(self, message: str, progress: Optional[float] = None):
        """Update status message and optional progress"""
        self.status_var.set(message)
        if progress is not None:
            self.progress_var.set(progress)
        self.root.update_idletasks()
    
    # File handling methods
    def browse_source_file(self):
        """Browse for source Excel file"""
        initial_dir = Config.SOURCE_FILES_DIR if Config.SOURCE_FILES_DIR.exists() else Path.cwd()
        
        filename = filedialog.askopenfilename(
            title="Select Source Excel File",
            initialdir=str(initial_dir),
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.source_file_path.set(filename)
            self.update_status(f"Source file selected: {Path(filename).name}")
    
    def browse_destination_file(self):
        """Browse for destination Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Destination Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.destination_file_path.set(filename)
            self.update_status(f"Destination file selected: {Path(filename).name}")
    
    def read_excel_data(self, file_path: str) -> pd.DataFrame:
        """Read Excel or CSV data"""
        file_path = Path(file_path)
        
        try:
            if file_path.suffix.lower() == '.csv':
                return pd.read_csv(file_path, encoding='utf-8')
            else:
                return pd.read_excel(file_path)
        except Exception as e:
            raise Exception(f"Error reading file {file_path.name}: {str(e)}")
    
    def load_headers(self):
        """Load column headers from both files"""
        # Validate file selection
        if not self.source_file_path.get():
            messagebox.showerror("Error", "Please select a source file")
            return
        
        if not self.destination_file_path.get():
            messagebox.showerror("Error", "Please select a destination file")
            return
        
        try:
            self.show_progress(True)
            
            # Load source file
            self.update_status("Loading source file...", 25)
            self.source_df = self.read_excel_data(self.source_file_path.get())
            self.source_headers = self.source_df.columns.tolist()
            
            # Load destination file
            self.update_status("Loading destination file...", 50)
            self.destination_df = self.read_excel_data(self.destination_file_path.get())
            self.destination_headers = self.destination_df.columns.tolist()
            
            # Update UI
            self.update_status("Updating interface...", 75)
            self.populate_source_tree()
            self.create_mapping_widgets()
            self.update_preview()
            
            # Enable buttons
            self.copy_button.config(state=tk.NORMAL)
            self.history_button.config(state=tk.NORMAL)
            
            self.update_status(
                f"Files loaded successfully - Source: {len(self.source_headers)} columns, "
                f"Destination: {len(self.destination_headers)} columns", 100
            )
            
            # Hide progress bar after a moment
            self.root.after(1000, lambda: self.show_progress(False))
            
            messagebox.showinfo(
                "Success",
                f"Files loaded successfully!\n\n"
                f"Source: {len(self.source_headers)} columns\n"
                f"Destination: {len(self.destination_headers)} columns"
            )
            
        except Exception as e:
            self.show_progress(False)
            messagebox.showerror("Error", f"Failed to load files:\n{str(e)}")
            self.update_status("Error loading files")
    
    def populate_source_tree(self):
        """Populate source tree with column headers and sample data"""
        # Clear existing items
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
        
        # Add column headers with sample data
        for header in self.source_headers:
            sample_data = self.get_sample_data(header)
            self.source_tree.insert('', 'end', iid=header, text=header, values=(sample_data,))
    
    def get_sample_data(self, column_name: str, max_samples: int = 3) -> str:
        """Get sample data from a column"""
        try:
            if column_name not in self.source_df.columns:
                return "No data"
            
            column_data = self.source_df[column_name].dropna()
            if len(column_data) == 0:
                return "All empty"
            
            # Get unique values
            unique_values = column_data.unique()[:max_samples]
            
            # Format sample strings
            samples = []
            for value in unique_values:
                value_str = str(value)
                if len(value_str) > 20:
                    value_str = value_str[:20] + "..."
                samples.append(value_str)
            
            result = ", ".join(samples)
            if len(unique_values) >= max_samples and len(column_data.unique()) > max_samples:
                result += "..."
            
            return result
            
        except Exception:
            return "Error reading"
    
    def create_mapping_widgets(self):
        """Create mapping widgets for destination columns"""
        # Clear existing widgets
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        self.mapping_combos.clear()
        
        # Prepare combo values
        self.combo_values = ["-- Select Source Column --"] + self.source_headers
        
        # Create mapping widgets
        for i, dest_header in enumerate(self.destination_headers):
            self.create_mapping_row(i, dest_header)
        
        # Update canvas scroll region
        self.mapping_frame.update_idletasks()
        self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
    
    def create_mapping_row(self, row: int, dest_header: str):
        """Create a single mapping row"""
        row_frame = ttk.Frame(self.mapping_frame, padding="5")
        row_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=2)
        row_frame.columnconfigure(1, weight=1)
        
        # Destination column label
        ttk.Label(
            row_frame,
            text=dest_header,
            style='Body.TLabel',
            width=25
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # Mapping combobox
        combo = ttk.Combobox(
            row_frame,
            values=self.combo_values,
            width=30,
            state="readonly"
        )
        combo.grid(row=0, column=1, sticky=(tk.W, tk.E))
        combo.set("-- Select Source Column --")
        
        # Store reference
        self.mapping_combos[dest_header] = combo
        
        # Bind selection event
        combo.bind('<<ComboboxSelected>>', lambda e, dest=dest_header: self.on_mapping_changed(dest))
    
    def on_mapping_changed(self, dest_column: str):
        """Handle mapping selection change"""
        selected_source = self.mapping_combos[dest_column].get()
        
        # Clear previous mapping
        old_source = self.column_mappings.get(dest_column)
        if old_source and old_source != selected_source:
            self.update_source_tree_mapping(old_source, "", False)
        
        # Update mapping
        if selected_source == "-- Select Source Column --":
            if dest_column in self.column_mappings:
                del self.column_mappings[dest_column]
        else:
            self.column_mappings[dest_column] = selected_source
            self.update_source_tree_mapping(selected_source, dest_column, True)
            self.stats_manager.update_mapping_created()
        
        # Update status and preview
        mappings_count = len(self.column_mappings)
        self.update_status(f"Column mappings configured: {mappings_count}")
        self.update_preview()
    
    def update_source_tree_mapping(self, source_column: str, dest_column: str, is_mapped: bool):
        """Update source tree to show mapping status"""
        if source_column in self.source_headers:
            if is_mapped:
                self.source_tree.item(source_column, text=f"✓ {source_column} → {dest_column}")
            else:
                self.source_tree.item(source_column, text=source_column)
    
    def update_preview(self):
        """Update the data preview"""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        
        if not self.column_mappings:
            preview_text = "No column mappings configured yet.\n\nConfigure mappings to see preview."
            self.preview_text.insert(tk.END, "No column mappings configured yet.\n\nConfigure mappings to see preview.")
        else:
            preview_text = "Current Column Mappings:\n" + "="*50 + "\n\n"
            
            for dest_col, source_col in self.column_mappings.items():
                preview_text += f"Destination: {dest_col}\n"
                preview_text += f"Source: {source_col}\n"
                
                # Show sample data transfer
                if hasattr(self, 'source_df') and source_col in self.source_df.columns:
                    sample = self.get_sample_data(source_col, 2)
                    preview_text += f"Sample Data: {sample}\n"
                
                preview_text += "\n"
        
        self.preview_text.insert(tk.END, preview_text)
        self.preview_text.config(state=tk.DISABLED)
    
    def clear_mappings(self):
        """Clear all column mappings"""
        # Reset comboboxes
        for combo in self.mapping_combos.values():
            combo.set("-- Select Source Column --")
        
        # Clear source tree mappings
        for header in self.source_headers:
            self.update_source_tree_mapping(header, "", False)
        
        # Clear mappings dictionary
        self.column_mappings.clear()
        
        # Update UI
        self.update_status("All mappings cleared")
        self.update_preview()
    
    def copy_mapped_data(self):
        """Copy mapped data from source to destination"""
        if not self.column_mappings:
            messagebox.showwarning("Warning", "No column mappings configured")
            return
        
        try:
            # Confirm operation
            mappings_text = "\n".join([f"{dest} ← {source}" for dest, source in self.column_mappings.items()])
            
            if not messagebox.askyesno(
                "Confirm Data Copy",
                f"Copy data with the following mappings?\n\n{mappings_text}\n\nContinue?"
            ):
                return
            
            self.show_progress(True)
            self.update_status("Preparing data transfer...", 10)
            
            # Create result dataframe
            result_df = self.destination_df.copy()
            
            # Copy data for each mapping
            total_mappings = len(self.column_mappings)
            for i, (dest_col, source_col) in enumerate(self.column_mappings.items()):
                progress = 20 + (60 * i / total_mappings)
                self.update_status(f"Copying {source_col} → {dest_col}...", progress)
                
                if source_col in self.source_df.columns:
                    source_data = self.source_df[source_col]
                    
                    # Handle different row counts
                    if len(source_data) > len(result_df):
                        additional_rows = len(source_data) - len(result_df)
                        empty_rows = pd.DataFrame(index=range(additional_rows), columns=result_df.columns)
                        result_df = pd.concat([result_df, empty_rows], ignore_index=True)
                    
                    # Copy data
                    result_df[dest_col] = source_data
            
            # Save file
            self.update_status("Saving file...", 80)
            
            save_path = filedialog.asksaveasfilename(
                title="Save Updated Destination File",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ],
                initialfile=f"{Path(self.destination_file_path.get()).stem}_updated.xlsx"
            )
            
            if save_path:
                save_path = Path(save_path)
                
                if save_path.suffix.lower() == '.csv':
                    result_df.to_csv(save_path, index=False)
                else:
                    result_df.to_excel(save_path, index=False)
                
                # Save history
                self.save_mapping_history(str(save_path))
                
                # Update statistics
                self.stats_manager.update_file_processed(len(self.column_mappings))
                
                self.update_status(f"Data copied successfully to {save_path.name}", 100)
                
                # Hide progress bar after showing completion
                self.root.after(2000, lambda: self.show_progress(False))
                
                messagebox.showinfo(
                    "Success",
                    f"Data copied successfully!\n\n"
                    f"Mapped {len(self.column_mappings)} columns\n"
                    f"Output file: {save_path.name}"
                )
            else:
                self.show_progress(False)
                self.update_status("Save cancelled")
        
        except Exception as e:
            self.show_progress(False)
            messagebox.showerror("Error", f"Failed to copy data:\n{str(e)}")
            self.update_status("Error copying data")
    
    def save_mapping_history(self, output_file_path: str):
        """Save mapping history to CSV file"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            history_data = []
            for dest_col, source_col in self.column_mappings.items():
                history_data.append({
                    'Timestamp': timestamp,
                    'Source_File': Path(self.source_file_path.get()).name,
                    'Destination_File': Path(self.destination_file_path.get()).name,
                    'Output_File': Path(output_file_path).name,
                    'Source_Column': source_col,
                    'Destination_Column': dest_col
                })
            
            new_history_df = pd.DataFrame(history_data)
            
            # Append to existing history or create new
            if Config.HISTORY_FILE.exists():
                existing_history_df = pd.read_csv(Config.HISTORY_FILE)
                combined_history_df = pd.concat([existing_history_df, new_history_df], ignore_index=True)
            else:
                combined_history_df = new_history_df
            
            combined_history_df.to_csv(Config.HISTORY_FILE, index=False)
            
        except Exception as e:
            print(f"Warning: Could not save mapping history: {e}")
    
    def view_mapping_history(self):
        """View mapping history file"""
        if not Config.HISTORY_FILE.exists():
            messagebox.showinfo("No History", "No mapping history found. Perform some data transfers first.")
            return
        
        try:
            # Try to open with default application
            if os.name == 'nt':  # Windows
                os.startfile(str(Config.HISTORY_FILE))
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{Config.HISTORY_FILE}"')
        except Exception:
            messagebox.showinfo("History File", f"History file location:\n{Config.HISTORY_FILE}")
    
    def load_from_history(self):
        """Load column mappings from history"""
        if not Config.HISTORY_FILE.exists():
            messagebox.showinfo("No History", "No mapping history found.")
            return
        
        if not self.source_headers or not self.destination_headers:
            messagebox.showwarning("Warning", "Please load both source and destination files first.")
            return
        
        try:
            history_df = pd.read_csv(Config.HISTORY_FILE)
            
            if history_df.empty:
                messagebox.showinfo("No History", "No mapping history found.")
                return
            
            self.show_history_selection_dialog(history_df)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load history:\n{str(e)}")
    
    def show_history_selection_dialog(self, history_df: pd.DataFrame):
        """Show dialog to select mapping configuration from history"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Load Mapping from History")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Instructions
        ttk.Label(
            main_frame,
            text="Select a mapping configuration to load:",
            style='Subheading.TLabel'
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # History treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Group history by operation
        grouped_history = history_df.groupby(['Timestamp', 'Source_File', 'Destination_File', 'Output_File'])
        
        columns = ('timestamp', 'source_file', 'dest_file', 'output_file', 'mappings_count')
        history_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style='Modern.Treeview')
        
        # Configure columns
        history_tree.heading('timestamp', text='Date & Time')
        history_tree.heading('source_file', text='Source File')
        history_tree.heading('dest_file', text='Destination File')
        history_tree.heading('output_file', text='Output File')
        history_tree.heading('mappings_count', text='Mappings')
        
        history_tree.column('timestamp', width=150)
        history_tree.column('source_file', width=150)
        history_tree.column('dest_file', width=150)
        history_tree.column('output_file', width=150)
        history_tree.column('mappings_count', width=80)
        
        # Scrollbar
        history_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=history_tree.yview)
        history_tree.configure(yscrollcommand=history_scrollbar.set)
        
        history_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        history_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Populate tree
        history_records = {}
        for (timestamp, source_file, dest_file, output_file), group in grouped_history:
            mappings_count = len(group)
            item_id = history_tree.insert('', 'end', values=(
                timestamp, source_file, dest_file, output_file, mappings_count
            ))
            history_records[item_id] = group
        
        # Details section
        details_frame = ttk.LabelFrame(main_frame, text="Mapping Details", padding="10")
        details_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        details_frame.columnconfigure(0, weight=1)
        
        details_text = tk.Text(details_frame, height=6, wrap=tk.WORD, font=('Consolas', 9))
        details_scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, command=details_text.yview)
        details_text.configure(yscrollcommand=details_scrollbar.set)
        
        details_text.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        details_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        def on_history_select(event):
            selection = history_tree.selection()
            if selection:
                selected_group = history_records[selection[0]]
                details_text.delete(1.0, tk.END)
                details_text.insert(tk.END, "Column Mappings:\n")
                for _, row in selected_group.iterrows():
                    details_text.insert(tk.END, f"• {row['Destination_Column']} ← {row['Source_Column']}\n")
        
        history_tree.bind('<<TreeviewSelect>>', on_history_select)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=tk.E)
        
        def load_selected():
            selection = history_tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a mapping configuration.")
                return
            
            selected_group = history_records[selection[0]]
            
            # Check compatibility
            missing_source = []
            missing_dest = []
            
            for _, row in selected_group.iterrows():
                if row['Source_Column'] not in self.source_headers:
                    missing_source.append(row['Source_Column'])
                if row['Destination_Column'] not in self.destination_headers:
                    missing_dest.append(row['Destination_Column'])
            
            if missing_source or missing_dest:
                error_msg = "Cannot load this mapping:\n\n"
                if missing_source:
                    error_msg += f"Missing source columns: {', '.join(set(missing_source))}\n"
                if missing_dest:
                    error_msg += f"Missing destination columns: {', '.join(set(missing_dest))}\n"
                messagebox.showerror("Incompatible Mapping", error_msg)
                return
            
            # Apply mappings
            self.apply_mappings_from_history(selected_group)
            dialog.destroy()
            messagebox.showinfo("Success", f"Loaded {len(selected_group)} column mappings from history.")
        
        ttk.Button(button_frame, text="Load Selected", command=load_selected, style='Primary.TButton').pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy, style='Secondary.TButton').pack(side=tk.RIGHT)
    
    def apply_mappings_from_history(self, mappings_group: pd.DataFrame):
        """Apply mappings from history to current interface"""
        self.clear_mappings()
        
        for _, row in mappings_group.iterrows():
            source_col = row['Source_Column']
            dest_col = row['Destination_Column']
            
            if dest_col in self.mapping_combos:
                combo = self.mapping_combos[dest_col]
                combo.set(source_col)
                
                self.column_mappings[dest_col] = source_col
                self.update_source_tree_mapping(source_col, dest_col, True)
        
        mappings_count = len(self.column_mappings)
        self.update_status(f"Loaded {mappings_count} column mappings from history")
        self.update_preview()

def create_asset_files():
    """Create logo and favicon files for the application"""
    Config.ensure_directories()
    
    # Create a simple logo using text (you can replace with actual image files)
    logo_content = """
    Excel Column Mapper Logo
    
    You can replace this file with an actual PNG logo image.
    Recommended size: 256x256 pixels
    """
    
    favicon_content = """
    Excel Column Mapper Favicon
    
    You can replace this file with an actual ICO favicon.
    Recommended size: 32x32 pixels
    """
    
    # Write placeholder files if they don't exist
    if not Config.LOGO_FILE.exists():
        with open(Config.LOGO_FILE.with_suffix('.txt'), 'w') as f:
            f.write(logo_content)
    
    if not Config.FAVICON_FILE.exists():
        with open(Config.FAVICON_FILE.with_suffix('.txt'), 'w') as f:
            f.write(favicon_content)

def main():
    """Main application entry point"""
    # Create asset files if they don't exist
    create_asset_files()
    
    # Create main window
    root = tk.Tk()
    
    # Create application
    app = ExcelColumnMapper(root)
    
    # Set up proper window closing
    def on_closing():
        root.quit()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Start the application
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        # Clean up
        if hasattr(app, 'stats_manager'):
            app.stats_manager.save_stats()

if __name__ == "__main__":
    main()