import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

class ExcelHeaderComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Column Mapper & Data Transfer")
        self.root.geometry("1400x700")
        
        # Variables to store file paths
        self.source_file_path = tk.StringVar()
        self.destination_file_path = tk.StringVar()
        
        # Variables to store column headers and data
        self.source_headers = []
        self.destination_headers = []
        self.source_df = None
        self.destination_df = None
        
        # Dictionary to store column mappings {dest_column: source_column}
        self.column_mappings = {}
        
        # Dictionary to store comboboxes for each destination column
        self.mapping_combos = {}
        
        # Store combo values for search functionality
        self.combo_values = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Source file selection
        ttk.Label(main_frame, text="Source File:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        source_frame.columnconfigure(0, weight=1)
        
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_file_path, state="readonly")
        self.source_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(source_frame, text="Browse", command=self.browse_source_file).grid(row=0, column=1)
        
        # Destination file selection
        ttk.Label(main_frame, text="Destination File:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        dest_frame = ttk.Frame(main_frame)
        dest_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        dest_frame.columnconfigure(0, weight=1)
        
        self.dest_entry = ttk.Entry(dest_frame, textvariable=self.destination_file_path, state="readonly")
        self.dest_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(dest_frame, text="Browse", command=self.browse_destination_file).grid(row=0, column=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        ttk.Button(buttons_frame, text="Load Column Headers", command=self.load_headers).pack(side=tk.LEFT, padx=(0, 10))
        self.copy_button = ttk.Button(buttons_frame, text="Copy Mapped Data", command=self.copy_mapped_data, state=tk.DISABLED)
        self.copy_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Clear Mappings", command=self.clear_mappings).pack(side=tk.LEFT, padx=(0, 10))
        self.load_history_button = ttk.Button(buttons_frame, text="Load from History", command=self.load_from_history, state=tk.DISABLED)
        self.load_history_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="View Mapping History", command=self.view_mapping_history).pack(side=tk.LEFT)
        
        # Headers display frame
        headers_frame = ttk.Frame(main_frame)
        headers_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        headers_frame.columnconfigure(0, weight=3)  # More space for source with data samples
        headers_frame.columnconfigure(1, weight=2)  # Mapping column
        headers_frame.rowconfigure(1, weight=1)
        
        # Source headers column with checkboxes and mapping info
        ttk.Label(headers_frame, text="Source File Headers", font=("Arial", 10, "bold")).grid(row=0, column=0, pady=(0, 5))
        
        source_tree_frame = ttk.Frame(headers_frame)
        source_tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        source_tree_frame.columnconfigure(0, weight=1)
        source_tree_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for source headers with checkboxes and data samples
        self.source_tree = ttk.Treeview(source_tree_frame, columns=('data_sample'), show='tree headings', height=15)
        self.source_tree.heading('#0', text='✓ Source Column', anchor=tk.W)
        self.source_tree.heading('data_sample', text='Data Sample', anchor=tk.W)
        self.source_tree.column('#0', width=400, minwidth=200, stretch=False)
        self.source_tree.column('data_sample', width=100, minwidth=50, stretch=True)
        
        source_tree_scrollbar = ttk.Scrollbar(source_tree_frame, orient=tk.VERTICAL, command=self.source_tree.yview)
        self.source_tree.configure(yscrollcommand=source_tree_scrollbar.set)
        
        self.source_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        source_tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Destination headers and mapping column
        ttk.Label(headers_frame, text="Destination Headers → Source Mapping", font=("Arial", 10, "bold")).grid(row=0, column=1, pady=(0, 5))
        
        # Create a frame with canvas for scrolling the mapping area
        self.mapping_canvas_frame = ttk.Frame(headers_frame)
        self.mapping_canvas_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.mapping_canvas_frame.columnconfigure(0, weight=1)
        self.mapping_canvas_frame.rowconfigure(0, weight=1)
        
        # Canvas and scrollbar for mapping area
        self.mapping_canvas = tk.Canvas(self.mapping_canvas_frame, bg='white')
        mapping_v_scrollbar = ttk.Scrollbar(self.mapping_canvas_frame, orient=tk.VERTICAL, command=self.mapping_canvas.yview)
        self.mapping_canvas.configure(yscrollcommand=mapping_v_scrollbar.set)
        
        self.mapping_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        mapping_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Frame inside canvas for mapping widgets
        self.mapping_frame = ttk.Frame(self.mapping_canvas)
        self.mapping_canvas_window = self.mapping_canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw")
        
        # Bind canvas resize event
        self.mapping_canvas.bind('<Configure>', self.on_canvas_configure)
        self.mapping_frame.bind('<Configure>', self.on_frame_configure)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def on_canvas_configure(self, event):
        """Handle canvas resize"""
        self.mapping_canvas.itemconfig(self.mapping_canvas_window, width=event.width)
        
    def on_frame_configure(self, event):
        """Handle frame resize"""
        self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        
    def browse_source_file(self):
        """Browse and select source Excel file from source_files folder"""
        initial_dir = os.path.join(os.getcwd(), "source_files")
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
            
        filename = filedialog.askopenfilename(
            title="Select Source Excel File",
            initialdir=initial_dir,
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.source_file_path.set(filename)
            self.status_var.set(f"Source file selected: {os.path.basename(filename)}")
            
    def browse_destination_file(self):
        """Browse and select destination Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Destination Excel File (Excel_B)",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.destination_file_path.set(filename)
            self.status_var.set(f"Destination file selected: {os.path.basename(filename)}")
            
    def read_excel_data(self, file_path):
        """Read complete Excel file data"""
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")
            
    def load_headers(self):
        """Load and display column headers from both Excel files"""
        # Clear previous data
        self.clear_source_tree()
        self.clear_mapping_widgets()
        self.column_mappings.clear()
        
        # Check if both files are selected
        if not self.source_file_path.get():
            messagebox.showerror("Error", "Please select a source Excel file")
            return
            
        if not self.destination_file_path.get():
            messagebox.showerror("Error", "Please select a destination Excel file")
            return
            
        try:
            # Load source file
            self.status_var.set("Loading source file...")
            self.root.update()
            
            self.source_df = self.read_excel_data(self.source_file_path.get())
            self.source_headers = self.source_df.columns.tolist()
            
            # Load destination file
            self.status_var.set("Loading destination file...")
            self.root.update()
            
            self.destination_df = self.read_excel_data(self.destination_file_path.get())
            self.destination_headers = self.destination_df.columns.tolist()
            
            # Populate source headers tree
            self.populate_source_tree()
                
            # Create mapping widgets for destination headers
            self.create_mapping_widgets()
            
            # Enable copy button and load history button
            self.copy_button.config(state=tk.NORMAL)
            self.load_history_button.config(state=tk.NORMAL)
            
            # Update status
            source_count = len(self.source_headers)
            dest_count = len(self.destination_headers)
            self.status_var.set(f"Files loaded - Source: {source_count} columns, Destination: {dest_count} columns")
            
            # Show summary message
            messagebox.showinfo("Success", 
                              f"Files loaded successfully!\n\n"
                              f"Source file: {source_count} columns\n"
                              f"Destination file: {dest_count} columns\n\n"
                              f"Configure column mappings and click 'Copy Mapped Data'")
                              
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load files:\n{str(e)}")
            self.status_var.set("Error loading files")
            
    def clear_source_tree(self):
        """Clear the source tree"""
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
            
    def get_data_sample(self, column_name, max_samples=1):
        """Get a representative data sample from a column"""
        try:
            if column_name not in self.source_df.columns:
                return "No data"
            
            column_data = self.source_df[column_name]
            
            # Remove null/empty values for sampling
            non_null_data = column_data.dropna()
            if len(non_null_data) == 0:
                return "All empty"
            
            # Get unique values (up to max_samples)
            unique_values = non_null_data.unique()[:max_samples]
            
            # Convert to strings and truncate if too long
            sample_strings = []
            for value in unique_values:
                value_str = str(value)
                if len(value_str) > 20:  # Truncate long values
                    value_str = value_str[:20] + "..."
                sample_strings.append(value_str)
            
            # Join samples with separator
            if len(unique_values) >= max_samples and len(non_null_data.unique()) > max_samples:
                return ", ".join(sample_strings) + ", ..."
            else:
                return ", ".join(sample_strings)
                
        except Exception as e:
            return "Error reading"

    def get_source_header_data_samples(self):
        """Get the data samples for each source header columns"""
        for header in self.source_headers:
                data_sample = self.get_data_sample(header)
                self.source_tree.insert('', 'end', iid=header, text=f"☐ {header}", values=(data_sample, ''))

    def populate_source_tree(self):
        """Populate the source tree with headers and data samples"""
        self.clear_source_tree()
        self.get_source_header_data_samples()
       
    def update_source_tree_mapping(self, source_column, dest_column, is_mapped=True):
        """Update the source tree to show mapping status"""
        
        if source_column in self.source_headers:
            if is_mapped:
                # Show as checked and display destination column
                self.source_tree.item(source_column, text=f"☑ {source_column} → {dest_column}")
            else:
                # Show as unchecked and clear destination column
                self.source_tree.item(source_column, text=f"☐ {source_column}")
                
    def clear_mapping_widgets(self):
        """Clear all mapping widgets"""
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        self.mapping_combos.clear()
        
    def create_mapping_widgets(self):
        """Create mapping widgets for destination columns"""
        self.clear_mapping_widgets()
        
        # Prepare combobox values (source headers plus empty option)
        self.combo_values = ["-- Select Source Column --"] + self.source_headers
        
        # Create label and combobox for each destination column
        for i, dest_header in enumerate(self.destination_headers):
            # Frame for each mapping row
            row_frame = ttk.Frame(self.mapping_frame)
            row_frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=2, padx=5)
            row_frame.columnconfigure(1, weight=1)
            
            # Destination column label
            dest_label = ttk.Label(row_frame, text=dest_header, width=25, anchor=tk.W)
            dest_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            
            # Arrow label
            arrow_label = ttk.Label(row_frame, text="←", font=("Arial", 12))
            arrow_label.grid(row=0, column=1, padx=5)
            
            # Source column combobox with search functionality
            combo = ttk.Combobox(row_frame, values=self.combo_values, width=25)
            combo.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(5, 0))
            combo.set("-- Select Source Column --")
            
            # Store combobox reference
            self.mapping_combos[dest_header] = combo
            
            # Bind events for search functionality
            combo.bind('<KeyRelease>', lambda event, dest=dest_header: self.on_search_keyrelease(event, dest))
            combo.bind('<FocusIn>', lambda event, dest=dest_header: self.on_combo_focus_in(event, dest))
            combo.bind('<Button-1>', lambda event, dest=dest_header: self.on_combo_click(event, dest))
            combo.bind('<<ComboboxSelected>>', lambda event, dest=dest_header: self.on_mapping_selected(event, dest))
        
        # Configure the mapping frame column weight
        self.mapping_frame.columnconfigure(0, weight=1)
        
        # Update scroll region
        self.mapping_frame.update_idletasks()
        self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        
    def on_search_keyrelease(self, event, dest_column):
        """Handle search as user types in combobox"""
        combo = self.mapping_combos[dest_column]
        typed_text = combo.get().lower()
        
        # Don't filter if it's the default text
        if typed_text == "-- select source column --":
            return
            
        # Filter source headers based on typed text
        if typed_text:
            filtered_values = ["-- Select Source Column --"] + [
                header for header in self.source_headers 
                if typed_text in header.lower()
            ]
        else:
            filtered_values = self.combo_values
            
        # Update combobox values
        combo['values'] = filtered_values
        
        # Show dropdown if there are matches and user is typing
        if len(filtered_values) > 1 and typed_text:
            combo.event_generate('<Down>')
            
    def on_combo_focus_in(self, event, dest_column):
        """Handle combobox focus - clear default text and show all options"""
        combo = self.mapping_combos[dest_column]
        current_value = combo.get()
        
        # Clear default text when focused
        if current_value == "-- Select Source Column --":
            combo.set("")
            
        # Reset to show all options
        combo['values'] = self.combo_values
        
    def on_combo_click(self, event, dest_column):
        """Handle combobox click - ensure all options are shown"""
        combo = self.mapping_combos[dest_column]
        combo['values'] = self.combo_values
        
    def on_mapping_selected(self, event, dest_column):
        """Handle when a mapping is selected from dropdown"""
        self.on_mapping_changed(dest_column)
        
    def on_mapping_changed(self, dest_column):
        """Handle mapping selection change"""
        selected_source = self.mapping_combos[dest_column].get()
        
        # Clear previous mapping for this destination column
        old_source = self.column_mappings.get(dest_column)
        if old_source:
            # Check if this source column is still used by other destinations
            still_used = any(src == old_source for dest, src in self.column_mappings.items() if dest != dest_column)
            if not still_used:
                self.update_source_tree_mapping(old_source, '', False)
        
        # Validate the selection
        if selected_source == "-- Select Source Column --" or selected_source == "":
            # Remove mapping if default or empty option is chosen
            if dest_column in self.column_mappings:
                del self.column_mappings[dest_column]
        elif selected_source in self.source_headers:
            # Store the mapping only if it's a valid source column
            self.column_mappings[dest_column] = selected_source
            # Update source tree to show this mapping
            self.update_source_tree_mapping(selected_source, dest_column, True)
        else:
            # Handle case where user typed something that doesn't match exactly
            # Check if there's a case-insensitive match
            matching_headers = [h for h in self.source_headers if h.lower() == selected_source.lower()]
            if matching_headers:
                # Use the first exact match (case-insensitive)
                exact_match = matching_headers[0]
                self.mapping_combos[dest_column].set(exact_match)
                self.column_mappings[dest_column] = exact_match
                self.update_source_tree_mapping(exact_match, dest_column, True)
            else:
                # No match found, reset to default
                self.mapping_combos[dest_column].set("-- Select Source Column --")
                if dest_column in self.column_mappings:
                    del self.column_mappings[dest_column]
            
        # Update status with current mappings count
        mappings_count = len(self.column_mappings)
        self.status_var.set(f"Column mappings configured: {mappings_count}")
        
    def clear_mappings(self):
        """Clear all column mappings"""
        # Clear comboboxes
        for combo in self.mapping_combos.values():
            combo['values'] = self.combo_values  # Reset to full list
            combo.set("-- Select Source Column --")
        
        # Clear source tree mappings (but keep data samples)
        for header in self.source_headers:
            self.update_source_tree_mapping(header, '', False)
            
        self.column_mappings.clear()
        self.status_var.set("All mappings cleared")
        
    def save_mapping_history(self, output_file_path):
        """Save mapping history to a CSV file"""
        try:
            history_file = "log\mapping_history.csv"
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Prepare history data
            history_data = []
            for dest_col, source_col in self.column_mappings.items():
                history_data.append({
                    'Timestamp': timestamp,
                    'Source_File': os.path.basename(self.source_file_path.get()),
                    'Destination_File': os.path.basename(self.destination_file_path.get()),
                    'Output_File': os.path.basename(output_file_path),
                    'Source_Column': source_col,
                    'Destination_Column': dest_col
                })
            
            # Create DataFrame
            new_history_df = pd.DataFrame(history_data)
            
            # Check if history file exists
            if os.path.exists(history_file):
                # Append to existing history
                existing_history_df = pd.read_csv(history_file)
                combined_history_df = pd.concat([existing_history_df, new_history_df], ignore_index=True)
            else:
                # Create new history file
                combined_history_df = new_history_df
            
            # Save history
            combined_history_df.to_csv(history_file, index=False)
            
            return history_file
            
        except Exception as e:
            print(f"Warning: Could not save mapping history: {str(e)}")
            return None
            
    def view_mapping_history(self):
        """View the mapping history"""
        history_file = "log\mapping_history.csv"
        
        if not os.path.exists(history_file):
            messagebox.showinfo("No History", "No mapping history found. Perform some data copies first.")
            return
            
        try:
            # Open history file with default application
            if os.name == 'nt':  # Windows
                os.startfile(history_file)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{history_file}"')
        except:
            messagebox.showinfo("History File", f"Mapping history saved in: {os.path.abspath(history_file)}")
            
    def load_from_history(self):
        """Load column mappings from history"""
        history_file = "log\mapping_history.csv"
        
        if not os.path.exists(history_file):
            messagebox.showinfo("No History", "No mapping history found. Perform some data copies first.")
            return
            
        if not self.source_headers or not self.destination_headers:
            messagebox.showwarning("Warning", "Please load both source and destination files first.")
            return
            
        try:
            # Read history file
            history_df = pd.read_csv(history_file)
            
            if history_df.empty:
                messagebox.showinfo("No History", "No mapping history found.")
                return
            
            # Create history selection dialog
            self.show_history_selection_dialog(history_df)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load history:\n{str(e)}")
            
    def show_history_selection_dialog(self, history_df):
        """Show dialog to select mapping from history"""
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Load Mapping from History")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(main_frame, text="Select a mapping configuration to load:", 
                 font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Create treeview for history records
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Group history by unique operations (timestamp + files)
        grouped_history = history_df.groupby(['Timestamp', 'Source_File', 'Destination_File', 'Output_File'])
        
        columns = ('timestamp', 'source_file', 'dest_file', 'output_file', 'mappings_count')
        history_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=10)
        
        # Define headings
        history_tree.heading('timestamp', text='Date & Time')
        history_tree.heading('source_file', text='Source File')
        history_tree.heading('dest_file', text='Destination File') 
        history_tree.heading('output_file', text='Output File')
        history_tree.heading('mappings_count', text='Mappings')
        
        # Configure column widths
        history_tree.column('timestamp', width=130)
        history_tree.column('source_file', width=150)
        history_tree.column('dest_file', width=150)
        history_tree.column('output_file', width=150)
        history_tree.column('mappings_count', width=80)
        
        # Add scrollbar
        history_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=history_tree.yview)
        history_tree.configure(yscrollcommand=history_scrollbar.set)
        
        history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate tree with grouped data
        history_records = {}
        for (timestamp, source_file, dest_file, output_file), group in grouped_history:
            mappings_count = len(group)
            item_id = history_tree.insert('', 'end', values=(
                timestamp, source_file, dest_file, output_file, mappings_count
            ))
            history_records[item_id] = group
        
        # Details frame
        details_frame = ttk.LabelFrame(main_frame, text="Mapping Details", padding="5")
        details_frame.pack(fill=tk.X, pady=(0, 10))
        
        details_text = tk.Text(details_frame, height=6, wrap=tk.WORD)
        details_scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, command=details_text.yview)
        details_text.configure(yscrollcommand=details_scrollbar.set)
        
        details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        details_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def on_history_select(event):
            """Handle history selection"""
            selection = history_tree.selection()
            if selection:
                selected_group = history_records[selection[0]]
                details_text.delete(1.0, tk.END)
                details_text.insert(tk.END, "Column Mappings:\n")
                for _, row in selected_group.iterrows():
                    details_text.insert(tk.END, f"• {row['Destination_Column']} ← {row['Source_Column']}\n")
        
        history_tree.bind('<<TreeviewSelect>>', on_history_select)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)
        
        def load_selected_mapping():
            """Load the selected mapping"""
            selection = history_tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a mapping configuration to load.")
                return
                
            selected_group = history_records[selection[0]]
            
            # Check if current files have the required columns
            missing_source = []
            missing_dest = []
            
            for _, row in selected_group.iterrows():
                if row['Source_Column'] not in self.source_headers:
                    missing_source.append(row['Source_Column'])
                if row['Destination_Column'] not in self.destination_headers:
                    missing_dest.append(row['Destination_Column'])
            
            if missing_source or missing_dest:
                error_msg = "Cannot load this mapping configuration:\n\n"
                if missing_source:
                    error_msg += f"Missing source columns: {', '.join(set(missing_source))}\n"
                if missing_dest:
                    error_msg += f"Missing destination columns: {', '.join(set(missing_dest))}\n"
                messagebox.showerror("Incompatible Mapping", error_msg)
                return
            
            # Apply the mappings
            self.apply_mappings_from_history(selected_group)
            dialog.destroy()
            messagebox.showinfo("Success", f"Loaded {len(selected_group)} column mappings from history.")
        
        ttk.Button(buttons_frame, text="Load Selected", command=load_selected_mapping).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(buttons_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)
        
    def apply_mappings_from_history(self, mappings_group):
        """Apply mappings from history to current interface"""
        # Clear current mappings
        self.clear_mappings()
        
        # Apply each mapping
        for _, row in mappings_group.iterrows():
            source_col = row['Source_Column']
            dest_col = row['Destination_Column']
            
            # Set the combobox value
            if dest_col in self.mapping_combos:
                combo = self.mapping_combos[dest_col]
                combo['values'] = self.combo_values  # Ensure full list is available
                combo.set(source_col)
                
                # Update internal mapping
                self.column_mappings[dest_col] = source_col
                
                # Update source tree visualization
                self.update_source_tree_mapping(source_col, dest_col, True)
        
        # Update status
        mappings_count = len(self.column_mappings)
        self.status_var.set(f"Loaded {mappings_count} column mappings from history")
        
    def copy_mapped_data(self):
        """Copy data from source to destination based on mappings"""
        if not self.column_mappings:
            messagebox.showwarning("Warning", "No column mappings configured. Please map at least one column.")
            return
            
        try:
            # Ask user for confirmation
            mappings_text = "\n".join([f"{dest} ← {source}" for dest, source in self.column_mappings.items()])
            confirm = messagebox.askyesno(
                "Confirm Data Copy", 
                f"This will copy data with the following mappings:\n\n{mappings_text}\n\n"
                f"Continue?"
            )
            
            if not confirm:
                return
                
            self.status_var.set("Copying data...")
            self.root.update()
            
            # Create a copy of destination dataframe
            result_df = self.destination_df.copy()
            
            # Copy data for each mapping
            copied_columns = []
            for dest_col, source_col in self.column_mappings.items():
                if source_col in self.source_df.columns:
                    # Get source data
                    source_data = self.source_df[source_col]
                    
                    # Handle case where source has more rows than destination
                    if len(source_data) > len(result_df):
                        # Extend destination dataframe if needed
                        additional_rows = len(source_data) - len(result_df)
                        empty_rows = pd.DataFrame(index=range(additional_rows), columns=result_df.columns)
                        result_df = pd.concat([result_df, empty_rows], ignore_index=True)
                    
                    # Copy the data
                    result_df[dest_col] = source_data
                    copied_columns.append(f"{dest_col} ← {source_col}")
                    
            # Ask user where to save the result
            save_path = filedialog.asksaveasfilename(
                title="Save Updated Destination File",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("All files", "*.*")
                ],
                initialfile=os.path.splitext(os.path.basename(self.destination_file_path.get()))[0] + "_updated.xlsx"
            )
            
            if save_path:
                # Save the result
                result_df.to_excel(save_path, index=False)
                
                # Save mapping history
                history_file = self.save_mapping_history(save_path)
                
                self.status_var.set(f"Data copied successfully to {os.path.basename(save_path)}")
                
                # Show success message
                success_message = (
                    f"Data copied successfully!\n\n"
                    f"Copied columns:\n" + "\n".join(copied_columns) + "\n\n"
                    f"Output file: {os.path.basename(save_path)}"
                )
                
                if history_file:
                    success_message += f"\n\nMapping history saved to: {history_file}"
                
                messagebox.showinfo("Success", success_message)
            else:
                self.status_var.set("Save cancelled")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy data:\n{str(e)}")
            self.status_var.set("Error copying data")

def main():
    # Create the main window
    root = tk.Tk()
    
    # Create the application
    app = ExcelHeaderComparator(root)
    
    # Start the GUI event loop
    root.mainloop()

if __name__ == "__main__":
    main()