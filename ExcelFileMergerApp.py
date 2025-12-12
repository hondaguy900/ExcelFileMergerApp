import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Merger")
        self.root.geometry("1100x650")
        
        # Initialize variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.duplicate_handling = tk.StringVar(value="keep_all")
        self.file1_sheet = tk.StringVar()
        self.file2_sheet = tk.StringVar()
        self.file1_header_row = tk.IntVar(value=1)  # 1-based index for display
        self.file2_header_row = tk.IntVar(value=1)  # 1-based index for display
        self.superior_file = tk.IntVar(value=1)  # 1 = File 1, 2 = File 2
        self.merge_type = tk.StringVar(value="outer")  # merge type options

        # Track the merge option radio buttons
        self.merge_option_buttons = []
        
        # For storing data
        self.file1_columns = []
        self.file2_columns = []
        self.file1_sheets = []
        self.file2_sheets = []
        
        # Create and arrange widgets
        self.create_widgets()
        
    def create_widgets(self):
        # Create main frame with scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)
        
        # Add canvas for scrolling
        canvas = tk.Canvas(main_frame)
        canvas.pack(side="left", fill="both", expand=True)
        
        # Add scrollbar
        # scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        # scrollbar.pack(side="right", fill="y")
        
        # Configure canvas
        # canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Create frame for content
        content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        # Binding mousewheel to scroll
        def _on_mousewheel(event):
           pass
        
        # Create left and right panes
        left_pane = ttk.Frame(content_frame)
        left_pane.grid(row=0, column=0, sticky="nw", padx=10, pady=5)  # Reduced vertical padding
        
        # Add separator between panes
        ttk.Separator(content_frame, orient="vertical").grid(row=0, column=1, sticky="ns", padx=10, pady=5)
        
        right_pane = ttk.Frame(content_frame)
        right_pane.grid(row=0, column=2, sticky="ne", padx=10, pady=5)  # Reduced vertical padding
        
        # Configure grid weights for content_frame
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(2, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)
        
        # Create frames for left pane with reduced padding
        input_frame = ttk.LabelFrame(left_pane, text="Input Files")
        input_frame.pack(fill="x", expand="no", padx=0, pady=3)
        
        sheet_frame = ttk.LabelFrame(left_pane, text="Sheet Selection")
        sheet_frame.pack(fill="x", expand="no", padx=0, pady=3)
        
        options_frame = ttk.LabelFrame(left_pane, text="Merge Options")
        options_frame.pack(fill="x", expand="no", padx=0, pady=3)
        
        output_frame = ttk.LabelFrame(left_pane, text="Output File")
        output_frame.pack(fill="x", expand="no", padx=0, pady=3)
        
        # Create mapping frame for right pane
        column_frame = ttk.LabelFrame(right_pane, text="Column Mapping")
        column_frame.pack(fill="both", expand="yes", padx=0, pady=10)
        
        # File 1 selection - reduced padding
        ttk.Label(input_frame, text="First Excel File:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(input_frame, textvariable=self.file1_path, width=45).grid(column=1, row=0, padx=5, pady=2)
        ttk.Button(input_frame, text="Browse...", command=self.browse_file1).grid(column=2, row=0, padx=5, pady=2)
        
        # File 2 selection - reduced padding
        ttk.Label(input_frame, text="Second Excel File:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(input_frame, textvariable=self.file2_path, width=45).grid(column=1, row=1, padx=5, pady=2)
        ttk.Button(input_frame, text="Browse...", command=self.browse_file2).grid(column=2, row=1, padx=5, pady=2)
        
        # Configure grid for input_frame to center the load button
        input_frame.columnconfigure(0, weight=1)
        input_frame.columnconfigure(1, weight=10)
        input_frame.columnconfigure(2, weight=1)
        
        # Load files button - centered
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=3)
        ttk.Button(button_frame, text="Load Files", command=self.load_files, width=15).pack(pady=2)
        
        # Sheet selection with reduced padding
        ttk.Label(sheet_frame, text="File 1 Sheet:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=2)
        self.file1_sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.file1_sheet, width=25, state="readonly")
        self.file1_sheet_combo.grid(column=1, row=0, padx=5, pady=2)
        
        ttk.Label(sheet_frame, text="Header Row:").grid(column=2, row=0, sticky=tk.W, padx=5, pady=2)
        ttk.Spinbox(sheet_frame, from_=1, to=10, textvariable=self.file1_header_row, width=5).grid(column=3, row=0, padx=5, pady=2)
        
        ttk.Label(sheet_frame, text="File 2 Sheet:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=2)
        self.file2_sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.file2_sheet, width=25, state="readonly")
        self.file2_sheet_combo.grid(column=1, row=1, padx=5, pady=2)
        
        ttk.Label(sheet_frame, text="Header Row:").grid(column=2, row=1, sticky=tk.W, padx=5, pady=2)
        ttk.Spinbox(sheet_frame, from_=1, to=10, textvariable=self.file2_header_row, width=5).grid(column=3, row=1, padx=5, pady=2)
        
        # Configure grid for sheet_frame to center the load button
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.columnconfigure(1, weight=10)
        sheet_frame.columnconfigure(2, weight=1)
        sheet_frame.columnconfigure(3, weight=1)
        
        # Load columns button - centered
        button_frame2 = ttk.Frame(sheet_frame)
        button_frame2.grid(row=2, column=0, columnspan=4, pady=3)
        ttk.Button(button_frame2, text="Load Columns", command=self.load_columns, width=15).pack(pady=2)
        
        # Column selection area
        self.column_selection_frame = ttk.Frame(column_frame)
        self.column_selection_frame.pack(fill="both", expand="yes", padx=5, pady=5)
        
        # Duplicate handling options section
        ttk.Label(options_frame, text="Duplicate Handling:", font=("TkDefaultFont", 9)).grid(
            column=0, row=0, sticky=tk.W, padx=5, pady=5)
        
        # Radio buttons for duplicate handling options - declare first
        duplicate_options = [
            ("Keep all duplicates", "keep_all"),
            ("Keep first occurrence", "first"),
            ("Keep last occurrence", "last"),
            ("Flag duplicates", "flag")
        ]
        
        # Duplicate handling options section - reduced vertical padding
        for i, (text, value) in enumerate(duplicate_options):
            ttk.Radiobutton(
                options_frame, 
                text=text, 
                value=value, 
                variable=self.duplicate_handling
            ).grid(column=1, row=i, sticky=tk.W, padx=20, pady=1)  # Reduced padding
        
        # Add separator after duplicate handling section
        ttk.Separator(options_frame, orient="horizontal").grid(
            column=0, row=4, columnspan=2, sticky="ew", padx=5, pady=8)
        
        # Superior file section with reduced padding
        ttk.Label(options_frame, text="Superior File:", font=("TkDefaultFont", 9)).grid(
            column=0, row=5, sticky=tk.W, padx=5, pady=2)
        
        # Superior file radio buttons - on the same line
        radio_frame = ttk.Frame(options_frame)
        radio_frame.grid(column=1, row=5, sticky=tk.W, padx=20, pady=1)
        
        # Add trace to superior_file variable to update merge type labels when changed
        self.superior_file.trace_add("write", self.update_merge_type_labels)
        
        ttk.Radiobutton(
            radio_frame, 
            text="File 1", 
            value=1, 
            variable=self.superior_file
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Radiobutton(
            radio_frame, 
            text="File 2", 
            value=2, 
            variable=self.superior_file
        ).pack(side=tk.LEFT)
        
        ttk.Label(options_frame, text="(Values from this file will take precedence for duplicates)").grid(
            column=1, row=6, sticky=tk.W, padx=20, pady=(0, 2))  # Reduced bottom padding
        
        # Add separator after superior file section
        ttk.Separator(options_frame, orient="horizontal").grid(
            column=0, row=7, columnspan=2, sticky="ew", padx=5, pady=8)
        
        # Create a frame to hold the merge type options
        self.merge_type_frame = ttk.Frame(options_frame)
        self.merge_type_frame.grid(column=1, row=8, rowspan=4, sticky=tk.W, padx=20, pady=1)
        
        # Merge type selection - reduced padding
        ttk.Label(options_frame, text="Merge Type:", font=("TkDefaultFont", 9,)).grid(
            column=0, row=8, sticky=tk.W, padx=5, pady=2)
        
        # Create the merge type radio buttons with initial labels
        self.create_merge_type_buttons()
        
        # Output file selection with reduced padding
        ttk.Label(output_frame, text="Output Excel File:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(output_frame, textvariable=self.output_path, width=45).grid(column=1, row=0, padx=5, pady=2)
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).grid(column=2, row=0, padx=5, pady=2)
        
        # Merge button - centered and with reduced padding
        button_frame3 = ttk.Frame(left_pane)
        button_frame3.pack(pady=5)  # Reduced padding
        ttk.Button(button_frame3, text="Merge Files", command=self.merge_files, width=20).pack(pady=2)  # Reduced padding
    
    def create_merge_type_buttons(self):
        """Create the merge type radio buttons with appropriate labels"""
        # Clear any existing buttons first
        for widget in self.merge_type_frame.winfo_children():
            widget.destroy()
        
        self.merge_option_buttons = []
        
        # Get labels based on current superior file selection
        merge_options = self.get_merge_options()
        
        # Create radio buttons for merge options
        for i, (text, value) in enumerate(merge_options):
            radio_btn = ttk.Radiobutton(
                self.merge_type_frame, 
                text=text, 
                value=value, 
                variable=self.merge_type
            )
            radio_btn.grid(row=i, column=0, sticky=tk.W, pady=1)
            self.merge_option_buttons.append((radio_btn, value))
    
    def update_merge_type_labels(self, *args):
        """Update merge type radio button labels when superior file changes"""
        merge_options = self.get_merge_options()
        
        # Update the text of each radio button
        for i, ((radio_btn, value), (text, _)) in enumerate(zip(self.merge_option_buttons, merge_options)):
            radio_btn.config(text=text)
    
    def get_merge_options(self):
        """Return merge options with labels based on current superior file selection"""
        superior = self.superior_file.get()
        
        if superior == 1:  # File 1 is superior
            return [
                ("Keep all rows from both files", "outer"),
                ("Keep all rows from File 1 (superior)", "left"),
                ("Keep all rows from File 2 (inferior)", "right"),
                ("Only keep rows that match in both files", "inner")
            ]
        else:  # File 2 is superior
            return [
                ("Keep all rows from both files", "outer"),
                ("Keep all rows from File 2 (superior)", "left"),
                ("Keep all rows from File 1 (inferior)", "right"),
                ("Only keep rows that match in both files", "inner")
            ]
        
    def browse_file1(self):
        filename = filedialog.askopenfilename(
            title="Select first Excel file",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.file1_path.set(filename)
            
    def browse_file2(self):
        filename = filedialog.askopenfilename(
            title="Select second Excel file",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.file2_path.set(filename)
            
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save merged Excel file as",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.output_path.set(filename)
            
    def load_files(self):
        """Load Excel files and populate sheet selection dropdowns"""
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showerror("Error", "Please select both Excel files first.")
            return
            
        try:
            import pandas as pd
            
            # Load workbooks to get sheet names
            self.file1_sheets = pd.ExcelFile(self.file1_path.get()).sheet_names
            self.file2_sheets = pd.ExcelFile(self.file2_path.get()).sheet_names
            
            # Update comboboxes
            self.file1_sheet_combo['values'] = self.file1_sheets
            self.file2_sheet_combo['values'] = self.file2_sheets
            
            # Set default values
            if self.file1_sheets:
                self.file1_sheet.set(self.file1_sheets[0])
            if self.file2_sheets:
                self.file2_sheet.set(self.file2_sheets[0])
                
            # Set header row to 1 (1-based for display)
            self.file1_header_row.set(1)
            self.file2_header_row.set(1)
            
            messagebox.showinfo("Success", 
                              f"File 1 sheets: {', '.join(self.file1_sheets)}\n"
                              f"File 2 sheets: {', '.join(self.file2_sheets)}\n\n"
                              f"Please select the sheets to use and the header row positions, then click 'Load Columns'.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel files: {str(e)}")
    
    def load_columns(self):
        """Load columns from selected sheets and header rows"""
        # Clear previous widgets
        for widget in self.column_selection_frame.winfo_children():
            widget.destroy()
            
        # Validate sheet selections
        if not self.file1_sheet.get() or not self.file2_sheet.get():
            messagebox.showerror("Error", "Please select sheets from both Excel files.")
            return
            
        try:
            import pandas as pd
            
            # Load data from the selected sheets with specified header rows
            # Convert from 1-based (UI) to 0-based (pandas)
            df1 = pd.read_excel(self.file1_path.get(), sheet_name=self.file1_sheet.get(), header=self.file1_header_row.get()-1)
            df2 = pd.read_excel(self.file2_path.get(), sheet_name=self.file2_sheet.get(), header=self.file2_header_row.get()-1)
            
            self.file1_columns = df1.columns.tolist()
            self.file2_columns = df2.columns.tolist()
            
            # Create the column mapping interface
            self.display_column_mapping_interface()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load columns: {str(e)}")
    
    def display_column_mapping_interface(self):
        # Create a simple mapping interface showing all columns from both files
        ttk.Label(self.column_selection_frame, text="Map columns between files:").grid(
            column=0, row=0, columnspan=4, sticky=tk.W, padx=5, pady=5
        )
        
        # Headers
        ttk.Label(self.column_selection_frame, text="Use for\nMatching", font="TkDefaultFont 9 bold").grid(
            column=0, row=1, padx=5, pady=5)
        ttk.Label(self.column_selection_frame, text="File 1 Column", font="TkDefaultFont 9 bold").grid(
            column=1, row=1, padx=5, pady=5)
        ttk.Label(self.column_selection_frame, text="Maps to", font="TkDefaultFont 9 bold").grid(
            column=2, row=1, padx=5, pady=5)
        ttk.Label(self.column_selection_frame, text="File 2 Column", font="TkDefaultFont 9 bold").grid(
            column=3, row=1, padx=5, pady=5)
        
        # Create rows of mapping options
        self.mapping_rows = []
        
        # Determine number of mapping rows - start with min 5, or up to 10 based on column counts
        num_rows = max(5, min(10, max(len(self.file1_columns), len(self.file2_columns))))
        
        for i in range(num_rows):
            use_var = tk.BooleanVar(value=False)
            file1_var = tk.StringVar()
            file2_var = tk.StringVar()
            
            # If we have columns for this row, set them as defaults
            if i < len(self.file1_columns):
                file1_var.set(self.file1_columns[i])
            if i < len(self.file2_columns):
                file2_var.set(self.file2_columns[i])
                
            # If columns happen to have the same name in same position, select by default
            if (i < len(self.file1_columns) and i < len(self.file2_columns) and 
                self.file1_columns[i] == self.file2_columns[i]):
                use_var.set(True)
            
            self.mapping_rows.append((use_var, file1_var, file2_var))
            
            # Create the row widgets
            ttk.Checkbutton(self.column_selection_frame, variable=use_var).grid(
                column=0, row=i+2, padx=5, pady=2)
            
            file1_combo = ttk.Combobox(self.column_selection_frame, textvariable=file1_var, 
                                      values=self.file1_columns, width=25)
            file1_combo.grid(column=1, row=i+2, padx=5, pady=2)
            
            ttk.Label(self.column_selection_frame, text="➔").grid(
                column=2, row=i+2, padx=5, pady=2)
            
            file2_combo = ttk.Combobox(self.column_selection_frame, textvariable=file2_var,
                                      values=self.file2_columns, width=25)
            file2_combo.grid(column=3, row=i+2, padx=5, pady=2)
            
        # Configure grid to center the button
        self.column_selection_frame.columnconfigure(0, weight=1)
        self.column_selection_frame.columnconfigure(1, weight=10)
        self.column_selection_frame.columnconfigure(2, weight=1)
        self.column_selection_frame.columnconfigure(3, weight=10)
        
        # Add button to add more mapping rows - centered
        button_frame = ttk.Frame(self.column_selection_frame)
        button_frame.grid(row=num_rows+2, column=0, columnspan=4, pady=10)
        ttk.Button(
            button_frame,
            text="Add More Mapping Rows",
            command=self.add_mapping_row,
            width=24
        ).pack(pady=5)

    def add_mapping_row(self):
        # Add another row to the mapping interface
        row_index = len(self.mapping_rows) + 2  # +2 for headers
        
        use_var = tk.BooleanVar(value=False)
        file1_var = tk.StringVar()
        file2_var = tk.StringVar()
        
        self.mapping_rows.append((use_var, file1_var, file2_var))
        
        # Create the row widgets
        ttk.Checkbutton(self.column_selection_frame, variable=use_var).grid(
            column=0, row=row_index, padx=5, pady=2)
        
        file1_combo = ttk.Combobox(self.column_selection_frame, textvariable=file1_var, 
                                  values=self.file1_columns, width=25)
        file1_combo.grid(column=1, row=row_index, padx=5, pady=2)
        
        ttk.Label(self.column_selection_frame, text="➔").grid(
            column=2, row=row_index, padx=5, pady=2)
        
        file2_combo = ttk.Combobox(self.column_selection_frame, textvariable=file2_var,
                                  values=self.file2_columns, width=25)
        file2_combo.grid(column=3, row=row_index, padx=5, pady=2)
        
        # Move the "Add More Mapping Rows" button down
        for widget in self.column_selection_frame.grid_slaves():
            if widget.winfo_class() == 'TFrame':  # Look for the button frame
                widget.grid_forget()
                widget.grid(row=row_index+1, column=0, columnspan=4, pady=10)
            
    def merge_files(self):
        # Validate inputs
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showerror("Error", "Please select both Excel files.")
            return
            
        if not self.output_path.get():
            messagebox.showerror("Error", "Please specify the output file.")
            return
            
        if not self.file1_sheet.get() or not self.file2_sheet.get():
            messagebox.showerror("Error", "Please select sheets from both Excel files.")
            return
            
        # Check if we have mapping rows
        if not hasattr(self, 'mapping_rows') or not self.mapping_rows:
            messagebox.showerror("Error", "Please load columns first.")
            return
            
        # Get selected columns for mapping
        selected_mappings = []
        for use_var, file1_var, file2_var in self.mapping_rows:
            if use_var.get() and file1_var.get() and file2_var.get():
                selected_mappings.append((file1_var.get(), file2_var.get()))
        
        if not selected_mappings:
            messagebox.showerror("Error", "Please select at least one column mapping for merging.")
            return
            
        try:
            # Load data from specific sheets and header rows
            # Convert from 1-based (UI) to 0-based (pandas)
            df1 = pd.read_excel(
                self.file1_path.get(), 
                sheet_name=self.file1_sheet.get(), 
                header=self.file1_header_row.get()-1
            )
            
            df2 = pd.read_excel(
                self.file2_path.get(), 
                sheet_name=self.file2_sheet.get(), 
                header=self.file2_header_row.get()-1
            )
            
            # Create renamed temporary dataframes for the merge
            df1_temp = df1.copy()
            df2_temp = df2.copy()
            
            # Rename columns in each dataframe to a common name for merging
            merge_columns = []
            rename_dict1 = {}
            rename_dict2 = {}
            
            # Keep track of original column names for each file
            file1_orig_cols = {}
            file2_orig_cols = {}
            
            for i, (col1, col2) in enumerate(selected_mappings):
                merge_col = f"merge_col_{i}"
                
                rename_dict1[col1] = merge_col
                rename_dict2[col2] = merge_col
                merge_columns.append(merge_col)
                
                # Store original column names
                file1_orig_cols[merge_col] = col1
                file2_orig_cols[merge_col] = col2
            
            # Apply the renames
            df1_temp = df1_temp.rename(columns=rename_dict1)
            df2_temp = df2_temp.rename(columns=rename_dict2)
            
            # Set up the merge dataframes in a consistent way, not swapping them
            left_df = df1_temp
            right_df = df2_temp
            left_suffix = '_file1'
            right_suffix = '_file2'
            
            # Determine the correct 'how' parameter based on merge type and superior file
            merge_type_value = self.merge_type.get()
            superior_file_value = self.superior_file.get()
            
            # Adjust how parameter based on merge type and superior file
            if merge_type_value == "outer" or merge_type_value == "inner":
                # These types work the same regardless of superior file
                how_param = merge_type_value
            elif merge_type_value == "left":
                # Keep all rows from superior file
                how_param = "left" if superior_file_value == 1 else "right"
            elif merge_type_value == "right":
                # Keep all rows from inferior file
                how_param = "right" if superior_file_value == 1 else "left"
            
            # Perform merge on renamed columns
            merged_df = pd.merge(
                left_df, right_df, 
                on=merge_columns, 
                how=how_param,  # Use adjusted merge type
                suffixes=(left_suffix, right_suffix), 
                indicator=True
            )
            
            # Handle duplicates based on user selection
            duplicate_option = self.duplicate_handling.get()
            
            if duplicate_option == "keep_all":
                # Default behavior - keep all records
                pass
            elif duplicate_option == "first":
                # Keep only the first occurrence of each combination of matching columns
                merged_df = merged_df.drop_duplicates(subset=merge_columns, keep='first')
            elif duplicate_option == "last":
                # Keep only the last occurrence of each combination of matching columns
                merged_df = merged_df.drop_duplicates(subset=merge_columns, keep='last')
            elif duplicate_option == "flag":
                # Flag duplicates by adding a column
                # Create a column that counts occurrences of each unique combination
                merged_df['duplicate_count'] = merged_df.groupby(merge_columns)[merge_columns[0]].transform('count')
                merged_df['is_duplicate'] = merged_df['duplicate_count'] > 1
            
            # Count duplicates if applicable
            duplicate_count = 0
            if duplicate_option == "flag":
                duplicate_count = merged_df['is_duplicate'].sum()
            
            # Add a column indicating which file(s) the data came from
            merged_df['source'] = merged_df['_merge'].map({
                'left_only': 'File 1 Only',
                'right_only': 'File 2 Only',
                'both': 'Both Files'
            })
            
            # Rename merge columns back to original names
            # Use the superior file's column names
            rename_back = {}
            for merge_col in merge_columns:
                if self.superior_file.get() == 1:
                    rename_back[merge_col] = file1_orig_cols[merge_col]
                else:
                    rename_back[merge_col] = file2_orig_cols[merge_col]
                    
            merged_df = merged_df.rename(columns=rename_back)
            
            # Save to output file
            output_file = self.output_path.get()
            merged_df.to_excel(output_file, index=False)
            
            # Create a dictionary for duplicate handling descriptions
            dup_options = {
                "keep_all": "Kept all records",
                "first": "Kept first occurrence", 
                "last": "Kept last occurrence", 
                "flag": "Flagged duplicates"
            }
            
            # Create a dictionary for superior file descriptions
            superior_options = {
                1: "File 1 (values from File 1 take precedence)",
                2: "File 2 (values from File 2 take precedence)"
            }
            
            # Create a dictionary for merge type descriptions that matches the dynamic labels
            merge_type_options = self.get_merge_type_descriptions()
            
            # Create success message
            success_msg = f"Files merged successfully!\n"
            success_msg += f"Number of rows in file 1: {len(df1)}\n"
            success_msg += f"Number of rows in file 2: {len(df2)}\n"
            success_msg += f"Number of rows in merged file: {len(merged_df)}\n"
            success_msg += f"Number of columns: {len(merged_df.columns)}\n"
            success_msg += f"Duplicate handling: {dup_options[self.duplicate_handling.get()]}\n"
            success_msg += f"Superior file: {superior_options[self.superior_file.get()]}\n"
            success_msg += f"Merge type: {merge_type_options[self.merge_type.get()]}\n"
            success_msg += f"Output file: {output_file}"
            
            
            
            # Add duplicate count if flagging was used
            if duplicate_option == "flag":
                success_msg += f"\nNumber of duplicates found: {duplicate_count}"
                
            # Create a custom dialog with "Open File" button
            self.show_success_with_open_button(success_msg, output_file)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge files: {str(e)}\n\nError details: {repr(e)}")
            # Print stack trace for debugging
            import traceback
            traceback.print_exc()
    
    def get_merge_type_descriptions(self):
        """Return merge type descriptions based on current superior file selection"""
        superior = self.superior_file.get()
        
        if superior == 1:  # File 1 is superior
            return {
                "outer": "All rows from both files",
                "left": "All rows from File 1 (superior)",
                "right": "All rows from File 2 (inferior)",
                "inner": "Only rows that match in both files"
            }
        else:  # File 2 is superior
            return {
                "outer": "All rows from both files",
                "left": "All rows from File 2 (superior)",
                "right": "All rows from File 1 (inferior)",
                "inner": "Only rows that match in both files"
            }
            
    def show_success_with_open_button(self, message, filepath):
        """Show success message with a button to open the output file"""
        # Create a custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Merge Successful")
        dialog.geometry("400x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)  # Set as transient window to parent
        dialog.grab_set()  # Make dialog modal
        
        # Add icon if available
        try:
            dialog.iconbitmap(default=self.root.iconbitmap())
        except:
            pass  # If no icon is set, ignore
            
        # Message area
        message_frame = ttk.Frame(dialog, padding=10)
        message_frame.pack(fill="both", expand=True)
        
        # Message label with scrolling
        msg_container = ttk.Frame(message_frame)
        msg_container.pack(fill="both", expand=True)
        
        # Scrollable text widget for the message
        from tkinter import scrolledtext
        text_area = scrolledtext.ScrolledText(
            msg_container, 
            wrap=tk.WORD, 
            width=50, 
            height=15,
            font=("TkDefaultFont", 9)
        )
        text_area.pack(fill="both", expand=True, padx=5, pady=5)
        text_area.insert(tk.INSERT, message)
        text_area.config(state="disabled")  # Read-only
        
        # Button frame
        button_frame = ttk.Frame(dialog, padding=(10, 0, 10, 10))
        button_frame.pack(fill="x")
        
        # Function to open the Excel file
        def open_excel_file():
            try:
                import os
                import platform
                import subprocess
                
                # Normalize path
                filepath_norm = os.path.normpath(filepath)
                
                # Open file based on operating system
                if platform.system() == 'Windows':
                    os.startfile(filepath_norm)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.call(['open', filepath_norm])
                else:  # Linux and other OS
                    subprocess.call(['xdg-open', filepath_norm])
                    
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {str(e)}")
        
        # Function to close the dialog
        def close_dialog():
            dialog.destroy()
        
        # Add buttons
        ttk.Button(
            button_frame, 
            text="Open in Excel", 
            command=open_excel_file,
            width=15
        ).pack(side="left", padx=5)
        
        ttk.Button(
            button_frame, 
            text="OK", 
            command=close_dialog,
            width=10
        ).pack(side="right", padx=5)
        
        # Center the dialog on parent window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Make the dialog wait for user input
        dialog.wait_window(dialog)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
	
# Version: 2025.04.10.001