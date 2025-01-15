# pip install pandas, docxtpl

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
import re
import os
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import string



class ExcelToWordMapper:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to Word Mapper")
        self.root.geometry("800x600")
        self.excel_path = tk.StringVar()
        self.word_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.filename_pattern = tk.StringVar(value="doc-{}-{}.docx")  # Default pattern
        self.excel_data = None
        self.word_template = None
        self.template_fields = []
        self.mapping = {}
        self.filename_columns = []
        
        self.create_ui()
    
    def create_ui(self):
        file_frame = ttk.LabelFrame(self.root, text="File Selection", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2)        
        ttk.Label(file_frame, text="Word Template:").grid(row=1, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.word_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_word).grid(row=1, column=2)
        ttk.Label(file_frame, text="Output Directory:").grid(row=2, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=2, column=2)
        ttk.Button(file_frame, text="Load Files", command=self.load_files).grid(row=3, column=1, pady=10)
        
        self.filename_frame = ttk.LabelFrame(self.root, text="Filename Pattern", padding=10)
        self.filename_frame.pack(fill="x", padx=10, pady=5)
        
        self.mapping_frame = ttk.LabelFrame(self.root, text="Field Mapping", padding=10)
        self.mapping_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.progress_frame = ttk.Frame(self.root, padding=10)
        self.progress_frame.pack(fill="x", padx=10)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=5)        
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_label.pack()
        ttk.Button(self.root, text="Generate Documents", command=self.generate_documents).pack(pady=10)
    
    def create_filename_ui(self, excel_columns):
        for widget in self.filename_frame.winfo_children():
            widget.destroy()
            
        ttk.Label(self.filename_frame, text="Select columns for filename pattern:").pack(anchor="w", pady=5)
        col_frame = ttk.Frame(self.filename_frame)
        col_frame.pack(fill="x", pady=5)

        self.filename_columns = []
        for i in range(2):
            combo = ttk.Combobox(col_frame, values=[''] + excel_columns, width=20)
            combo.grid(row=0, column=i, padx=5)
            self.filename_columns.append(combo)

        preview_frame = ttk.Frame(self.filename_frame)
        preview_frame.pack(fill="x", pady=5)
        ttk.Label(preview_frame, text="Preview: ").pack(side="left")
        self.preview_label = ttk.Label(preview_frame, text="doc-value1-value2-1.docx")
        self.preview_label.pack(side="left")
        
        for combo in self.filename_columns:
            combo.bind('<<ComboboxSelected>>', self.update_filename_preview)
    
    def update_filename_preview(self, event=None):
        selected_columns = [combo.get() for combo in self.filename_columns if combo.get()]
        if selected_columns:
            preview = f"doc-{'-'.join(['value' + str(i+1) for i in range(len(selected_columns))])}-1.docx"
            self.preview_label.config(text=preview)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)
    
    def browse_word(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.word_path.set(filename)
    
    def browse_output(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)
    
    def sanitize_filename(self, filename):
        # Remove invalid characters and replace spaces with underscores
        valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
        filename = ''.join(c for c in filename if c in valid_chars)
        filename = filename.replace(' ', '_')
        return filename
    
    def extract_template_fields(self, template):
        """Extract template fields from the Word document using regex pattern matching."""
        try:
            variables = template.undeclared_template_variables
            if variables:
                return list(variables)
            else:
                template_xml = template.get_xml()
                fields = list(set(re.findall(r'\{\{(\w+)\}\}', template_xml)))
                return fields
        except Exception as e:
            messagebox.showerror("Error", f"Error extracting template fields: {str(e)}")
            return []

    def load_files(self):
        try:
            if not self.excel_path.get() or not self.word_path.get():
                messagebox.showwarning("Warning", "Please select both Excel and Word files.")
                return

            # Load Excel
            self.excel_data = pd.read_excel(self.excel_path.get())
            excel_columns = list(self.excel_data.columns)
            
            # Create filename pattern UI
            self.create_filename_ui(excel_columns)
            
            # Load Word template and extract fields
            self.word_template = DocxTemplate(self.word_path.get())
            self.template_fields = self.extract_template_fields(self.word_template)
            
            if not self.template_fields:
                messagebox.showwarning("Warning", "No template fields found in the Word document.\nMake sure fields are formatted as {{field_name}}")
                return
            
            # Clear existing mapping widgets
            for widget in self.mapping_frame.winfo_children():
                widget.destroy()
            
            # Create mapping UI
            ttk.Label(self.mapping_frame, text="Template Field").grid(row=0, column=0, padx=5, pady=5)
            ttk.Label(self.mapping_frame, text="Excel Column").grid(row=0, column=1, padx=5, pady=5)
            
            self.mapping_vars = {}
            for i, field in enumerate(self.template_fields, 1):
                ttk.Label(self.mapping_frame, text=field).grid(row=i, column=0, padx=5, pady=2)
                var = tk.StringVar()
                self.mapping_vars[field] = var
                combo = ttk.Combobox(self.mapping_frame, textvariable=var, values=[''] + excel_columns)
                combo.grid(row=i, column=1, padx=5, pady=2)
                
                if field in excel_columns:
                    var.set(field)
            
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "The Excel file is empty.")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading files: {str(e)}")

    def generate_single_document(self, template_path, context, output_path):
        """Generate a single document with proper cleanup."""
        try:
            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(output_path)
            
            try:
                verify_doc = Document(output_path)
                verify_doc = None
            except PackageNotFoundError:
                with open(output_path, 'rb') as broken_file:
                    content = broken_file.read()
                with open(output_path, 'wb') as fixed_file:
                    fixed_file.write(content)
            
            return True
        except Exception as e:
            print(f"Error generating document: {str(e)}")
            return False
    
    def generate_documents(self):
        try:
            if not self.output_dir.get():
                messagebox.showwarning("Warning", "Please select an output directory.")
                return

            filename_cols = [combo.get() for combo in self.filename_columns if combo.get()]
            if not filename_cols:
                messagebox.showwarning("Warning", "Please select at least one column for filename pattern.")
                return

            # Create mapping
            self.mapping = {field: var.get() for field, var in self.mapping_vars.items()}
            
            # Validate mapping
            unmapped_fields = [field for field, col in self.mapping.items() if not col]
            if unmapped_fields:
                unmapped_list = ", ".join(unmapped_fields)
                messagebox.showwarning("Warning", f"Please map the following template fields:\n{unmapped_list}")
                return
            
            total_rows = len(self.excel_data)
            successful_generations = 0
            
            # Generate documents for each row
            for index, row in self.excel_data.iterrows():
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"Processing document {index + 1} of {total_rows}")
                self.root.update_idletasks()
                
                # Create context dictionary for template
                context = {field: str(row[col]) for field, col in self.mapping.items()}
                
                # Generate filename
                filename_parts = [str(row[col]) for col in filename_cols]
                filename_parts = [self.sanitize_filename(part) for part in filename_parts]
                filename = f"doc-{'-'.join(filename_parts)}-{index+1}.docx"
                output_path = os.path.join(self.output_dir.get(), filename)
                
                # Generate document
                if self.generate_single_document(self.word_path.get(), context, output_path):
                    successful_generations += 1
            
            # Reset progress bar and update label
            self.progress_var.set(0)
            self.progress_label.config(text=f"Generated {successful_generations} of {total_rows} documents successfully!")
            
            if successful_generations == total_rows:
                messagebox.showinfo("Success", f"Successfully generated all {total_rows} documents!")
            else:
                messagebox.showwarning("Warning", 
                    f"Generated {successful_generations} of {total_rows} documents.\n"
                    "Some documents may need to be checked for errors.")
            
        except Exception as e:
            self.progress_label.config(text="Error during generation")
            messagebox.showerror("Error", f"Error generating documents: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToWordMapper(root)
    root.mainloop()