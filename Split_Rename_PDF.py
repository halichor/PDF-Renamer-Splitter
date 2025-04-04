import fitz  # PyMuPDF 
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES  # Import tkinterdnd2
import re
import ttkbootstrap as ttkb


class SplitPDFApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF Splitter & Renamer")
        self.geometry("740x550")
        self.configure(bg="#f7f7f7")

        self.pdf_path = None
        self.csv_path = None
        self.output_folder = None
        self.filename_format = None
        self.output_folder_var = tk.StringVar()

        # Configure column weights for even layout
        self.grid_columnconfigure(0, weight=1, uniform="equal")
        self.grid_columnconfigure(1, weight=1, uniform="equal")

        # Header
        self.header_label = ttkb.Label(self, text="PDF Splitter & Renamer", font=("Helvetica", 18, "bold"))
        self.header_label.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky="w")

        # PDF Drop Area
        self.pdf_label = ttkb.Label(self, text="Drag and Drop PDF File here", font=("Helvetica", 12))
        self.pdf_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.pdf_drop = tk.Label(self, text="No file selected", relief="solid", width=35, height=5)
        self.pdf_drop.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        self.pdf_drop.drop_target_register(DND_FILES)
        self.pdf_drop.dnd_bind('<<Drop>>', self.on_pdf_drop)

        # CSV Drop Area
        self.csv_label = ttkb.Label(self, text="Drag and Drop CSV File here", font=("Helvetica", 12))
        self.csv_label.grid(row=1, column=1, padx=20, pady=5, sticky="w")

        self.csv_drop = tk.Label(self, text="No file selected", relief="solid", width=35, height=5)
        self.csv_drop.grid(row=2, column=1, padx=20, pady=5, sticky="ew")
        self.csv_drop.drop_target_register(DND_FILES)
        self.csv_drop.dnd_bind('<<Drop>>', self.on_csv_drop)

        # Pages per split
        self.split_size_label = ttkb.Label(self, text="Pages per split file:", font=("Helvetica", 12))
        self.split_size_label.grid(row=3, column=0, padx=20, pady=10, sticky="w")

        self.split_size_entry = ttkb.Entry(self, font=("Helvetica", 12))
        self.split_size_entry.grid(row=3, column=1, padx=20, pady=10, sticky="ew")

        # Output Folder
        self.output_folder_label = ttkb.Label(self, text="Select output folder:", font=("Helvetica", 12))
        self.output_folder_label.grid(row=4, column=0, padx=20, pady=10, sticky="w")

        output_folder_frame = ttkb.Frame(self)
        output_folder_frame.grid(row=4, column=1, padx=20, pady=5, sticky="ew")
        output_folder_frame.grid_columnconfigure(0, weight=1)

        self.output_entry = ttkb.Entry(output_folder_frame, textvariable=self.output_folder_var, font=("Helvetica", 12), state="readonly")
        self.output_entry.grid(row=0, column=0, sticky="ew")

        self.output_browse = ttkb.Button(output_folder_frame, text="Browse", command=self.select_output_folder)
        self.output_browse.grid(row=0, column=1, padx=(5, 0))

        # Filename Format
        self.filename_format_label = ttkb.Label(self, text="Enter filename format:", font=("Helvetica", 12))
        self.filename_format_label.grid(row=5, column=0, padx=20, pady=10, sticky="w")

        self.filename_format_entry = ttkb.Entry(self, font=("Helvetica", 12))
        self.filename_format_entry.grid(row=5, column=1, padx=20, pady=5, sticky="ew")

        self.filename_format_instructions = ttkb.Label(
            self,
            text="Use column names from CSV in curly brackets.\n(e.g., Invitation to {province}, {LGU})",
            font=("Helvetica", 10),
            wraplength=450,
            anchor="w",
            justify="left"
        )
        self.filename_format_instructions.grid(row=6, column=1, padx=20, pady=(0, 10), sticky="w")

        # Style for Checkbutton
        style = ttkb.Style()
        style.configure("Custom.TCheckbutton", font=("Helvetica", 11))

        # Numbering Checkbox & Instructions
        self.numbering_var = tk.BooleanVar()
        self.numbering_checkbox = ttkb.Checkbutton(
            self,
            text="Append numbering to filenames",
            variable=self.numbering_var,
            style="Custom.TCheckbutton"
        )
        self.numbering_checkbox.grid(row=7, column=1, padx=20, pady=(0, 5), sticky="w")

        self.numbering_format_instructions = ttkb.Label(
            self,
            text="Check this box to add numbering to filenames.\nUse {num_count} in the format to pick its location.",
            font=("Helvetica", 10),
            wraplength=450,
            anchor="w",
            justify="left"
        )
        self.numbering_format_instructions.grid(row=8, column=1, padx=20, pady=(0, 10), sticky="w")

        # Buttons
        self.clear_button = ttkb.Button(self, text="Clear", command=self.clear_fields, bootstyle="danger")
        self.clear_button.grid(row=9, column=0, pady=(20, 10), padx=20, sticky="ew")

        self.run_button = ttkb.Button(self, text="Start", command=self.run_split, bootstyle="success")
        self.run_button.grid(row=9, column=1, pady=(20, 10), padx=20, sticky="ew")

        # Success Message
        self.success_message_label = ttkb.Label(self, text="", font=("Helvetica", 12), foreground="#4CAF50")
        self.success_message_label.grid(row=10, column=0, columnspan=2, pady=5)

    def clear_fields(self):
        self.pdf_path = None
        self.csv_path = None
        self.output_folder = None
        self.filename_format = None

        self.pdf_drop.config(text="No file selected")
        self.csv_drop.config(text="No file selected")
        self.output_selected_label.config(text="No folder selected")
        self.filename_format_entry.delete(0, tk.END)
        self.split_size_entry.delete(0, tk.END)
        self.success_message_label.config(text="")
        self.numbering_var.set(False)

    def on_pdf_drop(self, event):
        self.pdf_path = event.data.strip("{}")
        
        # Check if the file exists
        if not os.path.isfile(self.pdf_path):
            messagebox.showerror("Error", "Please upload a valid PDF file.")
            return
        
        # Check if the file is a PDF by checking its extension
        file_extension = os.path.splitext(self.pdf_path)[1].lower()
        if file_extension != '.pdf':
            messagebox.showerror("Incorrect File Format", "The uploaded file is not in PDF format. Please upload a valid PDF file.")
            return
        
        self.pdf_drop.config(text=f"Selected PDF: {os.path.basename(self.pdf_path)}")

    def on_csv_drop(self, event):
        self.csv_path = event.data.strip("{}")
        
        # Check if the file exists and has a .csv extension
        if not os.path.isfile(self.csv_path):
            messagebox.showerror("Error", "Please upload a valid CSV file.")
            return
        
        file_extension = os.path.splitext(self.csv_path)[1].lower()
        if file_extension != '.csv':
            messagebox.showerror("Incorrect File Format", "The uploaded file is not in CSV format. Please upload a valid CSV file.")
            return
        
        self.csv_drop.config(text=f"Selected CSV: {os.path.basename(self.csv_path)}")

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder = folder
            self.output_folder_var.set(folder)

    def run_split(self):
        if not self.pdf_path or not self.csv_path or not self.output_folder:
            messagebox.showerror("Missing Output Folder Path", "Please select an output folder.")
            return

        # Check if the filename format is empty
        self.filename_format = self.filename_format_entry.get().strip()
        if not self.filename_format:
            messagebox.showerror("Missing Filename Format", "Please enter a valid filename format.")
            return

        try:
            split_size = int(self.split_size_entry.get())
            if split_size <= 0:
                messagebox.showerror("Error", "Please enter a valid number greater than zero for the number of pages per split.")
                return
        except ValueError:
            messagebox.showerror("Invalid Page Amount", "Please enter a valid number greater than zero for the number of pages per split.")
            return

        try:
            df = pd.read_csv(self.csv_path)
            column_names = list(df.columns)
            placeholder_columns = self.find_columns_in_filename(self.filename_format)

            missing_columns = [col for col in placeholder_columns if col not in column_names and col != "num_count"]
            if missing_columns:
                messagebox.showerror("Missing CSV Column", f"You're missing these columns in the CSV file: {', '.join(missing_columns)}")
                return
        except Exception as e:
            messagebox.showerror("Can't Read CSV File", f"There was a problem reading the CSV file: {e}")
            return

        try:
            count = self.split_pdf(self.pdf_path, split_size, df, placeholder_columns, self.output_folder)
            self.success_message_label.config(text=f"PDF split and files renamed successfully! {count} files created.")
        except Exception as e:
            messagebox.showerror("Splitting Error", f"An error occurred while attempting to split the PDF: {str(e)}")

    def find_columns_in_filename(self, filename_format):
        return re.findall(r"\{(.*?)\}", filename_format)

    def split_pdf(self, pdf_path, split_size, df, placeholder_columns, output_folder):
        pdf_document = fitz.open(pdf_path)
        total_pages = pdf_document.page_count
        count = 0
        total_outputs = (total_pages + split_size - 1) // split_size
        num_digits = len(str(total_outputs))
        use_numbering = self.numbering_var.get()

        for i in range(0, total_pages, split_size):
            output_pdf = fitz.open()
            end_page = min(i + split_size, total_pages)
            for page_num in range(i, end_page):
                output_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

            row_index = min(i, len(df) - 1)
            number_str = str(count + 1).zfill(num_digits) if total_outputs > 1 else ""

            filename = self.generate_filename(df.iloc[row_index], placeholder_columns, number_str, use_numbering, total_outputs)
            output_path = os.path.join(output_folder, filename)
            output_pdf.save(output_path)
            count += 1

        return count

    def generate_filename(self, row, placeholder_columns, number_str, use_numbering, total_outputs):
        filename = self.filename_format
        for col in placeholder_columns:
            if col != "num_count":
                filename = filename.replace(f"{{{col}}}", str(row[col]))

        if total_outputs > 1:
            if use_numbering:
                if "{num_count}" in filename:
                    filename = filename.replace("{num_count}", number_str)
                else:
                    filename += f"_{number_str}"
            else:
                # Numbering is OFF, but still multiple outputs → append anyway
                filename += f"_{number_str}"

        return filename + ".pdf"

if __name__ == "__main__":
    app = SplitPDFApp()
    app.mainloop()
