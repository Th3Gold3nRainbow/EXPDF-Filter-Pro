import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file(title, file_types):
    """Opens a file dialog to select a file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path

def select_save_path(title, default_name):
    """Opens a dialog to choose where to save the result."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title=title, defaultextension=".pdf", initialfile=default_name)
    root.destroy()
    return file_path

def run_extraction():
    # 1. Select the Source PDF
    source_pdf = select_file("Select the Source PDF", [("PDF files", "*.pdf")])
    if not source_pdf: return

    # 2. Select the Excel List
    excel_file = select_file("Select the Excel List", [("Excel files", "*.xlsx *.xls")])
    if not excel_file: return

    # 3. Select Save Location
    output_pdf = select_save_path("Save the Result As", "filtered_document.pdf")
    if not output_pdf: return

    print(f"--- STARTING EXTRACTION ---")
    
    # --- STEP 1: LOAD EXCEL ---
    try:
        df = pd.read_excel(excel_file)
        col_name = next((c for c in df.columns if str(c).strip().lower() == 'name'), None)
        
        if col_name is None:
            messagebox.showerror("Error", f"Column 'Name' not found.\nColumns found: {df.columns.tolist()}")
            return

        name_list = df[col_name].dropna().astype(str).str.strip().unique().tolist()
        print(f"List loaded: {len(name_list)} names found.")
    except Exception as e:
        messagebox.showerror("Excel Error", f"Could not read Excel: {e}")
        return

    # --- STEP 2: PROCESS PDF ---
    try:
        reader = PdfReader(source_pdf)
        writer = PdfWriter()
        pages_found = 0

        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                text_upper = text.upper()
                for name in name_list:
                    if name.upper() in text_upper:
                        writer.add_page(page)
                        pages_found += 1
                        print(f" > Match found on Page {i+1}: {name}")
                        break 

        # --- STEP 3: SAVE ---
        if pages_found > 0:
            with open(output_pdf, "wb") as out_f:
                writer.write(out_f)
            messagebox.showinfo("Success", f"Done! {pages_found} pages extracted to:\n{output_pdf}")
        else:
            messagebox.showwarning("No Matches", "No names from the list were found in the PDF.")

    except Exception as e:
        messagebox.showerror("PDF Error", f"An error occurred: {e}")

if __name__ == "__main__":
    run_extraction()
