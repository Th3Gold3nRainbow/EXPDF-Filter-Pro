import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_PDF = os.path.join(BASE_DIR, "document.pdf")
EXCEL_FILE = os.path.join(BASE_DIR, "list.xlsx")
OUTPUT_PDF = os.path.join(BASE_DIR, "final_result.pdf")

def extract_matching_pdf_pages():
    print("--- STARTING EXTRACTION ---")
    
    # 1. Load names from Excel (Column 'Name')
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Look for the 'Name' column (case-insensitive and ignores spaces)
        col_name = next((c for c in df.columns if str(c).strip().lower() == 'name'), None)
        
        if col_name is None:
            print(f"Error: 'Name' column not found. Available columns: {df.columns.tolist()}")
            return

        # Clean the list: remove empty rows, convert to string, strip spaces, and get unique values
        name_list = df[col_name].dropna().astype(str).str.strip().unique().tolist()
        print(f"List loaded: {len(name_list)} names to search for.")
        
    except Exception as e:
        print(f"Error while reading Excel file: {e}")
        return

    # 2. Read source PDF and prepare the new PDF writer
    try:
        reader = PdfReader(SOURCE_PDF)
        writer = PdfWriter()
        total_source_pages = len(reader.pages)
        pages_found = 0

        print(f"Analyzing {total_source_pages} pages in the document...")

        for i in range(total_source_pages):
            page = reader.pages[i]
            page_text = page.extract_text()
            
            if page_text:
                uppercase_text = page_text.upper()
                
                # Check if any name from the list is present on this page
                for name in name_list:
                    if name.upper() in uppercase_text:
                        writer.add_page(page)
                        pages_found += 1
                        print(f" > Page {i+1} added (Match found: {name})")
                        break # Move to next page to prevent duplicate entries for the same page

        # 3. Save the final PDF file
        if pages_found > 0:
            with open(OUTPUT_PDF, "wb") as output_file:
                writer.write(output_file)
            print("-" * 40)
            print(f"SUCCESS: The file '{os.path.basename(OUTPUT_PDF)}' has been created.")
            print(f"Total pages extracted: {pages_found}")
        else:
            print("FAILED: No pages matched the names in the list.")

    except Exception as e:
        print(f"Error during PDF processing: {e}")

if __name__ == "__main__":
    extract_matching_pdf_pages()