# PDF Page Extractor by Name List

This Python utility automates the process of filtering and extracting specific pages from large PDF documents based on a list of names provided in an Excel file.

## 🚀 Features
* **Automated Filtering:** Quickly scans through hundreds of PDF pages.
* **Excel Integration:** Reads search terms directly from an `.xlsx` or `.xls` file.
* **Smart Detection:** Automatically identifies the "Name" column (case-insensitive).
* **Duplicate Prevention:** Ensures each matching page is only added once to the final result.
* **Graphical Interface (GUI):** Easy-to-use pop-up windows to select your files.

---

## 🛠️ Prerequisites

Before running the script, ensure you have the required libraries installed:

```bash
pip install pandas openpyxl PyPDF2
```
📂 Project Structure

To run the script successfully, organize your files as follows:
```bash
project-folder/
├── extract_pdf.py    # The Python script (GUI version)
├── document.pdf      # Your source PDF file
└── list.xlsx         # Excel file containing a 'Name' column
```

## 💻 How to Use
* Run the script: Execute python extract_pdf.py in your terminal or IDE.

* Select Source: A window will pop up asking you to select your Source PDF.

* Select List: Another window will ask for your Excel file containing the names.

* Save Result: Choose where you want to save the generated PDF and give it a name.

* Confirmation: A success message will appear showing how many pages were matched and extracted.

---

##  ⚠️ Important Notes
* Text Recognition: The PDF must contain selectable text. Scanned images without OCR (Optical Character Recognition) cannot be read by this script.

* Column Headers: Ensure your Excel sheet has a header named "Name" or "name".

  ---

##  📜 License
* This project is open-source and free to use.
