# 🚀 EXPDF-Filter Pro (PEF)

**EXPDF-Filter Pro** is a modern, high-performance desktop application designed to automate the extraction of specific pages from large PDF documents based on criteria defined in an Excel spreadsheet.

Whether you need to filter 700+ pages by Name, ID, or Serial Number, this tool handles the heavy lifting with a clean, professional user interface.

---

## ✨ Features
* **Modern UI:** Built with a sleek Dark Mode interface using `CustomTkinter`.
* **Smart Column Detection:** Automatically reads your Excel file and lets you choose which column to use for filtering.
* **Multi-Threaded Processing:** Handles large PDFs (hundreds of pages) without freezing or crashing.
* **Real-time Progress:** Visual progress bar and page counter to track the extraction.
* **Zero Installation for Users:** Compiled into a standalone `.exe` (no Python required for the end-user).

---

## 🛠️ Built With
This software was developed using the following technologies:
* **Language:** [Python 3.x](https://www.python.org/)
* **UI Framework:** [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) (Modernized Tkinter)
* **PDF Logic:** [PyPDF2](https://pypi.org/project/PyPDF2/) (Reading and Writing PDF streams)
* **Data Handling:** [Pandas](https://pandas.pydata.org/) & [Openpyxl](https://openpyxl.readthedocs.io/) (Excel processing)
* **Concurrency:** `Threading` library for background task management.
* **Compilation:** [PyInstaller](https://pyinstaller.org/) (Executable bundling).

---

## 🚀 How It Works



1.  **Input:** You provide a source PDF and an Excel file containing your list (e.g., a list of employees or invoice numbers).
2.  **Selection:** You select the specific column in the Excel file that matches the text present in the PDF pages.
3.  **Process:** The app scans every page of the PDF. If a value from your Excel list is found in the text of a page, that page is captured.
4.  **Output:** A new PDF is generated in the same folder, containing only the relevant pages.

---

## 💻 Installation (For Developers)
If you want to run the source code:

1. Clone the repo:
   ```bash
   git clone [https://github.com/Th3Gold3nRainbow/EXPDF-Filter-Pro.git](https://github.com/Th3Gold3nRainbow/EXPDF-Filter-Pro.git)
