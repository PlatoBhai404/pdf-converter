# 📄 Personal Ultimate Document Converter

A fully functional, asynchronous local desktop application for converting documents between PDF and other popular formats. Built with Python and CustomTkinter.

---

## ✨ Features

- **PDF → Word** — Structurally converts PDF layout into an editable `.docx` file
- **PDF → Excel** — Extracts embedded tables from PDFs into an `.xlsx` workbook
- **PDF → HTML** — Converts PDF components into clean HTML web code
- **PDF → Images** — Renders every page of a PDF as individual `.png` files
- **PDF → Text** — Extracts text from PDFs while preserving visual layout
- **Images → PDF** — Compiles multiple `.png`/`.jpg` images into a single PDF
- **Word → PDF** — Converts `.docx` files to PDF via local Office automation
- **Text → PDF** — Renders plain `.txt` files into formatted PDFs
- **HTML → PDF** — Converts HTML pages to PDF via WeasyPrint
- **PowerPoint → PDF** — Converts `.pptx` slides to PDF via Windows COM automation

---

## 🗂 Project Structure

```
pdf-converter/
│
├── app.py                  # CustomTkinter GUI & async threading engine
├── requirements.txt        # Third-party dependency list
├── .gitignore              # Git ignore rules
│
└── converters/
    ├── __init__.py         # Package namespace (empty)
    ├── pdf_to_all.py       # Extraction engine (PDF → X)
    └── all_to_pdf.py       # Generation engine (X → PDF)
```

---

## 🖥 Requirements

- **Python** 3.9 or higher
- **Windows OS** (required for `.docx` and `.pptx` → PDF via COM automation)
- **Poppler** — Required by `pdf2image` for PDF-to-image conversion
  - Download: https://github.com/oschwartz10612/poppler-windows/releases
  - Add the `bin/` folder to your system `PATH`
- **Microsoft Office** — Required for `.docx` → PDF and `.pptx` → PDF conversion

---

## ⚙️ Installation

**1. Clone the repository**
```bash
git clone https://github.com/your-username/pdf-converter.git
cd pdf-converter
```

**2. Create and activate a virtual environment**
```bash
python -m venv venv
venv\Scripts\activate
```

**3. Install dependencies**
```bash
pip install -r requirements.txt
```

**4. Run the application**
```bash
python app.py
```

---

## 🚀 Usage

1. Launch the app by running `python app.py`
2. Click **Select File(s)** to choose your source document(s)
3. The format dropdown will automatically populate based on the input file type
4. Select your desired **target format** from the dropdown
5. Click **Convert Now** — conversion runs on a background thread so the UI stays responsive
6. On success, click **📂 Open Destination Location** to reveal the output file

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `customtkinter` | Modern GUI framework |
| `pypdf` | Core PDF reading |
| `pdfplumber` | Layout-aware text & table extraction |
| `pdf2image` | PDF page → image rendering |
| `pdf2docx` | PDF → Word/HTML structural conversion |
| `openpyxl` | Excel workbook generation |
| `camelot-py` | Advanced PDF table extraction |
| `weasyprint` | HTML → PDF rendering |
| `docx2pdf` | Word → PDF via Office automation |
| `reportlab` | Text → PDF generation |
| `pywin32` | Windows COM automation (PPTX → PDF) |
| `Pillow` | Image processing for image → PDF |

---

## ⚠️ Known Limitations

- **Windows only** — `docx2pdf` and `pptx_to_pdf` rely on Windows COM (Microsoft Office must be installed)
- **Poppler required** — PDF-to-image conversion will fail without Poppler in your system `PATH`
- **Table extraction** — `pdf_to_excel` only succeeds if the PDF contains detectable grid tables; it returns `False` otherwise
- **Complex PDFs** — Heavily styled or scanned PDFs may produce imperfect layout conversions

---

## 📄 License

This project is released for personal and educational use. Feel free to fork and extend it.
