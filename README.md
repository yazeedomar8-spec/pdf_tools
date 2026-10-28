# PDF Tools

Two small Python scripts for batch document work on Windows.

[قراءة الدليل بالعربي](README.ar.md)

| Script | What it does | Library |
| --- | --- | --- |
| `excel_to_pdf.py` | Converts every Excel workbook in a folder to PDF in one run, using Excel itself so formatting is preserved | `pywin32` |
| `pdf_tool.py` | Extracts or deletes selected pages from every PDF in a folder | `pypdf` |

Both scripts write results to a new folder and never modify the original files.

## Requirements

- Windows with Python 3
- Microsoft Excel installed (for `excel_to_pdf.py` only)

```
py -m pip install pywin32 pypdf
```

## Usage

### Excel to PDF

1. Set `INPUT_FOLDER` at the top of `excel_to_pdf.py` to your Excel folder.
2. Run:
   ```
   py excel_to_pdf.py
   ```
3. PDFs are saved in a `PDF_Output` folder inside the input folder.

### Extract or delete PDF pages

1. Set `INPUT_FOLDER` at the top of `pdf_tool.py` to your PDF folder.
2. Run:
   ```
   py pdf_tool.py
   ```
3. Choose `1` to extract or `2` to delete, then enter the pages:

   | Format | Example |
   | --- | --- |
   | Single page | `5` |
   | Range | `3-6` |
   | List | `1,3,7` |
   | Mixed | `1,3-5,9` |

4. Results are saved in `Extracted` or `After_Delete`.

Password-protected PDFs are skipped with a warning.
