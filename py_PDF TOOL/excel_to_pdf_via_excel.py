"""
Convert a folder of Excel files to PDF using Excel itself
=============================================================
- Windows only (requires Microsoft Excel installed)
- Run this from Python on Windows directly (not from WSL)

Setup (one time only):
1. Make sure Python is installed on Windows (not WSL) - check with: py --version
2. Install the required library:
   py -m pip install pywin32
3. Edit the INPUT_FOLDER line below with the path to your Excel files
4. Run: py excel_to_pdf_via_excel.py
"""

import os
import win32com.client as win32

# ============ EDIT THIS LINE ONLY ============
INPUT_FOLDER = r"D:\المخططات الجديدة\مخطط 716\طلبات الاستلام\الحديقه\req"
# ===============================================

OUTPUT_FOLDER = os.path.join(INPUT_FOLDER, "PDF_Output")


def convert_folder_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel_files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith((".xlsx", ".xls", ".xlsm"))
    ]

    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    print(f"Found {len(excel_files)} file(s). Starting conversion...\n")

    # Start Excel in the background (invisible)
    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    success_count = 0
    failed_files = []

    for filename in excel_files:
        input_path = os.path.join(input_folder, filename)
        output_filename = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(output_folder, output_filename)

        try:
            wb = excel_app.Workbooks.Open(
                input_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )

            if wb is None:
                print(f"[FAIL] {filename} - Excel refused to open the file "
                      f"(try opening it manually in Excel to check for a warning dialog or protection)")
                failed_files.append(filename)
                continue

            # 0 = xlTypePDF
            wb.ExportAsFixedFormat(0, output_path)
            wb.Close(SaveChanges=False)
            print(f"[OK]   {filename}")
            success_count += 1
        except Exception as e:
            print(f"[FAIL] {filename} - reason: {e}")
            failed_files.append(filename)

    excel_app.Quit()

    print(f"\n----- Summary -----")
    print(f"Succeeded: {success_count} / {len(excel_files)}")
    if failed_files:
        print(f"Failed files: {', '.join(failed_files)}")
    print(f"Files saved to: {output_folder}")


if __name__ == "__main__":
    try:
        convert_folder_to_pdf(INPUT_FOLDER, OUTPUT_FOLDER)
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        input("\nDone. Press Enter to close this window...")
