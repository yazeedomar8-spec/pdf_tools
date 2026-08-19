"""
PDF Tool: Extract or delete page(s) from all PDF files in a folder
=====================================================================
Setup (one time only):
1. Open PowerShell and run:
   py -m pip install pypdf

2. Edit INPUT_FOLDER below

3. Run:
   py pdf_tool.py

Page selection format:
- Single page:         5
- Range:                3-6      (means pages 3 to 6)
- Multiple pages:       1,3,7    (means pages 1, 3, and 7)
- Mixed:                1,3-5,9
"""

import os
from pypdf import PdfReader, PdfWriter

# ============ EDIT THIS LINE ONLY ============
INPUT_FOLDER = r"D:\المخططات الجديدة\مخطط 716\طلبات الاستلام\السيول\CB\New folder\PDF_Output"
# ===============================================


def parse_page_selection(selection_text, total_pages):
    """Converts text like '1,3-5,9' into a list of page indices (0-based)"""
    page_indices = []
    parts = selection_text.replace(" ", "").split(",")

    for part in parts:
        if "-" in part:
            start, end = part.split("-")
            start, end = int(start), int(end)
            for p in range(start, end + 1):
                if 1 <= p <= total_pages:
                    page_indices.append(p - 1)
        else:
            p = int(part)
            if 1 <= p <= total_pages:
                page_indices.append(p - 1)

    return sorted(set(page_indices))


def process_files(action, selection_text):
    output_folder_name = "Extracted" if action == "1" else "After_Delete"
    output_folder = os.path.join(INPUT_FOLDER, output_folder_name)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    pdf_files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("No PDF files found in the specified folder.")
        return

    print(f"\nFound {len(pdf_files)} file(s). Starting processing...\n")

    success_count = 0
    failed_files = []

    for filename in pdf_files:
        input_path = os.path.join(INPUT_FOLDER, filename)
        suffix = "_extracted" if action == "1" else "_deleted"
        output_filename = os.path.splitext(filename)[0] + suffix + ".pdf"
        output_path = os.path.join(output_folder, output_filename)

        try:
            reader = PdfReader(input_path)

            if reader.is_encrypted:
                print(f"[SKIP] {filename} - file is password protected")
                failed_files.append(filename)
                continue

            total_pages = len(reader.pages)
            selected_indices = parse_page_selection(selection_text, total_pages)

            if not selected_indices:
                print(f"[FAIL] {filename} - the page number(s) entered don't exist "
                      f"in this file (it has {total_pages} pages)")
                failed_files.append(filename)
                continue

            writer = PdfWriter()

            if action == "1":
                # Extract: keep only the selected pages
                for idx in selected_indices:
                    writer.add_page(reader.pages[idx])
            else:
                # Delete: keep everything except the selected pages
                for idx in range(total_pages):
                    if idx not in selected_indices:
                        writer.add_page(reader.pages[idx])

            if len(writer.pages) == 0:
                print(f"[FAIL] {filename} - result has 0 pages (check your selection)")
                failed_files.append(filename)
                continue

            with open(output_path, "wb") as f:
                writer.write(f)

            print(f"[OK]   {filename} -> {output_filename} ({len(writer.pages)} page(s))")
            success_count += 1

        except Exception as e:
            print(f"[FAIL] {filename} - reason: {e}")
            failed_files.append(filename)

    print(f"\n----- Summary -----")
    print(f"Succeeded: {success_count} / {len(pdf_files)}")
    if failed_files:
        print(f"Failed/skipped files: {', '.join(failed_files)}")
    print(f"Output saved to: {output_folder}")


def main():
    print("=" * 50)
    print("PDF Tool - Extract or Delete Pages")
    print("=" * 50)
    print("\nWhat do you want to do?")
    print("1) Extract specific page(s)  (discard the rest)")
    print("2) Delete specific page(s)   (keep the rest)")

    action = input("\nEnter 1 or 2: ").strip()

    if action not in ("1", "2"):
        print("Invalid choice. Please run the script again and enter 1 or 2.")
        return

    print("\nPage number format:")
    print("  Single page:     5")
    print("  Range:           3-6")
    print("  Multiple pages:  1,3,7")
    print("  Mixed:           1,3-5,9")

    selection_text = input("\nEnter page number(s): ").strip()

    process_files(action, selection_text)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        input("\nDone. Press Enter to close this window...")
