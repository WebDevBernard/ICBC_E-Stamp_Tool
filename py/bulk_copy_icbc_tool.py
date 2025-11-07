# main.py
from pathlib import Path
import sys
import time

from utils import (
    load_excel_mapping,
    scan_icbc_pdfs,
    copy_pdfs,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

# --- Main execution ---
if __name__ == "__main__":
    print("📄 ICBC Copy Tool\n")

    # Load folders and producer mapping from Excel
    mapping_path = Path.cwd() / "config.xlsx"
    mapping_data = load_excel_mapping(mapping_path, ws_index=1)
    input_folder = mapping_data.get("input_folder")
    output_folder = mapping_data.get("output_folder")
    producer_mapping = mapping_data.get("producer_mapping", {})

    # --- Check input/output folders ---
    folders_missing = False

    # Check input folder
    if not input_folder or not input_folder.exists():
        print(f"⚠️ Input folder '{input_folder}' does not exist.")
        folders_missing = True

    # Check output folder
    if not output_folder:
        print(f"⚠️ Output folder path not set in config.")
        folders_missing = True
    else:
        if output_folder.parent.exists():
            output_folder.mkdir(exist_ok=True)
            print(f"✅ Output folder ready: {output_folder}")
        else:
            print(
                f"⚠️ Parent path '{output_folder.parent}' does not exist. Cannot create output folder."
            )
            folders_missing = True

    # Exit with 10-second countdown if any folder issue
    if folders_missing:
        print("Please correct the folder paths in 'config.xlsx'.")
        print("Exiting in ", end="", flush=True)
        for i in range(10, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print("\n👋 Done.")
        sys.exit(1)

    # --- Scan PDFs ---
    scanned_data, non_icbc_files = scan_icbc_pdfs(
        input_folder, regex_patterns=ICBC_PATTERNS, page_rects=PAGE_RECTS
    )

    # --- Report for non-ICBC documents ---
    report_path = Path.cwd() / "report.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("Documents Not Copied\n")
        f.write("=" * 50 + "\n\n")
        for file_path in non_icbc_files:
            f.write(file_path + "\n")

    # --- Copy PDFs ---
    copied_count = copy_pdfs(
        icbc_data=scanned_data,
        output_root_dir=output_folder,
        producer_mapping=producer_mapping,
        create_subfolders=True,
    )

    # --- Summary + Exit Countdown ---
    print("\n📊 Summary")
    print(f"Total PDFs scanned: {len(scanned_data)}")
    print(f"Total PDFs copied:  {copied_count}")
    print(f"\n📝 ICBC copy report written to: {report_path}")
    print("\nExiting in ", end="", flush=True)
    for i in range(10, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print("\n👋 Done.")
