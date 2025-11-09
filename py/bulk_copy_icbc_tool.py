# main.py
from pathlib import Path
import sys
import time

from utils import (
    load_excel_mapping,
    scan_icbc_pdfs,
    copy_pdfs,
    move_pdfs,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

DEFAULTS = {
    "create_subfolders": True,
    "use_alt_name": True,
    "copy_with_no_producer_two": True,
}


def bulk_copy_icbc_tool():
    print("📄 Bulk Copy ICBC Tool\n")

    mapping_path = Path.cwd() / "config.xlsx"
    mapping_data = load_excel_mapping(mapping_path, sheet_index=1, start_row=4)

    input_folder = mapping_data.get("b1")
    output_folder = mapping_data.get("b2")
    producer_mapping = mapping_data.get("producer_mapping", {})

    folders_missing = False

    if not input_folder or not input_folder.exists():
        print(f"⚠️ Input folder '{input_folder}' does not exist.")
        folders_missing = True

    if not output_folder:
        print(f"⚠️ Output folder path not set in config.")
        folders_missing = True
    else:
        if output_folder.parent.exists():
            output_folder.mkdir(exist_ok=True)
            print(f"✅ Output folder path: {output_folder}")
        else:
            print(
                f"⚠️ Parent path '{output_folder.parent}' does not exist. Cannot create output folder."
            )
            folders_missing = True

    if folders_missing:
        print("Please correct the folder paths in 'config.xlsx'.")
        print("Exiting in ", end="", flush=True)
        for i in range(7, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print("\n👋 Done.")
        sys.exit(1)

    # ------------------- PDF Scanning and Copy -------------------
    scanned_data, non_icbc_files = scan_icbc_pdfs(
        input_folder,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=None,
        copy_mode=True,
    )

    copied_files = copy_pdfs(
        icbc_data=scanned_data,
        output_root_dir=output_folder,
        producer_mapping=producer_mapping,
        create_subfolders=DEFAULTS["create_subfolders"],
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        use_alt_name=DEFAULTS["use_alt_name"],
    )

    copied_count = len(copied_files)

    # ------------------- ICBC File Mover -------------------
    moved_files = []
    moved_files = move_pdfs(
        copied_files, copy_with_no_producer_two=DEFAULTS["copy_with_no_producer_two"]
    )

    # ------------------- Consolidated Log -------------------
    log_path = Path.cwd() / "log.txt"
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("=== PDF Copy Summary ===\n")
        log.write(f"Total PDFs scanned: {len(scanned_data)}\n")
        log.write(f"Total PDFs copied:  {copied_count}\n\n")

        log.write("=== PDFs not copied to output folder ===\n")
        for file_path in non_icbc_files:
            log.write(str(file_path) + "\n")
        log.write("\n")

        log.write("=== PDFs moved to subfolder with no producer two ===\n")
        for file_path in moved_files:
            log.write(str(file_path) + "\n")

    print(f"\n📝 Log saved to: {log_path}")

    print("\nExiting in ", end="", flush=True)
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print("\n👋 Done.")


# --- Main execution ---
if __name__ == "__main__":
    bulk_copy_icbc_tool()
