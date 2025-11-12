# main.py
from pathlib import Path
import sys
import time

from utils import (
    load_excel_mapping,
    scan_icbc_pdfs,
    copy_pdfs,
    move_pdfs,
    auto_archive,
    reincrement_pdfs,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

DEFAULTS = {
    "copy_with_no_producer_two": True,
    "min_age_to_archive": 2,
}


def bulk_copy_icbc_tool():
    print("📄 Bulk Copy ICBC Tool\n")

    mapping_path = Path.cwd() / "config.xlsx"

    # --- Check if config file exists ---
    if not mapping_path.exists():
        print(f"⚠️ Missing configuration file: '{mapping_path}'")
        print("Please create or place 'config.xlsx' in the current working directory.")
        print("Exiting in ", end="", flush=True)
        for i in range(7, 0, -1):
            print(f"{i} ", end="", flush=True)
            time.sleep(1)
        print("\n👋 Done.")
        sys.exit(1)

    mapping_path = Path.cwd() / "config.xlsx"
    mapping_data = load_excel_mapping(mapping_path, sheet_index=1, start_row=4)

    input_folder = mapping_data.get("b1")
    output_folder = mapping_data.get("b2")
    producer_mapping = mapping_data.get("producer_mapping", {})

    folders_missing = False
    if input_folder.exists():
        print(f"✅ Input folder path: {input_folder}")
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
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
    )

    # ------------------- ICBC File Mover -------------------
    moved_files = []
    moved_files = move_pdfs(
        files=copied_files,
        copy_with_no_producer_two=DEFAULTS["copy_with_no_producer_two"],
        root_folder=output_folder,
    )

    # ------------------- Auto Archive -------------------
    archived_files = auto_archive(
        root_path=output_folder, min_age_to_archive=DEFAULTS["min_age_to_archive"]
    )
    if archived_files:
        reincrement_pdfs(root_dir=output_folder)

    # ------------------- Consolidated Log -------------------
    log_path = Path.cwd() / "log.txt"
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("=== PDF Copy Summary ===\n")
        log.write(f"Total PDFs scanned:  {len(scanned_data)}\n")
        log.write(f"Total PDFs copied:   {len(copied_files)}\n")
        log.write(f"Total PDFs moved:    {len(moved_files)}\n")

        log.write("=== PDFs not copied to output folder ===\n")
        for file_path in non_icbc_files:
            log.write(str(file_path) + "\n")
        log.write("\n")

        log.write(
            "=== PDFs with no producer two code moved to a 'root/subfolder' ===\n"
        )
        for file_path in moved_files:
            log.write(str(file_path) + "\n")

    print(f"\n📝 Log saved to: {log_path}")

    print("\nExiting in ", end="", flush=True)
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


# --- Main execution ---
if __name__ == "__main__":
    bulk_copy_icbc_tool()
