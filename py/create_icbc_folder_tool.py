from pathlib import Path
import sys
import time

from utils import (
    load_excel_mapping,
    scan_icbc_pdfs,
    copy_pdfs,
    match_pdfs,
    auto_archive,
    reincrement_pdfs,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

DEFAULTS = {
    "copy_with_no_producer_two": True,
    "min_age_to_archive": 1,
    "ignore_archive": False,
}


def _countdown(seconds: int) -> None:
    print("Exiting in ", end="", flush=True)
    for i in range(seconds, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print()


def create_icbc_folder_tool():
    print("Create ICBC Folder Tool\n")

    mapping_path = Path.cwd() / "config.xlsx"

    if not mapping_path.exists():
        print(f"Missing configuration file: '{mapping_path}'")
        print("Please create or place 'config.xlsx' in the current working directory.")
        _countdown(7)
        print("Done.")
        sys.exit(1)

    # ---- Load config ------------------------------------------------
    mapping = load_excel_mapping(
        sheet_name="Create ICBC Folder Tool",
        input_folder_row=1,
        output_folder_row=2,
        agency_number_row=None,
    )

    input_folder = mapping.input_folder
    output_folder = mapping.output_folder
    producer_mapping = mapping.producer_mapping

    # ---- Validate folders ------------------------------------------
    folders_missing = False

    if input_folder and input_folder.exists():
        print(f"Input folder path: {input_folder}")
    else:
        print(f"Input folder '{input_folder}' does not exist.")
        folders_missing = True

    if not output_folder:
        print("Output folder path not set in config.")
        folders_missing = True
    elif not output_folder.parent.exists():
        print(
            f"Parent path '{output_folder.parent}' does not exist. Cannot create output folder."
        )
        folders_missing = True
    else:
        output_folder.mkdir(exist_ok=True)
        print(f"Output folder path: {output_folder}")

    if folders_missing:
        print("Please correct the folder paths in 'config.xlsx'.")
        _countdown(7)
        print("Done.")
        sys.exit(1)

    # ---- Scan -------------------------------------------------------
    scan = scan_icbc_pdfs(
        input_folder,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=None,
        copy_mode=True,
    )

    # ---- Copy -------------------------------------------------------
    copied_files = copy_pdfs(
        documents=scan.documents,
        output_root_dir=output_folder,
        producer_mapping=producer_mapping,
        ignore_archive=DEFAULTS["ignore_archive"],
    )

    # ---- Match to producer subfolders -------------------------------
    files_without_producer = [f for f in copied_files if f.parent == output_folder]
    matched_files = match_pdfs(
        files=files_without_producer,
        copy_with_no_producer_two=DEFAULTS["copy_with_no_producer_two"],
        root_folder=output_folder,
    )

    # ---- Archive ----------------------------------------------------
    archived_files = auto_archive(
        root_path=output_folder,
        min_age_years=DEFAULTS["min_age_to_archive"],
    )
    if archived_files:
        reincrement_pdfs(root_dir=output_folder)

    # ---- Log --------------------------------------------------------
    log_path = Path.cwd() / "log.txt"
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("=== Create ICBC Folder Tool Summary ===\n")

        if scan.non_icbc:
            log.write("=== Non ICBC PDFs found ===\n")
            log.writelines(f"{p}\n" for p in scan.non_icbc)
            log.write("\n")

        if scan.payment_plans:
            log.write("=== Payment Plan Agreements and Receipts ===\n")
            log.writelines(f"{p}\n" for p in scan.payment_plans)
            log.write("\n")

        if scan.unreadable:
            log.write("=== PDFs that could NOT be opened ===\n")
            log.writelines(f"{p}\n" for p in scan.unreadable)
            log.write("\n")

        if matched_files:
            log.write("=== ICBC PDFs matched to a producer subfolder ===\n")
            log.writelines(f"{p}\n" for p in matched_files)

    print(f"\n Log saved to: {log_path}")
    _countdown(3)


if __name__ == "__main__":
    create_icbc_folder_tool()
