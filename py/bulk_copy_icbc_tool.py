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

DEFAULTS = {
    "create_subfolders": True,
    "use_alt_name": True,
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
            print(f"✅ Output folder ready: {output_folder}")
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

    scanned_data, non_icbc_files = scan_icbc_pdfs(
        input_folder,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=None,
        copy_mode=True,
    )

    report_path = Path.cwd() / "log.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("PDFs Not Copied\n")
        f.write("=" * 50 + "\n\n")
        for file_path in non_icbc_files:
            f.write(file_path + "\n")

    copied_count = copy_pdfs(
        icbc_data=scanned_data,
        output_root_dir=output_folder,
        producer_mapping=producer_mapping,
        create_subfolders=DEFAULTS["create_subfolders"],
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        use_alt_name=DEFAULTS["use_alt_name"],
    )

    print("\n📊 Summary")
    print(f"Total PDFs scanned: {len(scanned_data)}")
    print(f"Total PDFs copied:  {copied_count}")
    print(f"\n📝 ICBC copy report written to: {report_path}")
    print("\nExiting in ", end="", flush=True)
    for i in range(7, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print("\n👋 Done.")


# --- Main execution ---
if __name__ == "__main__":
    bulk_copy_icbc_tool()
