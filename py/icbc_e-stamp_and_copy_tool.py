import fitz
import timeit
import time
from pathlib import Path
from datetime import datetime

from utils import (
    _file_key,
    _extract_filename_timestamp,
    progressbar,
    scan_icbc_pdfs,
    load_excel_mapping,
    copy_pdfs,
    match_pdfs,
    auto_archive,
    reincrement_pdfs,
    PFX_STAMPING,
    validation_stamp,
    stamp_time_of_validation,
    save_batch_copy,
    save_customer_copy,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

# ────────────── Constants ────────────── #
DEFAULTS = {
    "number_of_pdfs": 10,
    "copy_with_no_producer_two": True,
    "min_age_to_archive": 1,
    "ignore_archive": False,
}


# ────────────── Main Tool ────────────── #
def icbc_e_stamp_tool() -> None:
    print("ICBC E-Stamp & Copy Tool\n")
    start_total = timeit.default_timer()

    # ── Define Desktop stamping folder
    desktop_path = Path.home() / "Desktop"
    if not desktop_path.exists():
        desktop_path = Path.cwd()
    STAMP_OUTPUT_FOLDER = desktop_path / "ICBC E-Stamp Copies"
    STAMP_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # ── Load Excel mapping
    mapping = load_excel_mapping()
    input_folder: Path = mapping.input_folder or (Path.home() / "Downloads")
    COPY_OUTPUT_FOLDER: Path = mapping.output_folder
    producer_mapping = mapping.producer_mapping
    copy_mode = bool(COPY_OUTPUT_FOLDER and COPY_OUTPUT_FOLDER.exists())

    # ── Stage 1: Scan PDFs
    scan = scan_icbc_pdfs(
        input_dir=input_folder,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=DEFAULTS["number_of_pdfs"],
        stamping_mode=True,
        copy_mode=copy_mode,
        config_agency_number=mapping.agency_number,
    )
    total_scanned = len(scan.documents)

    # ── Stage 2: Stamping → Desktop folder

    existing_cache: dict[str, set[str]] = {}
    batch_dir = STAMP_OUTPUT_FOLDER / "ICBC Batch Copies"

    for search_root in (STAMP_OUTPUT_FOLDER, batch_dir):
        if not search_root.exists():
            continue
        for pdf in search_root.rglob("*.pdf"):
            key = _file_key(pdf.stem)
            ts = _extract_filename_timestamp(pdf)
            if ts:
                existing_cache.setdefault(key, set()).add(ts)

    stamped_counter = 0
    for path, document in progressbar(
        list(reversed(list(scan.documents.items()))), prefix=PFX_STAMPING, size=10
    ):
        if not document.transaction_timestamp or not document.validation_stamp_coords:
            continue

        stamp_key = _file_key(document.stamp_name())
        base_key = _file_key(document.base_name())
        existing = existing_cache.get(stamp_key, set()) | existing_cache.get(
            base_key, set()
        )

        if document.transaction_timestamp in existing:
            continue

        ts_dt = datetime.strptime(document.transaction_timestamp, "%Y%m%d%H%M%S")

        try:
            with fitz.open(path) as doc:
                doc = validation_stamp(doc, document, ts_dt)
                doc = stamp_time_of_validation(doc, document, ts_dt)
                save_batch_copy(doc, document, STAMP_OUTPUT_FOLDER)
                save_customer_copy(doc, document, STAMP_OUTPUT_FOLDER)

            existing_cache.setdefault(stamp_key, set()).add(
                document.transaction_timestamp
            )
            existing_cache.setdefault(base_key, set()).add(
                document.transaction_timestamp
            )
            stamped_counter += 1
        except Exception as e:
            print(f"Error processing {path}: {e}")

    if stamped_counter > 0:
        print(
            "\n\033[1m\033[4m Stamping complete! ICBC E-Stamp Copies folder is ready now!\033[0m\n"
        )

    # ── Stage 3: Copy → Excel folder
    copied_files = []
    if copy_mode:
        copied_files = copy_pdfs(
            documents=scan.documents,
            output_root_dir=COPY_OUTPUT_FOLDER,
            producer_mapping=producer_mapping,
            ignore_archive=DEFAULTS["ignore_archive"],
        )

        files_without_producer = [
            f for f in copied_files if f.parent == COPY_OUTPUT_FOLDER
        ]
        if files_without_producer:
            match_pdfs(
                files=files_without_producer,
                copy_with_no_producer_two=DEFAULTS["copy_with_no_producer_two"],
                root_folder=COPY_OUTPUT_FOLDER,
            )

        archived_files = auto_archive(
            root_path=COPY_OUTPUT_FOLDER,
            min_age_years=DEFAULTS["min_age_to_archive"],
        )
        if archived_files:
            reincrement_pdfs(root_dir=COPY_OUTPUT_FOLDER)
    else:
        print(
            f"ICBC Copies folder '{COPY_OUTPUT_FOLDER}' not found or invalid — skipping copy step."
        )

    # ── Summary
    elapsed = timeit.default_timer() - start_total
    print(f"\nTotal PDFs scanned: {total_scanned}")
    print(f"Total PDFs stamped: {stamped_counter}")
    print(f"Total PDFs copied:  {len(copied_files)}")
    print(f"Total execution time: {elapsed:.2f} seconds")

    # ── Exit countdown
    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


if __name__ == "__main__":
    icbc_e_stamp_tool()
