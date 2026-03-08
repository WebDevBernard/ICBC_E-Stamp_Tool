import fitz
import timeit
import time
from pathlib import Path
from datetime import datetime

from utils import (
    ICBCDocument,
    _file_key,
    _extract_filename_timestamp,
    unique_file_path,
    progressbar,
    scan_icbc_pdfs,
    load_excel_mapping,
    copy_pdfs,
    match_pdfs,
    auto_archive,
    reincrement_pdfs,
    PFX_STAMPING,
)
from constants import ICBC_PATTERNS, PAGE_RECTS

# ────────────── Constants ────────────── #
VALIDATION_STAMP_OFFSET = (-4.25, 23.77, 1.58, 58.95)
TIME_OF_VALIDATION_OFFSET = (0.0, 10.35, 0.0, 40.0)
TIME_STAMP_OFFSET = (0.0, 13.0, 0.0, 0.0)
TIME_OF_VALIDATION_AM_OFFSET = (0.0, -0.6, 0.0, 0.0)
TIME_OF_VALIDATION_PM_OFFSET = (0.0, 21.2, 0.0, 0.0)

DEFAULTS = {
    "number_of_pdfs": 10,
    "copy_with_no_producer_two": True,
    "min_age_to_archive": 1,
    "ignore_archive": False,
}


# ────────────── PDF Stamping Functions ────────────── #
def validation_stamp(
    doc: fitz.Document, document: ICBCDocument, ts_dt: datetime
) -> fitz.Document:
    for page_num, (x0, y0, x1, y1) in document.validation_stamp_coords:
        dx0, dy0, dx1, dy1 = VALIDATION_STAMP_OFFSET
        agency_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        date_rect = fitz.Rect(
            agency_rect.x0 + TIME_STAMP_OFFSET[0],
            agency_rect.y0 + TIME_STAMP_OFFSET[1],
            agency_rect.x1 + TIME_STAMP_OFFSET[2],
            agency_rect.y1 + TIME_STAMP_OFFSET[3],
        )
        page = doc[page_num]
        page.insert_textbox(
            agency_rect,
            document.agency_number,
            fontname="spacembo",
            fontsize=9,
            align=1,
        )
        page.insert_textbox(
            date_rect,
            ts_dt.strftime("%b %d, %Y"),
            fontname="spacemo",
            fontsize=9,
            align=1,
        )
    return doc


def stamp_time_of_validation(
    doc: fitz.Document, document: ICBCDocument, ts_dt: datetime
) -> fitz.Document:
    am_pm_offset = (
        TIME_OF_VALIDATION_AM_OFFSET
        if ts_dt.hour < 12
        else TIME_OF_VALIDATION_PM_OFFSET
    )
    for page_num, (x0, y0, x1, y1) in document.time_of_validation_coords:
        dx0, dy0, dx1, dy1 = TIME_OF_VALIDATION_OFFSET
        dx0 += am_pm_offset[0]
        dy0 += am_pm_offset[1]
        time_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        doc[page_num].insert_textbox(
            time_rect, ts_dt.strftime("%I:%M"), fontname="helv", fontsize=6, align=2
        )
    return doc


def save_batch_copy(
    doc: fitz.Document, document: ICBCDocument, output_folder: Path
) -> Path:
    batch_dir = output_folder / "ICBC Batch Copies"
    batch_dir.mkdir(parents=True, exist_ok=True)
    dest = unique_file_path(
        batch_dir / f"{document.base_name()} [{document.transaction_timestamp}].pdf"
    )
    doc.save(dest, garbage=4, deflate=True)
    return dest


def save_customer_copy(
    doc: fitz.Document, document: ICBCDocument, output_folder: Path
) -> Path:
    customer_pages = list(document.customer_copy_pages)
    if document.top and (doc.page_count - 1) not in customer_pages:
        customer_pages.append(doc.page_count - 1)
    pages_to_delete = [i for i in range(doc.page_count) if i not in customer_pages]
    for page_num in reversed(pages_to_delete):
        doc.delete_page(page_num)
    dest = unique_file_path(
        output_folder / f"{document.stamp_name()} (Customer Copy).pdf"
    )
    doc.save(dest, garbage=4, deflate=True)
    return dest


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
    mapping = load_excel_mapping(Path.cwd() / "config.xlsx")
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
    #
    # Pre-build a cache of (file_key → set[timestamp]) from all PDFs already
    # in the stamp output folder and its ICBC Batch Copies sub-dir.
    # Replaces per-document rglob calls inside the loop (was O(n²)).
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

        # O(1) cache lookup instead of two rglob calls per document
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

            # Keep cache consistent so later iterations in this run see it
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
