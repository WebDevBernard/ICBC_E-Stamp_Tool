import fitz
from pathlib import Path
from datetime import datetime
from utils import (
    get_stamp_name,
    unique_file_name,
    progressbar,
    find_existing_timestamps,
    scan_icbc_pdfs,
    load_excel_mapping,
    copy_pdfs,
    move_pdfs,
    auto_archive,
    reincrement_pdfs,
)
from constants import ICBC_PATTERNS, PAGE_RECTS
import timeit
import time

# -------------------- Local Constants -------------------- #

validation_stamp_coords_box_offset = (-4.25, 23.77, 1.58, 58.95)
time_of_validation_offset = (0.0, 10.35, 0.0, 40)
time_stamp_offset = (0, 13, 0, 0)
time_of_validation_am_offset = (0, -0.6, 0, 0)
time_of_validation_pm_offset = (0, 21.2, 0, 0)

# -------------------- Defaults -------------------- #

desktop_or_root = Path.home() / "Desktop"
if not desktop_or_root.exists():
    print("⚠️ Desktop Directory not found, using root directory instead.")
    desktop_or_root = Path.cwd()

output_dir = desktop_or_root / "ICBC E-Stamp Copies"
output_dir.mkdir(parents=True, exist_ok=True)

DEFAULTS = {
    "number_of_pdfs": 10,
    "output_dir": str(output_dir),
    "input_dir": str(Path.home() / "Downloads"),
    "copy_with_no_producer_two": True,
    "min_age_to_archive": 2,
}

# -------------------- PDF Stamping Functions -------------------- #


def validation_stamp(doc, info, ts_dt):
    for page_num, coords in info.get("validation_stamp_coords", []):
        page = doc[page_num]
        x0, y0, x1, y1 = coords
        dx0, dy0, dx1, dy1 = validation_stamp_coords_box_offset
        agency_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        date_rect = fitz.Rect(
            agency_rect.x0 + time_stamp_offset[0],
            agency_rect.y0 + time_stamp_offset[1],
            agency_rect.x1 + time_stamp_offset[2],
            agency_rect.y1 + time_stamp_offset[3],
        )
        page.insert_textbox(
            agency_rect, info["agency_number"], fontname="spacembo", fontsize=9, align=1
        )
        page.insert_textbox(
            date_rect,
            ts_dt.strftime("%b %d, %Y"),
            fontname="spacemo",
            fontsize=9,
            align=1,
        )
    return doc


def stamp_time_of_validation(doc, info, ts_dt):
    for page_num, coords in info.get("time_of_validation_coords", []):
        page = doc[page_num]
        x0, y0, x1, y1 = coords
        dx0, dy0, dx1, dy1 = time_of_validation_offset
        if ts_dt.hour < 12:
            dx0 += time_of_validation_am_offset[0]
            dy0 += time_of_validation_am_offset[1]
        else:
            dx0 += time_of_validation_pm_offset[0]
            dy0 += time_of_validation_pm_offset[1]
        time_rect = fitz.Rect(x0 + dx0, y0 + dy0, x1 + dx1, y1 + dy1)
        page.insert_textbox(
            time_rect, ts_dt.strftime("%I:%M"), fontname="helv", fontsize=6, align=2
        )
    return doc


def save_batch_copy(doc, info, output_dir):
    batch_dir = Path(output_dir) / "ICBC Batch Copies"
    batch_dir.mkdir(parents=True, exist_ok=True)
    stamp_name = get_stamp_name(info)
    batch_copy_path = batch_dir / f"{stamp_name}.pdf"
    batch_copy_path = Path(unique_file_name(batch_copy_path))
    doc.save(batch_copy_path, garbage=4, deflate=True)
    return batch_copy_path


def save_customer_copy(doc, info, output_dir):
    total_pages = doc.page_count
    customer_pages = info.get("customer_copy_pages", [])
    if info.get("temporary_operation_permit") and total_pages - 1 not in customer_pages:
        customer_pages.append(total_pages - 1)
    pages_to_delete = [i for i in range(total_pages) if i not in customer_pages]
    for page_num in reversed(pages_to_delete):
        doc.delete_page(page_num)
    stamp_name = get_stamp_name(info)
    customer_copy_name = f"{stamp_name} (Customer Copy).pdf"
    customer_copy_path = Path(output_dir) / customer_copy_name
    customer_copy_path = Path(unique_file_name(customer_copy_path))
    doc.save(customer_copy_path, garbage=4, deflate=True)
    return customer_copy_path


# -------------------- Main Function -------------------- #
def icbc_e_stamp_tool():
    print("📄 ICBC E-Stamp & Copy Tool\n")
    start_total = timeit.default_timer()

    input_dir = DEFAULTS["input_dir"]
    output_dir = DEFAULTS["output_dir"]

    copied_files = []

    # -------------------- Stage 1: Scan PDFs -------------------- #
    icbc_data, _ = scan_icbc_pdfs(
        input_dir=input_dir,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=DEFAULTS["number_of_pdfs"],
        stamping_mode=True,
        copy_mode=False,
    )
    total_scanned = len(icbc_data)

    # -------------------- Stage 2: Process PDFs -------------------- #
    stamped_counter = 0
    for path, info in progressbar(
        list(reversed(list(icbc_data.items()))),
        prefix="🖋️ Stamping PDFs:  ",
        size=10,
    ):
        ts = info.get("transaction_timestamp")
        if not ts or not info.get("validation_stamp_coords"):
            continue

        stamp_name = get_stamp_name(info)
        existing_timestamps = find_existing_timestamps(
            stamp_name, ICBC_PATTERNS["timestamp"], PAGE_RECTS["timestamp"], output_dir
        )
        if ts in existing_timestamps:
            continue

        ts_dt = datetime.strptime(ts, "%Y%m%d%H%M%S")

        try:
            doc = fitz.open(path)

            doc = validation_stamp(doc, info, ts_dt)
            doc = stamp_time_of_validation(doc, info, ts_dt)
            save_batch_copy(doc, info, output_dir)
            save_customer_copy(doc, info, output_dir)

            stamped_counter += 1

        except Exception as e:
            print(f"❌ Error processing {path}: {e}")

    # -------------------- Stage 3: Copy PDFs -------------------- #
    mapping_path = Path.cwd() / "config.xlsx"
    mapping_data = load_excel_mapping(mapping_path, sheet_index=0, start_row=3)
    output_folder = mapping_data.get("b1")
    producer_mapping = mapping_data.get("producer_mapping", {})

    copy_data, _ = scan_icbc_pdfs(
        input_dir=input_dir,
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=DEFAULTS["number_of_pdfs"],
        copy_mode=True,
    )

    if output_folder and Path(output_folder).exists():
        copied_files = copy_pdfs(
            icbc_data=copy_data,
            output_root_dir=output_folder,
            producer_mapping=producer_mapping,
            regex_patterns=ICBC_PATTERNS,
            page_rects=PAGE_RECTS,
        )
        move_pdfs(
            files=copied_files,
            copy_with_no_producer_two=DEFAULTS["copy_with_no_producer_two"],
            root_folder=output_folder,
        )
        archived_files = auto_archive(
            root_path=output_folder, min_age_to_archive=DEFAULTS["min_age_to_archive"]
        )
        if archived_files:
            reincrement_pdfs(root_dir=output_folder)
    else:
        print(
            f" ⚠️ICBC Copies folder: '{output_folder}' not found or invalid — skipping copy step."
        )

    # -------------------- Summary -------------------- #
    end_total = timeit.default_timer()
    print(f"\nTotal PDFs scanned: {total_scanned}")
    print(f"Total PDFs stamped: {stamped_counter}")
    print(f"Total PDFs copied:  {len(copied_files) if copied_files else 0}")
    print(f"✅ Total script execution time: {end_total - start_total:.2f} seconds")
    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


if __name__ == "__main__":
    icbc_e_stamp_tool()
