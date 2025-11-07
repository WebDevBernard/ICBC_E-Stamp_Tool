import fitz
from pathlib import Path
from datetime import datetime
from utils import (
    get_base_name,
    unique_file_name,
    progressbar,
    search_insured_name,
    reverse_insured_name,
    find_existing_timestamps,
)
from constants import ICBC_PATTERNS, PAGE_RECTS
from utils import (
    scan_icbc_pdfs,
    load_excel_mapping,
    copy_pdfs,
)
import timeit
import time

# -------------------- Local Constants -------------------- #

timestamp_pattern = ICBC_PATTERNS["timestamp"]
license_plate_pattern = ICBC_PATTERNS["license_plate"]
temporary_permit_pattern = ICBC_PATTERNS["temporary_permit"]
agency_number_pattern = ICBC_PATTERNS["agency_number"]
customer_copy_pattern = ICBC_PATTERNS["customer_copy"]
validation_stamp_pattern = ICBC_PATTERNS["validation_stamp"]
time_of_validation_pattern = ICBC_PATTERNS["time_of_validation"]

timestamp_rect = PAGE_RECTS["timestamp"]
customer_copy_rect = PAGE_RECTS["customer_copy"]

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
    base_name = get_base_name(info)
    batch_copy_path = batch_dir / f"{base_name}.pdf"
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
    base_name = get_base_name(info)
    customer_copy_name = f"{base_name} (Customer Copy).pdf"
    customer_copy_path = Path(output_dir) / customer_copy_name
    customer_copy_path = Path(unique_file_name(customer_copy_path))
    doc.save(customer_copy_path, garbage=4, deflate=True)
    return customer_copy_path


# -------------------- Main Function -------------------- #
def icbc_e_stamp_tool():
    print("📄 ICBC E-Stamp Tool")
    start_total = timeit.default_timer()

    input_dir = DEFAULTS["input_dir"]
    output_dir = DEFAULTS["output_dir"]
    max_docs = DEFAULTS["number_of_pdfs"]

    # Get PDFs to process
    pdf_files = sorted(
        Path(input_dir).glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True
    )[:max_docs]

    icbc_data = {}

    # -------------------- Stage 1: Scan PDFs -------------------- #
    print("🔍 Scanning PDFs...")
    for pdf_path in progressbar(pdf_files, prefix="Scanning PDFs: ", size=10):
        try:
            with fitz.open(pdf_path) as doc:
                if doc.page_count == 0:
                    continue

                first_page = doc[0]
                full_text_first_page = first_page.get_text("text")
                ts_text = first_page.get_text(clip=timestamp_rect)

                ts_match = timestamp_pattern.search(ts_text)
                license_plate_match = license_plate_pattern.search(full_text_first_page)
                license_plate = (
                    license_plate_match.group(1).strip().upper()
                    if license_plate_match
                    else None
                )
                insured_name = reverse_insured_name(
                    search_insured_name(full_text_first_page)
                )
                temp_permit_found = bool(
                    temporary_permit_pattern.search(full_text_first_page)
                )
                agency_match = agency_number_pattern.search(full_text_first_page)
                agency_number = (
                    agency_match.group(1).strip() if agency_match else "UNKNOWN"
                )

                info_preview = {
                    "transaction_timestamp": ts_match.group(1) if ts_match else "",
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                }
                base_name = get_base_name(info_preview)
                existing_timestamps = find_existing_timestamps(
                    base_name, timestamp_pattern, timestamp_rect, output_dir
                )
                timestamp = (
                    ts_match.group(1)
                    if ts_match and ts_match.group(1) not in existing_timestamps
                    else None
                )

                customer_copy_pages = []
                validation_stamp_coords = []
                time_of_validation_coords = []

                for page_num, page in enumerate(doc):
                    clipped_customer_copy = page.get_text(clip=customer_copy_rect)
                    if customer_copy_pattern.search(clipped_customer_copy):
                        customer_copy_pages.append(page_num)

                    for block in page.get_text("blocks"):
                        word_text, coords = block[4], block[:4]
                        if validation_stamp_pattern.search(word_text):
                            validation_stamp_coords.append((page_num, coords))
                        if time_of_validation_pattern.search(word_text):
                            time_of_validation_coords.append((page_num, coords))

                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "temporary_operation_permit": temp_permit_found,
                    "agency_number": agency_number,
                    "customer_copy_pages": customer_copy_pages,
                    "validation_stamp_coords": validation_stamp_coords,
                    "time_of_validation_coords": time_of_validation_coords,
                }

        except Exception as e:
            print(f"❌ Error scanning {pdf_path}: {e}")

    total_scanned = len(pdf_files)

    # -------------------- Stage 2: Process PDFs -------------------- #
    stamped_counter = 0

    for path, info in progressbar(
        list(reversed(list(icbc_data.items()))),
        prefix="Stamping PDFs: ",
        size=10,
    ):
        ts = info.get("transaction_timestamp")
        if not ts or not info.get("validation_stamp_coords"):
            continue

        base_name = get_base_name(info)
        existing_timestamps = find_existing_timestamps(
            base_name, timestamp_pattern, timestamp_rect, output_dir
        )
        if ts in existing_timestamps:
            continue

        ts_dt = datetime.strptime(ts, "%Y%m%d%H%M%S")

        try:
            doc_batch = fitz.open(path)
            doc_customer = fitz.open(path)

            doc_batch = validation_stamp(doc_batch, info, ts_dt)
            doc_batch = stamp_time_of_validation(doc_batch, info, ts_dt)
            doc_customer = validation_stamp(doc_customer, info, ts_dt)
            doc_customer = stamp_time_of_validation(doc_customer, info, ts_dt)

            save_batch_copy(doc_batch, info, output_dir)
            save_customer_copy(doc_customer, info, output_dir)

            stamped_counter += 1

        except Exception as e:
            print(f"❌ Error processing {path}: {e}")

    # -------------------- Stage 3: Copy PDFs -------------------- #
    copied_count = None

    mapping_path = Path.cwd() / "config.xlsx"
    mapping_data = load_excel_mapping(mapping_path, ws_index="active")
    input_folder = DEFAULTS["input_dir"]
    output_folder = mapping_data.get("output_folder")
    producer_mapping = mapping_data.get("producer_mapping", {})

    copy_data = scan_icbc_pdfs(
        input_dir=DEFAULTS["input_dir"],
        regex_patterns=ICBC_PATTERNS,
        page_rects=PAGE_RECTS,
        max_docs=DEFAULTS["number_of_pdfs"],
        suffix_mode=True,
    )

    if output_folder and producer_mapping:
        if not Path(output_folder).exists():
            print(f"⚠️ Path '{output_folder}' does not exist. Skipping copy operation.")
        else:
            missing_subfolders = [
                folder
                for folder in producer_mapping.values()
                if not Path(output_folder, folder).exists()
            ]
            if missing_subfolders:
                print(
                    f"⚠️ The following subfolders do not exist under '{output_folder}':"
                )
                for folder in missing_subfolders:
                    print(f"   - {folder}")
            copied_count = copy_pdfs(
                icbc_data=copy_data,
                output_root_dir=output_folder,
                producer_mapping=producer_mapping,
                create_subfolders=False,
            )

    else:
        print("ℹ️ config.xlsx file not found — skipping copy step.")

    # -------------------- Summary -------------------- #
    end_total = timeit.default_timer()
    print(f"\nTotal PDFs scanned: {total_scanned}")
    print(f"Total PDFs stamped: {stamped_counter}")
    if copied_count is not None:
        print(f"Total PDFs copied: {copied_count}")
    print(f"✅ Total script execution time: {end_total - start_total:.2f} seconds")
    print("\nExiting in ", end="")
    for i in range(3, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)


if __name__ == "__main__":
    icbc_e_stamp_tool()
