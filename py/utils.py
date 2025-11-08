import re
import os
import shutil
import time
import sys
from pathlib import Path
import fitz
import openpyxl


# -------------------- Progress Bar -------------------- #
def progressbar(it, prefix="", size=60, out=sys.stdout):
    count = len(it)
    start = time.time()

    def show(j):
        x = int(size * j / count)
        remaining = ((time.time() - start) / j) * (count - j) if j else 0
        mins, sec = divmod(remaining, 60)
        time_str = f"{int(mins):02}:{sec:03.1f}"
        print(
            f"{prefix}[{'█' * x}{'.' * (size - x)}] {j}/{count} Est wait {time_str}",
            end="\r",
            file=out,
            flush=True,
        )

    if len(it) > 0:
        show(0.1)
        for i, item in enumerate(it):
            yield item
            show(i + 1)
        print(flush=True, file=out)


# -------------------- Name Utilities -------------------- #
def clean_name(name):
    name = re.sub(r"[.:/\\*?\"<>|]", "", name)
    name = re.sub(r"\s+", " ", name).strip().title()
    return name


def format_name(name, lessor=False):
    name = clean_name(name)
    parts = name.split(" ")

    # Return as-is if 27 chars with ≥4 words, or ends with Inc/Ltd
    if (len(name) == 27 and len(parts) >= 4) or re.search(
        r"(Inc\.?|Ltd\.?)$", name, re.IGNORECASE
    ):
        return name

    if lessor:
        # Reverse short lessor names (<4 words and <27 chars)
        if len(parts) < 4 and len(name) < 27:
            return " ".join(parts[::-1])
        # Otherwise, truncate to first 3 words
        return " ".join(parts[:3])

    # Non-lessor names
    if len(parts) == 1:
        return name
    return " ".join(parts[1:] + [parts[0]])


def search_insured_name(full_text_first_page):
    lessor_match = re.search(
        r"\((?:LESSOR|LSR)\)\s*(.*?)\s*\((?:LESSEE|LSE)\)",
        full_text_first_page,
        re.IGNORECASE | re.DOTALL,
    )
    if lessor_match:
        return format_name(lessor_match.group(1), lessor=True)

    match = re.search(
        r"(?:Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))\s*\n([^\n]+)",
        full_text_first_page,
        re.IGNORECASE,
    )
    if match:
        return format_name(match.group(1))

    return None


# -------------------- File Utilities -------------------- #
def safe_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def unique_file_name(path: str) -> str:
    directory = os.path.dirname(path)
    filename, extension = os.path.splitext(os.path.basename(path))
    filename = safe_filename(filename)
    path = os.path.join(directory, filename + extension)
    counter = 1
    while Path(path).is_file():
        path = os.path.join(directory, f"{filename} ({counter}){extension}")
        counter += 1
    return path


def get_base_name(info, use_alt_name=False):
    transaction_timestamp = info.get("transaction_timestamp") or ""
    license_plate = (info.get("license_plate") or "").strip().upper()
    insured_name = (info.get("insured_name") or "").strip()
    insured_name = re.sub(r"[.:/\\*?\"<>|]", "", insured_name)
    insured_name = re.sub(r"\s+", " ", insured_name).strip()
    insured_name = insured_name.title() if insured_name else ""
    transaction_type = (info.get("transaction_type") or "").strip().title()

    top = info.get("top", False)
    storage = info.get("storage", False)
    cancellation = info.get("cancellation", False)
    rental = info.get("rental", False)
    special_risk = info.get("special_risk", False)
    garage = info.get("garage", False)

    if use_alt_name and insured_name:
        if license_plate and license_plate not in ("NONLIC", "STORAGE"):
            base_name = f"{insured_name} - {license_plate}"
        elif insured_name:
            base_name = insured_name
        elif transaction_timestamp:
            base_name = transaction_timestamp
        else:
            base_name = "UNKNOWN"
    else:
        if license_plate and license_plate not in ("NONLIC", "STORAGE"):
            base_name = license_plate
        elif insured_name:
            base_name = insured_name
        elif transaction_timestamp:
            base_name = transaction_timestamp
        else:
            base_name = "UNKNOWN"

    if top:
        base_name = f"{base_name} TOP"
    elif storage:
        base_name = f"{base_name} Storage Policy"
    elif cancellation:
        base_name = f"{base_name} Cancel"
    elif rental:
        base_name = f"{base_name} Rental Policy"
    elif special_risk:
        base_name = f"{base_name} Special Own Risk Damage"
    elif garage:
        base_name = f"{base_name} Garage Policy"
    elif transaction_type == "Change":
        base_name = f"{base_name} Change"
    elif license_plate == "NONLIC":
        base_name = f"{base_name} Registration"

    return base_name


def find_existing_timestamps(base_name, timestamp_pattern, timestamp_rect, folder_dir):
    timestamps = set()
    base_name = base_name.split(" ")[0].upper()

    for pdf_file in Path(folder_dir).glob(f"{base_name}*.pdf"):
        filename = pdf_file.stem.upper().split(" ")[0]
        if filename != base_name:
            continue

        try:
            with fitz.open(pdf_file) as doc:
                if doc.page_count > 0:
                    ts_match = timestamp_pattern.search(
                        doc[0].get_text(clip=timestamp_rect)
                    )
                    if ts_match:
                        timestamps.add(ts_match.group(1))
        except Exception:
            continue

    return timestamps


from pathlib import Path
import openpyxl


from pathlib import Path
import openpyxl


def load_excel_mapping(mapping_path, sheet_index=0, start_row=3):
    mapping_path = Path(mapping_path)
    if not mapping_path.exists():
        return {"b1": None, "b2": None, "producer_mapping": {}}

    wb = openpyxl.load_workbook(mapping_path)
    ws = wb.worksheets[sheet_index]

    input_folder = (
        Path(ws.cell(row=1, column=2).value).expanduser()
        if ws.cell(row=1, column=2).value
        else None
    )
    output_folder = (
        Path(ws.cell(row=2, column=2).value).expanduser()
        if ws.cell(row=2, column=2).value
        else None
    )

    producer_mapping = {
        str(row[0]).upper(): str(row[1])
        for row in ws.iter_rows(min_row=start_row, values_only=True)
        if row[0] and row[1]
    }

    return {
        "b1": input_folder,
        "b2": output_folder,
        "producer_mapping": producer_mapping,
    }


# -------------------- Scan PDFs -------------------- #
# This is the main function that checks for ICBC documents
def scan_icbc_pdfs(
    input_dir,
    regex_patterns,
    page_rects=None,
    max_docs=None,
    stamping_mode=False,
    copy_mode=False,
):
    input_dir = Path(input_dir)
    icbc_data = {}
    non_icbc_file_paths = []

    pdfs = sorted(
        input_dir.rglob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True
    )
    if max_docs:
        pdfs = pdfs[:max_docs]

    for pdf_path in progressbar(pdfs, prefix="Reading PDFs: ", size=10):
        try:
            with fitz.open(pdf_path) as doc:
                if doc.page_count == 0:
                    non_icbc_file_paths.append(str(pdf_path))
                    continue

                page = doc[0]

                def text(clip_name=None):
                    if page_rects and clip_name in page_rects:
                        return (page.get_text(clip=page_rects[clip_name]) or "").strip()
                    return (page.get_text() or "").strip()

                full_text = text()

                # Temporary workaround to not copy payment plan and payment plan receipts
                if not stamping_mode:
                    if (
                        "payment_plan" in regex_patterns
                        and regex_patterns["payment_plan"].search(full_text)
                    ) or (
                        "payment_plan_receipt" in regex_patterns
                        and regex_patterns["payment_plan_receipt"].search(full_text)
                    ):
                        continue
                # ======================================================
                # 🟢 PRIMARY SEARCH
                # ======================================================
                ts_match = regex_patterns["timestamp"].search(full_text)
                if not ts_match:
                    non_icbc_file_paths.append(str(pdf_path))
                    continue
                timestamp = ts_match.group(1)

                license_match = regex_patterns["license_plate"].search(full_text)
                license_plate = (
                    license_match.group(1).strip().upper() if license_match else None
                )

                insured_name = search_insured_name(full_text)

                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "producer_name": None,
                }

                # ======================================================
                # 🟢 STAMPING MODE
                # ======================================================
                if stamping_mode:
                    agency_number = None
                    if "agency_number" in regex_patterns:
                        agency_match = regex_patterns["agency_number"].search(full_text)
                        agency_number = (
                            agency_match.group(1).strip() if agency_match else "UNKNOWN"
                        )

                    customer_copy_pages = []
                    validation_stamp_coords = []
                    time_of_validation_coords = []

                    for page_num, p in enumerate(doc):
                        if "customer_copy" in regex_patterns:
                            if regex_patterns["customer_copy"].search(p.get_text()):
                                customer_copy_pages.append(page_num)

                        for block in p.get_text("blocks"):
                            x0, y0, x1, y1, block_text = *block[:4], block[4]
                            if "validation_stamp" in regex_patterns and regex_patterns[
                                "validation_stamp"
                            ].search(block_text):
                                validation_stamp_coords.append(
                                    (page_num, (x0, y0, x1, y1))
                                )
                            if (
                                "time_of_validation" in regex_patterns
                                and regex_patterns["time_of_validation"].search(
                                    block_text
                                )
                            ):
                                time_of_validation_coords.append(
                                    (page_num, (x0, y0, x1, y1))
                                )

                    icbc_data[pdf_path].update(
                        {
                            "agency_number": agency_number,
                            "customer_copy_pages": customer_copy_pages,
                            "validation_stamp_coords": validation_stamp_coords,
                            "time_of_validation_coords": time_of_validation_coords,
                        }
                    )

                # ======================================================
                # 🔹 COPY MODE
                # ======================================================
                if copy_mode:

                    producer_name = None
                    if "producer" in regex_patterns:
                        producer_match = regex_patterns["producer"].search(
                            text("producer")
                        )
                        producer_name = (
                            producer_match.group(1).upper() if producer_match else None
                        )

                    transaction_type = None
                    if "transaction_type" in regex_patterns:
                        trans_match = regex_patterns["transaction_type"].search(
                            full_text
                        )
                        transaction_type = (
                            trans_match.group(1).strip().title()
                            if trans_match
                            else None
                        )

                    icbc_data[pdf_path].update(
                        {
                            "producer_name": producer_name,
                            "transaction_type": transaction_type,
                            "top": bool(
                                regex_patterns.get("temporary_permit")
                                and regex_patterns["temporary_permit"].search(full_text)
                            ),
                            "storage": bool(
                                regex_patterns.get("storage_policy")
                                and regex_patterns["storage_policy"].search(full_text)
                            ),
                            "cancellation": bool(
                                regex_patterns.get("cancellation")
                                and regex_patterns["cancellation"].search(full_text)
                            ),
                            "special_risk": bool(
                                regex_patterns.get("special_risk_own_damage_policy")
                                and regex_patterns[
                                    "special_risk_own_damage_policy"
                                ].search(full_text)
                            ),
                            "rental": bool(
                                regex_patterns.get("rental_vehicle_policy")
                                and regex_patterns["rental_vehicle_policy"].search(
                                    full_text
                                )
                            ),
                            "garage": bool(
                                regex_patterns.get("garage_vehicle_certificate")
                                and regex_patterns["garage_vehicle_certificate"].search(
                                    full_text
                                )
                            ),
                        }
                    )

        except Exception as e:
            print(f"⚠️  Error processing {pdf_path.name}: {e}")

    return icbc_data, non_icbc_file_paths


# -------------------- Copy PDFs -------------------- #
def copy_pdfs(
    icbc_data,
    output_root_dir,
    producer_mapping=None,
    create_subfolders=False,
    regex_patterns=None,
    page_rects=None,
    use_alt_name=False,
):
    output_root_dir = Path(output_root_dir)
    producer_mapping = producer_mapping or {}
    copied_count = 0

    items_to_process = list(reversed(list(icbc_data.items())))

    for path, info in progressbar(items_to_process, prefix="Copying PDFs: ", size=10):

        if not regex_patterns or "timestamp" not in regex_patterns:
            continue
        if not page_rects or "timestamp" not in page_rects:
            continue

        producer_name = info.get("producer_name")
        if producer_name and producer_name in producer_mapping:
            producer_folder_name = safe_filename(producer_mapping[producer_name])
            subfolder_path = output_root_dir / producer_folder_name
        else:
            subfolder_path = output_root_dir

        if create_subfolders:
            subfolder_path.mkdir(parents=True, exist_ok=True)
        else:
            if not subfolder_path.exists():
                subfolder_path = output_root_dir

        base_name = get_base_name(info, use_alt_name)
        base_name = safe_filename(base_name)
        prefix_name = base_name.split(" ")[0]
        dest_file = subfolder_path / f"{base_name}{path.suffix}"

        # ----------------- Check for duplicates recursively in root folder ----------------- #
        duplicate_found = False
        timestamp_regex = regex_patterns["timestamp"]
        timestamp_rect = page_rects["timestamp"]

        for existing_file in output_root_dir.rglob(f"{prefix_name}*.pdf"):
            try:
                with fitz.open(existing_file) as doc:
                    if doc.page_count > 0:
                        page_text = doc[0].get_text(clip=timestamp_rect)
                        ts_match = timestamp_regex.search(page_text)
                        if ts_match and ts_match.group(1) == info.get(
                            "transaction_timestamp"
                        ):
                            duplicate_found = True
                            break
            except Exception:
                continue

        # ----------------- Copy if not duplicate ----------------- #
        if not duplicate_found:
            dest_file = Path(unique_file_name(str(dest_file)))
            try:
                shutil.copy2(path, dest_file)
                copied_count += 1
            except Exception as e:
                print(f"⚠️ Failed to copy '{path.name}': {e}")

    return copied_count
