import re
import os
import shutil
import time
import sys
from pathlib import Path
import fitz
import openpyxl
from collections import defaultdict


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
def reverse_insured_name(name):
    if not name:
        return ""
    name = re.sub(r"\s+", " ", name.strip())
    if name.endswith(("Ltd", "Inc")):
        return name
    name = name.replace(",", "")
    parts = name.split(" ")
    if len(parts) == 1:
        return name
    return " ".join(parts[1:] + [parts[0]])


def search_insured_name(full_text_first_page):
    match = re.search(
        r"(?:Owner\s|Applicant|Name of Insured \(surname followed by given name\(s\)\))\s*\n([^\n]+)",
        full_text_first_page,
        re.IGNORECASE,
    )
    if match:
        name = match.group(1)
        name = re.sub(r"[.:/\\*?\"<>|]", "", name)
        name = re.sub(r"\s+", " ", name).strip().title()
        return name
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


def get_base_name(info, add_suffix=False):
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

    if license_plate and license_plate not in ("NONLIC", "STORAGE"):
        base_name = license_plate
    elif insured_name:
        base_name = insured_name
    elif transaction_timestamp:
        base_name = transaction_timestamp
    else:
        base_name = "UNKNOWN"

    if add_suffix:
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
    base_name = base_name.strip().upper()

    for pdf_file in Path(folder_dir).glob("*.pdf"):
        filename = pdf_file.stem.upper()
        if not filename.startswith(base_name):
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


def load_excel_mapping(mapping_path, ws_index=None):
    mapping_path = Path(mapping_path)
    producer_mapping = {}
    input_folder = None
    output_folder = None

    if not mapping_path.exists():
        return {
            "input_folder": input_folder,
            "output_folder": output_folder,
            "producer_mapping": producer_mapping,
        }

    wb = openpyxl.load_workbook(mapping_path)
    ws = (
        wb.active
        if ws_index is None or ws_index == "active"
        else wb.worksheets[ws_index]
    )

    # Set input folder from cell B1
    root_cell_value = ws.cell(row=1, column=2).value
    if root_cell_value:
        input_folder = Path(root_cell_value).expanduser()

    # Set output folder only if ws_index == 1 (from cell B2)
    if ws_index == 1:
        output_cell_value = ws.cell(row=2, column=2).value
        if output_cell_value:
            output_folder = Path(output_cell_value).expanduser()

    # Decide starting row based on ws_index
    min_row = 4 if ws_index == 1 else 3

    # Populate producer mapping
    for row in ws.iter_rows(min_row=min_row, values_only=True):
        producer, folder = row[:2]
        if producer and folder:
            producer_mapping[str(producer).upper()] = str(folder)

    return {
        "input_folder": input_folder,
        "output_folder": output_folder,
        "producer_mapping": producer_mapping,
    }


# -------------------- Scan PDFs -------------------- #
# This function is used for copying pdfs
def scan_icbc_pdfs(
    input_dir, regex_patterns, page_rects=None, max_docs=None, suffix_mode=False
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

                # Mandatory fields
                ts_match = regex_patterns["timestamp"].search(full_text)
                if not ts_match:
                    non_icbc_file_paths.append(str(pdf_path))
                    continue
                timestamp = ts_match.group(1)

                license_match = regex_patterns["license_plate"].search(full_text)
                license_plate = (
                    license_match.group(1).strip().upper() if license_match else None
                )

                insured_name = reverse_insured_name(search_insured_name(full_text))

                icbc_data[pdf_path] = {
                    "transaction_timestamp": timestamp,
                    "license_plate": license_plate,
                    "insured_name": insured_name,
                    "producer_name": None,
                }

                if suffix_mode:
                    # Producer name from clipped region
                    producer_name = None
                    if "producer" in regex_patterns:
                        producer_match = regex_patterns["producer"].search(
                            text("producer")
                        )
                        producer_name = (
                            producer_match.group(1).upper() if producer_match else None
                        )

                    # Transaction type
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

                    # Boolean flags
                    icbc_data[pdf_path].update(
                        {
                            "producer_name": producer_name,
                            "transaction_type": transaction_type,
                            "top": bool(
                                regex_patterns.get("temporary_permit", None)
                                and regex_patterns["temporary_permit"].search(full_text)
                            ),
                            "storage": bool(
                                regex_patterns.get("storage_policy", None)
                                and regex_patterns["storage_policy"].search(full_text)
                            ),
                            "cancellation": bool(
                                regex_patterns.get("cancellation", None)
                                and regex_patterns["cancellation"].search(full_text)
                            ),
                            "special_risk": bool(
                                regex_patterns.get(
                                    "special_risk_own_damage_policy", None
                                )
                                and regex_patterns[
                                    "special_risk_own_damage_policy"
                                ].search(full_text)
                            ),
                            "rental": bool(
                                regex_patterns.get("rental_vehicle_policy", None)
                                and regex_patterns["rental_vehicle_policy"].search(
                                    full_text
                                )
                            ),
                            "garage": bool(
                                regex_patterns.get("garage_vehicle_certificate", None)
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
    icbc_data, output_root_dir, producer_mapping=None, create_subfolders=False
):
    output_root_dir = Path(output_root_dir)
    producer_mapping = producer_mapping or {}
    copied_count = 0
    seen_transactions_global = (
        set()
    )  # Track all transaction timestamps copied in this batch

    # --- Pre-scan output directory and organize by producer folder ---
    existing_transactions_by_folder = defaultdict(lambda: defaultdict(set))
    for pdf_file in output_root_dir.rglob("*.pdf"):
        folder = pdf_file.parent
        base_name = pdf_file.stem.split(" ")[0]  # assuming base_name logic
        ts = find_existing_timestamps(pdf_file)
        if ts is not None:
            existing_transactions_by_folder[folder][base_name].add(ts)

    # --- Group items by producer folder ---
    folder_to_items = defaultdict(list)
    items_to_process = list(reversed(list(icbc_data.items())))
    for path, info in items_to_process:
        producer_name = info.get("producer_name")
        if producer_name and producer_name in producer_mapping:
            producer_folder_name = safe_filename(producer_mapping[producer_name])
            subfolder_path = output_root_dir / producer_folder_name
        else:
            subfolder_path = output_root_dir
        folder_to_items[subfolder_path].append((path, info))

    # --- Process each folder in input order ---
    for subfolder_path, items in folder_to_items.items():
        # Handle folder existence
        if not subfolder_path.exists():
            if create_subfolders:
                try:
                    subfolder_path.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    print(f"⚠️ Failed to create folder '{subfolder_path}': {e}")
                    continue
            else:
                print(f"⚠️ Skipping folder '{subfolder_path}': does not exist.")
                continue

        for path, info in progressbar(
            items, prefix=f"Copying PDFs to {subfolder_path}: ", size=10
        ):
            base_name = get_base_name(info)
            transaction_ts = info.get("transaction_timestamp")
            dest_file = subfolder_path / f"{base_name}{path.suffix}"

            # Skip duplicates: batch or existing in this folder
            if transaction_ts in seen_transactions_global:
                continue
            if transaction_ts in existing_transactions_by_folder[subfolder_path].get(
                base_name, set()
            ):
                continue

            # Copy the file
            dest_file = Path(unique_file_name(str(dest_file)))
            try:
                shutil.copy2(path, dest_file)
                copied_count += 1
                seen_transactions_global.add(transaction_ts)
                existing_transactions_by_folder[subfolder_path][base_name].add(
                    transaction_ts
                )
            except Exception as e:
                print(f"⚠️ Failed to copy '{path.name}': {e}")

    return copied_count
