# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Desktop tool for ICBC (Insurance Corporation of British Columbia) insurance brokerages. It auto-detects ICBC policy document PDFs in the user's Downloads folder, stamps them with validation info (agency number + date/time), and backs up originals to a shared network folder organized by producer codes.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the tool directly (requires config.xlsx in CWD)
python py/icbc_e-stamp_and_copy_tool.py

# Build standalone Windows executable
pyinstaller --onefile --windowed --icon=icon.ico --name "icbc_e-stamp_and_copy_tool" py/icbc_e-stamp_and_copy_tool.py
```

There are no tests, no linter, and no type checker configured.

## Architecture

Two Python files — all business logic lives in `utils.py`, the CLI dispatcher is in `icbc_e-stamp_and_copy_tool.py`.

**Entry point** (`py/icbc_e-stamp_and_copy_tool.py`):
- Reads `config.xlsx` cell B3 to determine which mode to run
- `"Create ICBC Copies Folder Tool"` → `create_icbc_folder_tool()` — one-time setup: bulk-copies PDFs from Downloads to a new shared folder, generates `log.txt`
- Anything else (including empty) → `icbc_e_stamp_tool()` — daily use: stamps + copies the 10 most recently modified PDFs

**Business logic** (`py/utils.py`, ~1140 lines):
- **`ICBCDocument` dataclass** — central data model carrying all extracted PDF metadata (timestamp, plate, insured name, producer code, policy flags, stamping coordinates)
- **`scan_icbc_pdfs()`** — parallel PDF scanner using `ThreadPoolExecutor` (up to 8 workers). Determines if a PDF is an ICBC document by searching page 1 for "Transaction Timestamp" + 14-digit number in a specific page rect. Also extracts: license plate, insured name, producer two-code, policy flags (TOP, Storage, Cancel, etc.), and stamping coordinates for validation stamp placement
- **Name formatting** (`_format_insured_name()`) — ICBC prints names surname-first. Reversal is attempted only when confident the name belongs to a person (not a company). Logic: checks for BC Driver's Licence presence → if licence number is masked (****123), treats as person and reverses; if no licence, falls back to checking for corporate suffixes (Inc/Ltd/Corp). For 4+ part names with a recognized Chinese surname, applies compound surname rules. Particles (de/van/von) are also handled
- **`validation_stamp()`** / **`stamp_time_of_validation()`** — inserts text boxes into the PDF at detected "NOT VALID UNLESS STAMPED BY" and "TIME OF VALIDATION" coordinates using pymupdf
- **`copy_pdfs()`** — copies PDFs into producer subfolders (matched by producer two-code), with duplicate detection using insured name + transaction timestamp
- **`match_pdfs()`** — for files missing a producer two-code, searches existing backup for matching insured names and routes to the correct producer folder
- **`auto_archive()`** — moves files older than N years into `_Archive/YYYY/` subfolders. Can use either filename timestamp or last-modified date (controlled by `DEFAULTS["archive_by_timestamp"]`)
- **`reincrement_pdfs()`** — re-numbers duplicate filenames (e.g., `file (1).pdf`, `file (2).pdf`) after archiving

**Config** (`config.xlsx`, not in repo):
- B3: tool mode selector
- B7: input folder path (for Create ICBC Copies Folder Tool)
- B9: output folder path (for Create ICBC Copies Folder Tool)
- B13: shared ICBC copies folder path (for E-Stamp and Copy Tool)
- B15: agency number (for Certificate Replacement / Same Day Reprint)
- Row 18+: producer two-code → folder name mappings

## Key Patterns

- Duplicate prevention: files are deduplicated by `(insured_name_prefix, transaction_timestamp)` pair across all operations
- Customer copies: only pages tagged "Customer Copy" (plus the last page for TOP policies) are kept; all other pages deleted before saving
- Thread safety in scanning: a `threading.Lock` protects the shared progress counter
- Stamping coordinates are detected per-PDF by searching text blocks — not hardcoded positions
