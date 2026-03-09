# test_utils.py
# Run with: pytest test_utils.py -v

import re
import shutil
import fitz
import pytest
from pathlib import Path
from unittest.mock import patch
from datetime import datetime

from utils import (
    ICBCDocument,
    _sanitise,
    _format_insured_name,
    extract_insured_name,
    unique_file_path,
    _file_key,
    _extract_filename_timestamp,
    _extract_base_fields,
    _extract_copy_fields,
    copy_pdfs,
    auto_archive,
    reincrement_pdfs,
    save_customer_copy,
    save_batch_copy,
    validation_stamp,
    stamp_time_of_validation,
    load_excel_mapping,
    ICBC_PATTERNS,
)


# ═══════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════


def _make_fitz_pdf(num_pages: int = 1) -> fitz.Document:
    doc = fitz.open()
    for _ in range(num_pages):
        doc.new_page()
    return doc


def _make_pdf_file(path: Path, num_pages: int = 1, text: str = "") -> Path:
    doc = fitz.open()
    for _ in range(num_pages):
        page = doc.new_page()
        if text:
            page.insert_text((50, 50), text)
    doc.save(path)
    doc.close()
    return path


def _make_doc(
    tmp_path: Path,
    *,
    timestamp: str = "20240101120000",
    plate: str | None = "ABC123",
    name: str | None = "John Smith",
    producer: str | None = None,
    **flags,
) -> tuple[Path, ICBCDocument]:
    src = tmp_path / "input"
    src.mkdir(exist_ok=True)
    pdf = _make_pdf_file(src / "test.pdf")
    doc = ICBCDocument(
        path=pdf,
        transaction_timestamp=timestamp,
        license_plate=plate,
        insured_name=name,
        producer_name=producer,
        **flags,
    )
    return pdf, doc


# ═══════════════════════════════════════════════════════════════════
#  _sanitise
# ═══════════════════════════════════════════════════════════════════


def test_sanitise_strips_whitespace():
    assert _sanitise("  hello  ") == "hello"


def test_sanitise_collapses_internal_spaces():
    assert _sanitise("hello   world") == "hello world"


@pytest.mark.parametrize("char", list('.:/\\*?"<>|'))
def test_sanitise_removes_invalid_chars(char):
    assert char not in _sanitise(f"file{char}name")


def test_sanitise_empty_string():
    assert _sanitise("") == ""


# ═══════════════════════════════════════════════════════════════════
#  _format_insured_name
# ═══════════════════════════════════════════════════════════════════


def test_format_reverses_surname_firstname():
    assert _format_insured_name("Smith John") == "John Smith"


def test_format_no_reversal_when_bcdl_string_no_number():
    assert (
        _format_insured_name("Smith John", has_bcdl_string=True, has_bcdl_number=False)
        == "Smith John"
    )


def test_format_no_reversal_for_company_name():
    assert _format_insured_name("Acme Holdings Inc.") == "Acme Holdings Inc."


def test_format_single_word_unchanged():
    assert _format_insured_name("Madonna") == "Madonna"


def test_format_applies_title_case():
    result = _format_insured_name("SMITH JOHN")
    assert result == result.title()


# ═══════════════════════════════════════════════════════════════════
#  extract_insured_name
# ═══════════════════════════════════════════════════════════════════


def test_extract_from_lessor():
    text = "(LESSOR) Acme Leasing Ltd.\nsome other text"
    assert extract_insured_name(text) == "Acme Leasing Ltd."


def test_extract_from_owner_line():
    text = "Owner \nSmith John\nmore text"
    assert extract_insured_name(text) == "John Smith"


def test_extract_from_applicant_line():
    text = "Applicant\nDoe Jane\nmore text"
    assert extract_insured_name(text) == "Jane Doe"


def test_extract_returns_none_when_no_match():
    assert extract_insured_name("No relevant content here") is None


def test_extract_prefers_lessor_over_owner():
    text = "(LESSOR) Acme Corp\nOwner \nSmith John"
    result = extract_insured_name(text)
    assert "Acme" in result


# ═══════════════════════════════════════════════════════════════════
#  unique_file_path
# ═══════════════════════════════════════════════════════════════════


def test_unique_path_no_conflict(tmp_path):
    p = tmp_path / "file.pdf"
    assert unique_file_path(p) == p


def test_unique_path_one_conflict(tmp_path):
    p = tmp_path / "file.pdf"
    p.touch()
    result = unique_file_path(p)
    assert result == tmp_path / "file (1).pdf"


def test_unique_path_two_conflicts(tmp_path):
    (tmp_path / "file.pdf").touch()
    (tmp_path / "file (1).pdf").touch()
    result = unique_file_path(tmp_path / "file.pdf")
    assert result == tmp_path / "file (2).pdf"


def test_unique_path_strips_existing_counter(tmp_path):
    (tmp_path / "file (1).pdf").touch()
    result = unique_file_path(tmp_path / "file (1).pdf")
    # Should not produce "file (1) (1).pdf"
    assert "(1) (1)" not in result.name


# ═══════════════════════════════════════════════════════════════════
#  _file_key
# ═══════════════════════════════════════════════════════════════════


def test_file_key_takes_part_before_dash():
    assert _file_key("John Smith - ABC123") == "JOHN SMITH"


def test_file_key_falls_back_to_first_word():
    assert _file_key("ABC123 something") == "ABC123"


def test_file_key_uppercases():
    assert _file_key("john smith - plate") == "JOHN SMITH"


# ═══════════════════════════════════════════════════════════════════
#  _extract_filename_timestamp
# ═══════════════════════════════════════════════════════════════════


def test_extract_ts_from_filename():
    p = Path("John Smith - ABC123 [20240101120000].pdf")
    assert _extract_filename_timestamp(p) == "20240101120000"


def test_extract_ts_returns_none_when_absent():
    p = Path("John Smith - ABC123.pdf")
    assert _extract_filename_timestamp(p) is None


# ═══════════════════════════════════════════════════════════════════
#  ICBCDocument — base_name / stamp_name
# ═══════════════════════════════════════════════════════════════════


def test_base_name_uses_name_and_plate():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
    )
    assert doc.base_name() == "John Smith - ABC123"


def test_base_name_falls_back_to_name_when_plate_nonlic():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="NONLIC",
        insured_name="Smith John",
    )
    assert "John Smith" in doc.base_name()
    assert "NONLIC" not in doc.base_name()


def test_base_name_falls_back_to_timestamp():
    doc = ICBCDocument(path=Path("x.pdf"), transaction_timestamp="20240101120000")
    assert doc.base_name() == "20240101120000"


def test_base_name_appends_cancel():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
        cancellation=True,
    )
    assert doc.base_name().endswith(" Cancel")


def test_base_name_appends_change():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
        transaction_type="Change",
    )
    assert doc.base_name().endswith(" Change")


def test_base_name_first_flag_wins():
    # top takes priority over storage
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
        top=True,
        storage=True,
    )
    assert "Top" in doc.base_name()
    assert "Storage" not in doc.base_name()


def test_stamp_name_omits_cancel():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        cancellation=True,
    )
    assert "Cancel" not in doc.stamp_name()


def test_stamp_name_uses_plate_only():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
    )
    assert doc.stamp_name() == "ABC123"


def test_registration_suffix_for_nonlic():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="NONLIC",
        insured_name="Smith John",
    )
    assert doc.base_name().endswith("- Registration")


def test_no_registration_suffix_for_storage_plate():
    doc = ICBCDocument(
        path=Path("x.pdf"),
        transaction_timestamp="20240101120000",
        license_plate="STORAGE",
        insured_name="Smith John",
    )
    assert "Registration" not in doc.base_name()


# ═══════════════════════════════════════════════════════════════════
#  _extract_base_fields
# ═══════════════════════════════════════════════════════════════════


def test_extract_base_finds_timestamp():
    text = "Transaction Timestamp 20240101120000\nLicence Plate Number ABC123"
    ts, cert, same, plate, name, top = _extract_base_fields(text, ICBC_PATTERNS)
    assert ts == "20240101120000"
    assert plate == "ABC123"
    assert top is False


def test_extract_base_prefers_cert_replacement():
    text = (
        "Transaction Timestamp 20240101120000\n"
        "Certificate Replacement 20230601080000\n"
    )
    ts, cert, same, plate, name, top = _extract_base_fields(text, ICBC_PATTERNS)
    assert cert == "20230601080000"


def test_extract_base_raises_on_non_icbc():
    with pytest.raises(ValueError, match="not an ICBC document"):
        _extract_base_fields("Just some random PDF text", ICBC_PATTERNS)


def test_extract_base_sets_top_flag():
    text = (
        "Transaction Timestamp 20240101120000\n"
        "Temporary Operation Permit and Owner\u2019s Certificate of Insurance\n"
    )
    _, _, _, _, _, top = _extract_base_fields(text, ICBC_PATTERNS)
    assert top is True


# ═══════════════════════════════════════════════════════════════════
#  _extract_copy_fields
# ═══════════════════════════════════════════════════════════════════


@pytest.mark.parametrize(
    "flag,text",
    [
        ("storage", "Storage Policy"),
        ("cancellation", "Application for Cancellation"),
        ("special_risk", "Special Risk Own Damage Policy"),
        ("rental", "Rental Vehicle Policy"),
        ("garage", "Garage Vehicle Certificate"),
        ("manuscript", "Manuscript Certificate/Manuscript Policy"),
        ("binder", "Binder for Owner\u2019s Interim Certificate of Insurance"),
    ],
)
def test_extract_copy_detects_flag(flag, text):
    result = _extract_copy_fields(text, producer_text="", patterns=ICBC_PATTERNS)
    assert result[flag] is True


def test_extract_copy_all_false_when_nothing_matches():
    result = _extract_copy_fields(
        "Some generic text", producer_text="", patterns=ICBC_PATTERNS
    )
    for flag in (
        "storage",
        "cancellation",
        "special_risk",
        "rental",
        "garage",
        "manuscript",
        "binder",
    ):
        assert result[flag] is False


def test_extract_copy_transaction_type():
    result = _extract_copy_fields(
        "Transaction Type NEW", producer_text="", patterns=ICBC_PATTERNS
    )
    assert result["transaction_type"] == "New"


# ═══════════════════════════════════════════════════════════════════
#  copy_pdfs
# ═══════════════════════════════════════════════════════════════════


def test_copy_routes_to_producer_subfolder(tmp_path):
    src, doc = _make_doc(tmp_path, producer="JONES")
    output = tmp_path / "output"
    output.mkdir()

    copied = copy_pdfs(
        documents={src: doc},
        output_root_dir=output,
        producer_mapping={"JONES": "Jones Agency"},
    )

    assert len(copied) == 1
    assert copied[0].parent == output / "Jones Agency"


def test_copy_falls_back_to_root_when_no_match(tmp_path):
    src, doc = _make_doc(tmp_path, producer="UNKNOWN")
    output = tmp_path / "output"
    output.mkdir()

    copied = copy_pdfs(
        documents={src: doc},
        output_root_dir=output,
        producer_mapping={"JONES": "Jones Agency"},
    )

    assert len(copied) == 1
    assert copied[0].parent == output


def test_copy_skips_existing_timestamp_on_disk(tmp_path):
    src, doc = _make_doc(tmp_path, producer="JONES")
    output = tmp_path / "output" / "Jones Agency"
    output.mkdir(parents=True)
    # Pre-plant a file with the same timestamp
    (output / "John Smith - ABC123 [20240101120000].pdf").touch()

    copied = copy_pdfs(
        documents={src: doc},
        output_root_dir=tmp_path / "output",
        producer_mapping={"JONES": "Jones Agency"},
    )

    assert len(copied) == 0


def test_copy_skips_duplicate_within_same_run(tmp_path):
    src1, doc1 = _make_doc(tmp_path, producer="JONES")
    # Identical timestamp and name — second copy should be skipped
    src2 = tmp_path / "input" / "test2.pdf"
    shutil.copy(src1, src2)
    doc2 = ICBCDocument(
        path=src2,
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
        producer_name="JONES",
    )
    output = tmp_path / "output"
    output.mkdir()

    copied = copy_pdfs(
        documents={src1: doc1, src2: doc2},
        output_root_dir=output,
        producer_mapping={"JONES": "Jones Agency"},
    )

    assert len(copied) == 1


def test_copy_filename_includes_timestamp(tmp_path):
    src, doc = _make_doc(tmp_path)
    output = tmp_path / "output"
    output.mkdir()

    copied = copy_pdfs(documents={src: doc}, output_root_dir=output)

    assert "[20240101120000]" in copied[0].name


def test_copy_ignores_archive_when_flag_set(tmp_path):
    src, doc = _make_doc(tmp_path)
    output = tmp_path / "output"
    archive = output / "_Archive"
    archive.mkdir(parents=True)
    # Pre-plant a file in archive with same timestamp — should be ignored
    (archive / "John Smith - ABC123 [20240101120000].pdf").touch()

    copied = copy_pdfs(
        documents={src: doc},
        output_root_dir=output,
        ignore_archive=True,
    )

    assert len(copied) == 1


# ═══════════════════════════════════════════════════════════════════
#  auto_archive
# ═══════════════════════════════════════════════════════════════════


def test_auto_archive_moves_stale_file(tmp_path):
    pdf = _make_pdf_file(tmp_path / "old.pdf")
    # Set mtime to 3 years ago
    import time

    old_time = time.time() - (3 * 365 * 24 * 3600)
    import os

    os.utime(pdf, (old_time, old_time))

    archived = auto_archive(tmp_path, min_age_years=2)

    assert archived is not None
    assert len(archived) == 1
    assert "_Archive" in str(archived[0])


def test_auto_archive_leaves_recent_file(tmp_path):
    _make_pdf_file(tmp_path / "new.pdf")

    archived = auto_archive(tmp_path, min_age_years=2)

    assert archived is None


def test_auto_archive_does_not_re_archive(tmp_path):
    import os, time

    archive = tmp_path / "_Archive"
    archive.mkdir()
    pdf = _make_pdf_file(archive / "already.pdf")
    old_time = time.time() - (3 * 365 * 24 * 3600)
    os.utime(pdf, (old_time, old_time))

    archived = auto_archive(tmp_path, min_age_years=2)

    assert archived is None


def test_auto_archive_preserves_subfolder_structure(tmp_path):
    import os, time

    sub = tmp_path / "Jones Agency"
    sub.mkdir()
    pdf = _make_pdf_file(sub / "old.pdf")
    old_time = time.time() - (3 * 365 * 24 * 3600)
    os.utime(pdf, (old_time, old_time))

    archived = auto_archive(tmp_path, min_age_years=2)

    assert "Jones Agency" in str(archived[0])


# ═══════════════════════════════════════════════════════════════════
#  reincrement_pdfs
# ═══════════════════════════════════════════════════════════════════


def test_reincrement_leaves_single_file_alone(tmp_path):
    (tmp_path / "file.pdf").touch()
    reincrement_pdfs(tmp_path)
    assert (tmp_path / "file.pdf").exists()


def test_reincrement_renumbers_group(tmp_path):
    (tmp_path / "file (3).pdf").touch()
    (tmp_path / "file (5).pdf").touch()
    reincrement_pdfs(tmp_path)
    assert (tmp_path / "file.pdf").exists()
    assert (tmp_path / "file (1).pdf").exists()
    assert not (tmp_path / "file (3).pdf").exists()


def test_reincrement_removes_empty_subfolders(tmp_path):
    sub = tmp_path / "empty_sub"
    sub.mkdir()
    (sub / "file (1).pdf").touch()
    (sub / "file (2).pdf").touch()
    reincrement_pdfs(tmp_path)
    # After renaming, sub still has files so should remain
    assert sub.exists()


def test_reincrement_handles_nonexistent_dir(tmp_path):
    # Should not raise
    reincrement_pdfs(tmp_path / "does_not_exist")


# ═══════════════════════════════════════════════════════════════════
#  save_customer_copy
# ═══════════════════════════════════════════════════════════════════


def test_customer_copy_retains_correct_pages(tmp_path):
    doc = _make_fitz_pdf(3)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        customer_copy_pages=[1],
    )
    save_customer_copy(doc, document, tmp_path)
    saved = fitz.open(tmp_path / f"{document.stamp_name()} (Customer Copy).pdf")
    assert saved.page_count == 1


def test_customer_copy_top_appends_last_page(tmp_path):
    doc = _make_fitz_pdf(3)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        top=True,
        customer_copy_pages=[1],
        license_plate="ABC123",
    )
    save_customer_copy(doc, document, tmp_path)
    saved = fitz.open(tmp_path / f"{document.stamp_name()} (Customer Copy).pdf")
    assert saved.page_count == 2


def test_customer_copy_top_no_duplicate_last_page(tmp_path):
    doc = _make_fitz_pdf(3)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        top=True,
        customer_copy_pages=[1, 2],  # last page already included
        license_plate="ABC123",
    )
    save_customer_copy(doc, document, tmp_path)
    saved = fitz.open(tmp_path / f"{document.stamp_name()} (Customer Copy).pdf")
    assert saved.page_count == 2


# ═══════════════════════════════════════════════════════════════════
#  save_batch_copy
# ═══════════════════════════════════════════════════════════════════


def test_save_batch_copy_creates_folder(tmp_path):
    doc = _make_fitz_pdf(1)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
    )
    save_batch_copy(doc, document, tmp_path)
    assert (tmp_path / "ICBC Batch Copies").is_dir()


def test_save_batch_copy_filename_format(tmp_path):
    doc = _make_fitz_pdf(1)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        license_plate="ABC123",
        insured_name="Smith John",
    )
    dest = save_batch_copy(doc, document, tmp_path)
    assert "[20240101120000]" in dest.name
    assert dest.suffix == ".pdf"


# ═══════════════════════════════════════════════════════════════════
#  validation_stamp
# ═══════════════════════════════════════════════════════════════════


def test_validation_stamp_inserts_text(tmp_path):
    doc = _make_fitz_pdf(1)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        agency_number="12345",
        validation_stamp_coords=[(0, (100.0, 100.0, 200.0, 120.0))],
    )
    ts_dt = datetime(2024, 1, 1, 12, 0, 0)
    result = validation_stamp(doc, document, ts_dt)
    text = result[0].get_text()
    assert "12345" in text
    assert "Jan 01, 2024" in text


def test_validation_stamp_noop_when_no_coords(tmp_path):
    doc = _make_fitz_pdf(1)
    original_text = doc[0].get_text()
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101120000",
        agency_number="12345",
        validation_stamp_coords=[],
    )
    result = validation_stamp(doc, document, datetime(2024, 1, 1))
    assert result[0].get_text() == original_text


# ═══════════════════════════════════════════════════════════════════
#  stamp_time_of_validation
# ═══════════════════════════════════════════════════════════════════


def test_time_of_validation_am(tmp_path):
    doc = _make_fitz_pdf(1)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101090000",
        time_of_validation_coords=[(0, (100.0, 200.0, 200.0, 220.0))],
    )
    ts_dt = datetime(2024, 1, 1, 9, 30, 0)
    result = stamp_time_of_validation(doc, document, ts_dt)
    assert "09:30" in result[0].get_text()


def test_time_of_validation_pm(tmp_path):
    doc = _make_fitz_pdf(1)
    document = ICBCDocument(
        path=tmp_path / "x.pdf",
        transaction_timestamp="20240101150000",
        time_of_validation_coords=[(0, (100.0, 200.0, 200.0, 220.0))],
    )
    ts_dt = datetime(2024, 1, 1, 15, 30, 0)
    result = stamp_time_of_validation(doc, document, ts_dt)
    assert "03:30" in result[0].get_text()


# ═══════════════════════════════════════════════════════════════════
#  load_excel_mapping
# ═══════════════════════════════════════════════════════════════════


def test_load_mapping_returns_defaults_when_no_file(tmp_path):
    with patch("utils.Path.cwd", return_value=tmp_path):
        mapping = load_excel_mapping(tmp_path / "config.xlsx")
    assert mapping.copy_input_folder is None
    assert mapping.output_folder is None


def test_load_mapping_raises_on_missing_sheet(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.title = "wrong_sheet"
    path = tmp_path / "config.xlsx"
    wb.save(path)

    with pytest.raises(ValueError, match="Sheet 'config' not found"):
        load_excel_mapping(path)


# ═══════════════════════════════════════════════════════════════════
#  ICBC_PATTERNS — regex smoke tests
# ═══════════════════════════════════════════════════════════════════


@pytest.mark.parametrize(
    "key,text",
    [
        ("timestamp", "Transaction Timestamp 20240101120000"),
        ("certificate_replacement", "Certificate Replacement 20240101120000"),
        ("same_day_re-print", "Same day Re-print 20240101120000"),
        ("license_plate", "Licence Plate Number ABC 123"),
        ("cancellation", "Application for Cancellation"),
        ("storage_policy", "Storage Policy"),
        ("rental_vehicle_policy", "Rental Vehicle Policy"),
        ("payment_plan", "Payment Plan Agreement"),
        ("payment_plan_receipt", "Payment Plan Receipt"),
        ("binder", "Binder for Owner\u2019s Interim Certificate of Insurance"),
        ("manuscript", "Manuscript Certificate/Manuscript Policy"),
        ("customer_copy", "customer copy"),
        ("validation_stamp", "NOT VALID UNLESS STAMPED BY"),
    ],
)
def test_pattern_matches(key, text):
    assert ICBC_PATTERNS[key].search(text) is not None


def test_payment_plan_does_not_match_receipt():
    # These are separate patterns — receipt should not fire payment_plan
    assert ICBC_PATTERNS["payment_plan"].search("Payment Plan Receipt") is None


def test_binder_matches_straight_apostrophe():
    text = "Binder for Owner's Interim Certificate of Insurance"
    assert ICBC_PATTERNS["binder"].search(text) is not None
