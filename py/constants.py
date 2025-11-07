from pathlib import Path
import re
import fitz

# -----------------------------
# Modular regex patterns
# -----------------------------
ICBC_PATTERNS = {
    "timestamp": re.compile(r"Transaction Timestamp\s*(\d+)"),
    "license_plate": re.compile(
        r"Licence Plate Number\s*([A-Z0-9\- ]+)", re.IGNORECASE
    ),
    "temporary_permit": re.compile(
        r"Temporary Operation Permit and Owner’s Certificate of Insurance",
        re.IGNORECASE,
    ),
    "agency_number": re.compile(r"Agency Number\s*[:#]?\s*([A-Z0-9]+)", re.IGNORECASE),
    "customer_copy": re.compile(r"customer copy", re.IGNORECASE),
    "validation_stamp": re.compile(r"NOT VALID UNLESS STAMPED BY", re.IGNORECASE),
    "time_of_validation": re.compile(r"TIME OF VALIDATION", re.IGNORECASE),
    "producer": re.compile(r"-\s*([A-Za-z]+)\s*-", re.IGNORECASE),
    "transaction_type": re.compile(r"Transaction Type\s+([A-Z]+)", re.IGNORECASE),
    "cancellation": re.compile(r"Application for Cancellation"),
    "storage_policy": re.compile(r"Storage Policy"),
    "rental_vehicle_policy": re.compile(r"Rental Vehicle Policy"),
    "special_risk_own_damage_policy": re.compile(r"Special Risk Own Damage Policy"),
    "garage_vehicle_certificate": re.compile(r"Garage Vehicle Certificate"),
}

# -----------------------------
# Modular page clip rects
# -----------------------------
PAGE_RECTS = {
    "timestamp": fitz.Rect(409.979, 63.8488, 576.0, 83.7455),
    "payment_plan": fitz.Rect(425.402, 35.9664, 557.916, 48.3001),
    "payment_plan_receipt": fitz.Rect(461.071, 37.423, 575.922, 48.423),
    "producer": fitz.Rect(198.0, 761.04, 255.011, 769.977),
    "customer_copy": fitz.Rect(498.438, 751.953, 578.181, 769.977),
}
