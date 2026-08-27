import csv
import datetime
import re
import sys
from pathlib import Path

import pandas as pd


# ==============================================================================
# KUBO DOCUMENT MATCHING ENGINE — V5.5
# PRODUCTION VERSION
#
# DAILY SAP ↔ ACKNOWLEDGED DELIVERY PDF RECONCILIATION
#
# Purpose:
#   1. Locate the latest valid SAP export
#   2. Filter SAP records to the current year/month
#   3. Extract and normalize document numbers
#   4. Scan the acknowledged delivery PDF folder
#   5. Extract document numbers from PDF filenames
#   6. Reconcile SAP documents against PDFs
#   7. Detect duplicates, missing PDFs and untracked PDFs
#   8. Produce CSV + TXT audit reports
#
# IMPORTANT:
#   - Original SAP files are NEVER modified.
#   - Original PDF files are NEVER modified.
#   - Only reports/logs are written to C:\Kubo.
# ==============================================================================


# ==============================================================================
# CONFIGURATION
# ==============================================================================

NOW = datetime.datetime.now()

CURRENT_YEAR = NOW.year
CURRENT_YEAR_STR = NOW.strftime("%Y")
CURRENT_MONTH = NOW.strftime("%B").upper()
CURRENT_MONTH_NUMBER = NOW.month


# ------------------------------------------------------------------------------
# NETWORK LOCATIONS
# ------------------------------------------------------------------------------

BASE_DIR = Path(r"\\192.168.1.4\B1_Shr")

KUBO_SHARE = Path(r"\\vsvr1\IAL_Kubo$\KUBO")


# ------------------------------------------------------------------------------
# DELIVERY PDF LOCATION
# ------------------------------------------------------------------------------

DELIVERIES_DIR = (
    BASE_DIR
    / "Attachments"
    / "Acknowledged Deliveries"
    / CURRENT_YEAR_STR
    / CURRENT_MONTH
)


# ------------------------------------------------------------------------------
# LOCAL OUTPUT LOCATION
# ------------------------------------------------------------------------------

LOCAL_OUT_DIR = Path(r"C:\Kubo")
LOCAL_OUT_DIR.mkdir(parents=True, exist_ok=True)

MATCH_REPORT_CSV = (
    LOCAL_OUT_DIR
    / "kubo_document_matching_report.csv"
)

SUMMARY_REPORT = (
    LOCAL_OUT_DIR
    / "george_verification_report.txt"
)

LOG_FILE = (
    LOCAL_OUT_DIR
    / "kubo_matching_engine.log"
)


# ==============================================================================
# LOGGING
# ==============================================================================

def log(message: str) -> None:
    """
    Print a timestamped message and append it to the local log.
    """

    timestamp = datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    line = f"[{timestamp}] {message}"

    print(line)

    try:
        with open(
            LOG_FILE,
            "a",
            encoding="utf-8"
        ) as file:

            file.write(line + "\n")

    except Exception:
        # Logging failure must never stop reconciliation.
        pass


# ==============================================================================
# NORMALIZATION
# ==============================================================================

def normalize_document_id(value):
    """
    Normalize SAP/PDF document numbers.

    Examples:
        12345678      -> 12345678
        1234567       -> 01234567
        12345678.0    -> 12345678

    Invalid/non-numeric values return None.
    """

    if pd.isna(value):
        return None

    value = str(value).strip()

    if not value:
        return None

    # Remove Excel-style decimal suffix.
    value = re.sub(
        r"\.0$",
        "",
        value
    )

    if not value:
        return None

    if value.isdigit():
        return value.zfill(8)

    return None


def normalize_month(value):
    """
    Convert common month formats to full uppercase month names.
    """

    if pd.isna(value):
        return None

    value = str(value).strip().upper()

    if not value:
        return None

    month_map = {

        "1": "JANUARY",
        "01": "JANUARY",
        "JAN": "JANUARY",

        "2": "FEBRUARY",
        "02": "FEBRUARY",
        "FEB": "FEBRUARY",

        "3": "MARCH",
        "03": "MARCH",
        "MAR": "MARCH",

        "4": "APRIL",
        "04": "APRIL",
        "APR": "APRIL",

        "5": "MAY",
        "05": "MAY",

        "6": "JUNE",
        "06": "JUNE",
        "JUN": "JUNE",

        "7": "JULY",
        "07": "JULY",
        "JUL": "JULY",

        "8": "AUGUST",
        "08": "AUGUST",
        "AUG": "AUGUST",

        "9": "SEPTEMBER",
        "09": "SEPTEMBER",
        "SEP": "SEPTEMBER",
        "SEPT": "SEPTEMBER",

        "10": "OCTOBER",
        "OCT": "OCTOBER",

        "11": "NOVEMBER",
        "NOV": "NOVEMBER",

        "12": "DECEMBER",
        "DEC": "DECEMBER",
    }

    return month_map.get(
        value,
        value
    )


# ==============================================================================
# COLUMN DETECTION
# ==============================================================================

def find_column(df, possible_names):
    """
    Safely identify a column using exact normalized names.

    This deliberately avoids vague matching such as:
        'id' in column name
        'number' in column name

    because that can select the wrong SAP field.
    """

    normalized_map = {}

    for col in df.columns:

        normalized = (
            str(col)
            .strip()
            .lower()
            .replace("-", " ")
            .replace("_", " ")
        )

        normalized = re.sub(
            r"\s+",
            " ",
            normalized
        )

        normalized_map[normalized] = col

    for name in possible_names:

        normalized_name = (
            name
            .strip()
            .lower()
            .replace("-", " ")
            .replace("_", " ")
        )

        normalized_name = re.sub(
            r"\s+",
            " ",
            normalized_name
        )

        if normalized_name in normalized_map:
            return normalized_map[normalized_name]

    return None


# ==============================================================================
# ENVIRONMENT VALIDATION
# ==============================================================================

def validate_environment():
    """
    Validate the resources that are actually required.

    Important:
    KUBO_SHARE is NOT mandatory because SAP has a CSV fallback.

    BASE_DIR is required because acknowledged delivery PDFs live there.
    """

    log("Validating environment...")

    problems = []

    # Delivery source is mandatory.
    if not BASE_DIR.exists():

        problems.append(
            f"Base network path unavailable: {BASE_DIR}"
        )

    if not DELIVERIES_DIR.exists():

        problems.append(
            f"Delivery folder unavailable: {DELIVERIES_DIR}"
        )

    # KUBO_SHARE is optional because CSV fallback exists.
    if KUBO_SHARE.exists():

        log(
            f"[OK] SAP Kubo share available: {KUBO_SHARE}"
        )

    else:

        log(
            "[WARNING] SAP Kubo share unavailable. "
            "CSV fallback will be attempted."
        )

    if problems:

        for problem in problems:
            log(f"[ERROR] {problem}")

        return False

    log("[OK] Required network environment validated.")

    return True


# ==============================================================================
# SAP SOURCE DISCOVERY
# ==============================================================================

def locate_sap_source():
    """
    Locate the best available SAP source.

    Priority:
        1. Latest 'New SAP Combine*.xlsx' from KUBO_SHARE
        2. Exact current-month SAP CSV from BASE_DIR
        3. Other SAP-looking CSV files from BASE_DIR

    We deliberately avoid blindly selecting any CSV.
    """

    log("Searching for SAP source...")

    # --------------------------------------------------------------------------
    # OPTION 1 — SAP Excel from KUBO share
    # --------------------------------------------------------------------------

    if KUBO_SHARE.exists():

        try:

            excel_files = list(
                KUBO_SHARE.glob(
                    "New SAP Combine*.xlsx"
                )
            )

            if excel_files:

                latest_excel = max(
                    excel_files,
                    key=lambda f: f.stat().st_mtime
                )

                log(
                    f"[OK] Found SAP Excel: "
                    f"{latest_excel}"
                )

                return latest_excel

        except Exception as exc:

            log(
                f"[WARNING] Could not inspect KUBO SAP share: "
                f"{exc}"
            )

    # --------------------------------------------------------------------------
    # OPTION 2 — Exact current-month CSV
    # --------------------------------------------------------------------------

    if BASE_DIR.exists():

        exact_csv = (
            BASE_DIR
            / f"sap_{CURRENT_MONTH.capitalize()}_list.csv"
        )

        if exact_csv.exists():

            log(
                f"[OK] Found current-month SAP CSV: "
                f"{exact_csv}"
            )

            return exact_csv

    # --------------------------------------------------------------------------
    # OPTION 3 — Controlled CSV fallback
    # --------------------------------------------------------------------------

    if BASE_DIR.exists():

        csv_candidates = []

        try:

            for file in BASE_DIR.glob("*.csv"):

                name = file.name.lower()

                # Only consider files that look like SAP exports.
                if (
                    "sap" in name
                    or "combine" in name
                    or "document" in name
                    or "delivery" in name
                ):

                    csv_candidates.append(file)

        except Exception as exc:

            log(
                f"[WARNING] Could not inspect CSV fallback: "
                f"{exc}"
            )

            csv_candidates = []

        if csv_candidates:

            latest_csv = max(
                csv_candidates,
                key=lambda f: f.stat().st_mtime
            )

            log(
                f"[OK] Using controlled CSV fallback: "
                f"{latest_csv}"
            )

            return latest_csv

    log("[ERROR] No suitable SAP source found.")

    return None


# ==============================================================================
# SAP DOCUMENT EXTRACTION
# ==============================================================================

def get_sap_document_ids():

    source = locate_sap_source()

    if source is None:

        return (
            set(),
            [],
            None,
            0,
            0,
            [],
            "No source"
        )

    # --------------------------------------------------------------------------
    # Read SAP file
    # --------------------------------------------------------------------------

    try:

        if source.suffix.lower() == ".xlsx":

            sap_df = pd.read_excel(
                source,
                dtype=str
            )

        elif source.suffix.lower() == ".csv":

            sap_df = pd.read_csv(
                source,
                dtype=str
            )

        else:

            log(
                "[ERROR] Unsupported SAP source format."
            )

            return (
                set(),
                [],
                source.name,
                0,
                0,
                [],
                "Unsupported format"
            )

    except Exception as exc:

        log(
            f"[ERROR] Could not read SAP source: {exc}"
        )

        return (
            set(),
            [],
            source.name,
            0,
            0,
            [],
            "Read failure"
        )

    if sap_df.empty:

        log(
            "[ERROR] SAP source contains no data."
        )

        return (
            set(),
            [],
            source.name,
            0,
            0,
            [],
            "Empty source"
        )

    all_sap_rows = len(sap_df)

    log(
        f"SAP rows loaded: {all_sap_rows}"
    )

    # --------------------------------------------------------------------------
    # Locate document number
    # --------------------------------------------------------------------------

    doc_col = find_column(
        sap_df,
        [
            "DocNum",
            "Doc NO",
            "Doc_No",
            "DocNo",
            "Document No",
            "Document Number",
            "Document No."
        ]
    )

    if doc_col is None:

        log(
            "[FATAL] Could not locate SAP document "
            "number column."
        )

        return (
            set(),
            [],
            source.name,
            all_sap_rows,
            0,
            [],
            "Missing document column"
        )

    log(
        f"SAP document column detected: {doc_col}"
    )

    # --------------------------------------------------------------------------
    # Locate date/filter columns
    # --------------------------------------------------------------------------

    year_col = find_column(
        sap_df,
        ["Year"]
    )

    month_col = find_column(
        sap_df,
        ["Month"]
    )

    date_col = find_column(
        sap_df,
        [
            "DocDate",
            "Doc Date",
            "Document Date"
        ]
    )

    # --------------------------------------------------------------------------
    # Filter by Year + Month where possible
    # --------------------------------------------------------------------------

    if year_col is not None and month_col is not None:

        year_values = (
            sap_df[year_col]
            .astype(str)
            .str.strip()
            .str.replace(
                r"\.0$",
                "",
                regex=True
            )
        )

        month_values = (
            sap_df[month_col]
            .apply(normalize_month)
        )

        filtered_df = sap_df[
            (year_values == CURRENT_YEAR_STR)
            &
            (month_values == CURRENT_MONTH)
        ].copy()

        filtering_method = "Year + Month"

    # --------------------------------------------------------------------------
    # Otherwise use DocDate
    # --------------------------------------------------------------------------

    elif date_col is not None:

        parsed_dates = pd.to_datetime(
            sap_df[date_col],
            errors="coerce",
            dayfirst=True,
            format="mixed"
        )

        filtered_df = sap_df[
            (parsed_dates.dt.year == CURRENT_YEAR)
            &
            (
                parsed_dates.dt.month
                == CURRENT_MONTH_NUMBER
            )
        ].copy()

        filtering_method = "DocDate"

    else:

        log(
            "[FATAL] SAP source contains neither "
            "Year + Month nor DocDate."
        )

        return (
            set(),
            [],
            source.name,
            all_sap_rows,
            0,
            [],
            "No date filtering columns"
        )

    filtered_sap_rows = len(filtered_df)

    log(
        f"SAP filtering method: {filtering_method}"
    )

    log(
        f"SAP rows after "
        f"{CURRENT_MONTH} {CURRENT_YEAR} filter: "
        f"{filtered_sap_rows}"
    )

    if filtered_sap_rows == 0:

        log(
            f"[FATAL] No SAP records found for "
            f"{CURRENT_MONTH} {CURRENT_YEAR}."
        )

        return (
            set(),
            [],
            source.name,
            all_sap_rows,
            0,
            [],
            filtering_method
        )

    # --------------------------------------------------------------------------
    # Normalize document IDs
    # --------------------------------------------------------------------------

    sap_ids = set()
    invalid_sap_ids = []
    duplicate_sap_ids = []

    seen_ids = set()

    for value in filtered_df[doc_col]:

        clean_id = normalize_document_id(value)

        if clean_id is None:

            if pd.notna(value):

                invalid_sap_ids.append(
                    str(value)
                )

            continue

        if clean_id in seen_ids:

            duplicate_sap_ids.append(
                clean_id
            )

        else:

            seen_ids.add(clean_id)
            sap_ids.add(clean_id)

    log(
        f"Unique SAP document IDs: "
        f"{len(sap_ids)}"
    )

    log(
        f"Duplicate SAP records: "
        f"{len(duplicate_sap_ids)}"
    )

    log(
        f"Invalid SAP document values: "
        f"{len(invalid_sap_ids)}"
    )

    return (
        sap_ids,
        invalid_sap_ids,
        source.name,
        all_sap_rows,
        filtered_sap_rows,
        duplicate_sap_ids,
        filtering_method
    )


# ==============================================================================
# PDF DISCOVERY
# ==============================================================================

def find_delivery_pdfs():

    log(
        f"Scanning PDFs in: {DELIVERIES_DIR}"
    )

    if not DELIVERIES_DIR.exists():

        log(
            f"[ERROR] Delivery directory missing: "
            f"{DELIVERIES_DIR}"
        )

        return {}, {}, []

    pdf_dict = {}
    duplicate_dict = {}
    invalid_pdf_files = []

    total_pdf_files = 0

    try:

        for path in DELIVERIES_DIR.rglob("*.pdf"):

            total_pdf_files += 1

            # --------------------------------------------------------------
            # Extract 6-8 digit document number.
            #
            # We take the LAST valid number because filenames can contain
            # other numbers before the actual document number.
            # --------------------------------------------------------------

            matches = re.findall(
                r"(?<!\d)\d{6,8}(?!\d)",
                path.stem.strip()
            )

            if not matches:

                invalid_pdf_files.append(
                    path
                )

                continue

            clean_id = normalize_document_id(
                matches[-1]
            )

            if clean_id is None:

                invalid_pdf_files.append(
                    path
                )

                continue

            # --------------------------------------------------------------
            # First PDF for this ID
            # --------------------------------------------------------------

            if clean_id not in pdf_dict:

                pdf_dict[clean_id] = path

            # --------------------------------------------------------------
            # Additional PDF = duplicate
            # --------------------------------------------------------------

            else:

                if clean_id not in duplicate_dict:

                    duplicate_dict[clean_id] = [
                        pdf_dict[clean_id]
                    ]

                duplicate_dict[clean_id].append(
                    path
                )

    except Exception as exc:

        log(
            f"[ERROR] Interrupted while scanning PDFs: "
            f"{exc}"
        )

    log(
        f"Total PDF files scanned: "
        f"{total_pdf_files}"
    )

    log(
        f"Unique PDF document IDs: "
        f"{len(pdf_dict)}"
    )

    log(
        f"PDF document IDs with duplicates: "
        f"{len(duplicate_dict)}"
    )

    log(
        f"PDF files without valid document IDs: "
        f"{len(invalid_pdf_files)}"
    )

    return (
        pdf_dict,
        duplicate_dict,
        invalid_pdf_files
    )


# ==============================================================================
# RECONCILIATION
# ==============================================================================

def perform_matching(
    sap_ids: set,
    pdf_dict: dict,
    duplicate_dict: dict
):

    log(
        "Performing document reconciliation..."
    )

    results = []

    all_known_ids = (
        sap_ids
        .union(pdf_dict.keys())
        .union(duplicate_dict.keys())
    )

    for doc_id in sorted(all_known_ids):

        in_sap = doc_id in sap_ids
        in_pdf = doc_id in pdf_dict
        is_duplicate = doc_id in duplicate_dict

        # --------------------------------------------------------------
        # Determine status
        # --------------------------------------------------------------

        if in_sap and in_pdf and is_duplicate:

            status = "MATCHED_DUPLICATE_PDF"

        elif in_sap and in_pdf:

            status = "MATCHED"

        elif in_sap and not in_pdf:

            status = "MISSING_PDF"

        elif not in_sap and in_pdf and is_duplicate:

            status = "UNTRACKED_DUPLICATE_PDF"

        elif not in_sap and in_pdf:

            status = "UNTRACKED_PDF"

        else:

            status = "UNKNOWN"

        # --------------------------------------------------------------
        # Build PDF path field
        # --------------------------------------------------------------

        if is_duplicate:

            file_paths = [
                str(path)
                for path in duplicate_dict[doc_id]
            ]

            file_path = "; ".join(
                file_paths
            )

        elif in_pdf:

            file_path = str(
                pdf_dict[doc_id]
            )

        else:

            file_path = "N/A"

        results.append({

            "Document_ID": doc_id,

            "Status": status,

            "In_SAP": (
                "YES"
                if in_sap
                else "NO"
            ),

            "In_PDF": (
                "YES"
                if in_pdf
                else "NO"
            ),

            "Is_Duplicate": (
                "YES"
                if is_duplicate
                else "NO"
            ),

            "File_Path": file_path
        })

    return results


# ==============================================================================
# CSV REPORT
# ==============================================================================

def write_csv_report(results):

    fieldnames = [
        "Document_ID",
        "Status",
        "In_SAP",
        "In_PDF",
        "Is_Duplicate",
        "File_Path"
    ]

    with open(
        MATCH_REPORT_CSV,
        mode="w",
        newline="",
        encoding="utf-8"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=fieldnames
        )

        writer.writeheader()
        writer.writerows(results)

    log(
        f"CSV report created: "
        f"{MATCH_REPORT_CSV}"
    )


# ==============================================================================
# SEQUENCE GAP DETECTION
# ==============================================================================

def find_sequence_gaps(ids):

    numeric_ids = sorted(
        int(x)
        for x in ids
        if str(x).isdigit()
    )

    if len(numeric_ids) < 2:
        return []

    gaps = []

    for current, next_id in zip(
        numeric_ids,
        numeric_ids[1:]
    ):

        difference = (
            next_id - current
        )

        if difference > 1:

            skipped = list(
                range(
                    current + 1,
                    next_id
                )
            )

            gaps.append({

                "gap_after_id": current,

                "gap_before_id": next_id,

                "missing_count": len(
                    skipped
                ),

                "skipped_ids": skipped
            })

    return gaps


# ==============================================================================
# SUMMARY REPORT
# ==============================================================================

def write_summary_report(
    results,
    sap_count,
    pdf_count,
    duplicate_count,
    invalid_pdf_count,
    sap_source,
    invalid_sap_ids,
    all_sap_rows,
    filtered_sap_rows,
    duplicate_sap_ids,
    filtering_method
):

    matched = sum(
        1
        for r in results
        if r["Status"] == "MATCHED"
    )

    matched_duplicate = sum(
        1
        for r in results
        if r["Status"]
        == "MATCHED_DUPLICATE_PDF"
    )

    missing_pdf = sum(
        1
        for r in results
        if r["Status"] == "MISSING_PDF"
    )

    untracked_pdf = sum(
        1
        for r in results
        if r["Status"] == "UNTRACKED_PDF"
    )

    untracked_duplicate = sum(
        1
        for r in results
        if r["Status"]
        == "UNTRACKED_DUPLICATE_PDF"
    )

    # A duplicate PDF is still physically present,
    # therefore it counts as accounted for operationally,
    # but is separately flagged for attention.
    accounted_documents = (
        matched
        + matched_duplicate
    )

    match_rate = (
        accounted_documents
        / sap_count
        * 100
        if sap_count
        else 0
    )

    # --------------------------------------------------------------------------
    # Sequence gaps
    # --------------------------------------------------------------------------

    sap_ids_for_gap = {
        r["Document_ID"]
        for r in results
        if r["In_SAP"] == "YES"
    }

    pdf_ids_for_gap = {
        r["Document_ID"]
        for r in results
        if r["In_PDF"] == "YES"
    }

    sap_gaps = find_sequence_gaps(
        sap_ids_for_gap
    )

    pdf_gaps = find_sequence_gaps(
        pdf_ids_for_gap
    )

    # --------------------------------------------------------------------------
    # Write report
    # --------------------------------------------------------------------------

    with open(
        SUMMARY_REPORT,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            "=" * 78
            + "\n"
        )

        f.write(
            "KUBO DOCUMENT RECONCILIATION SUMMARY\n"
        )

        f.write(
            f"{CURRENT_MONTH} "
            f"{CURRENT_YEAR}\n"
        )

        f.write(
            "=" * 78
            + "\n\n"
        )

        # ------------------------------------------------------------------
        # RUN INFORMATION
        # ------------------------------------------------------------------

        f.write(
            "RUN INFORMATION\n"
            + "-" * 78
            + "\n"
        )

        f.write(
            f"Execution Time       : "
            f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        )

        f.write(
            f"SAP Source           : "
            f"{sap_source}\n"
        )

        f.write(
            f"Delivery Folder      : "
            f"{DELIVERIES_DIR}\n"
        )

        f.write(
            f"Filtering Method     : "
            f"{filtering_method}\n\n"
        )

        # ------------------------------------------------------------------
        # SAP DATA
        # ------------------------------------------------------------------

        f.write(
            "SAP DATA\n"
            + "-" * 78
            + "\n"
        )

        f.write(
            f"Total SAP Rows Loaded          : "
            f"{all_sap_rows}\n"
        )

        f.write(
            f"SAP Rows After Month Filter   : "
            f"{filtered_sap_rows}\n"
        )

        f.write(
            f"Unique SAP Documents           : "
            f"{sap_count}\n"
        )

        f.write(
            f"Duplicate SAP Records          : "
            f"{len(duplicate_sap_ids)}\n"
        )

        f.write(
            f"Invalid SAP Document Values    : "
            f"{len(invalid_sap_ids)}\n\n"
        )

        # ------------------------------------------------------------------
        # PDF DATA
        # ------------------------------------------------------------------

        f.write(
            "PDF DATA\n"
            + "-" * 78
            + "\n"
        )

        f.write(
            f"Unique PDF Documents           : "
            f"{pdf_count}\n"
        )

        f.write(
            f"PDF Document IDs with Duplicates: "
            f"{duplicate_count}\n"
        )

        f.write(
            f"PDF Files Without Document ID  : "
            f"{invalid_pdf_count}\n\n"
        )

        # ------------------------------------------------------------------
        # RECONCILIATION
        # ------------------------------------------------------------------

        f.write(
            "RECONCILIATION RESULTS\n"
            + "-" * 78
            + "\n"
        )

        f.write(
            f"Matched — Clean                : "
            f"{matched}\n"
        )

        f.write(
            f"Matched — Duplicate PDF        : "
            f"{matched_duplicate}\n"
        )

        f.write(
            f"Missing PDFs                    : "
            f"{missing_pdf}\n"
        )

        f.write(
            f"Untracked PDFs                  : "
            f"{untracked_pdf}\n"
        )

        f.write(
            f"Untracked Duplicate PDFs       : "
            f"{untracked_duplicate}\n"
        )

        f.write(
            f"Operational Match Rate         : "
            f"{match_rate:.2f}%\n\n"
        )

        # ------------------------------------------------------------------
        # SEQUENCE GAPS
        # ------------------------------------------------------------------

        f.write(
            "SAP DOCUMENT SEQUENCE GAPS\n"
            + "-" * 78
            + "\n"
        )

        if sap_gaps:

            for gap in sap_gaps:

                f.write(
                    f"Between "
                    f"{gap['gap_after_id']} "
                    f"and "
                    f"{gap['gap_before_id']} "
                    f"| Missing: "
                    f"{gap['missing_count']} "
                    f"| IDs: "
                    f"{gap['skipped_ids']}\n"
                )

        else:

            f.write(
                "No SAP sequence gaps detected.\n"
            )

        f.write("\n")

        f.write(
            "PDF DOCUMENT SEQUENCE GAPS\n"
            + "-" * 78
            + "\n"
        )

        if pdf_gaps:

            for gap in pdf_gaps:

                f.write(
                    f"Between "
                    f"{gap['gap_after_id']} "
                    f"and "
                    f"{gap['gap_before_id']} "
                    f"| Missing: "
                    f"{gap['missing_count']} "
                    f"| IDs: "
                    f"{gap['skipped_ids']}\n"
                )

        else:

            f.write(
                "No PDF sequence gaps detected.\n"
            )

        f.write("\n")

        # ------------------------------------------------------------------
        # DISCREPANCIES
        # ------------------------------------------------------------------

        f.write(
            "DISCREPANCIES REQUIRING ATTENTION\n"
            + "-" * 78
            + "\n"
        )

        discrepancies = [
            r
            for r in results
            if r["Status"] != "MATCHED"
        ]

        if discrepancies:

            for item in discrepancies:

                f.write(
                    f"ID: "
                    f"{item['Document_ID']} "
                    f"| Status: "
                    f"{item['Status']}\n"
                )

        else:

            f.write(
                "No discrepancies detected.\n"
            )

        f.write("\n")

        # ------------------------------------------------------------------
        # END
        # ------------------------------------------------------------------

        f.write(
            "=" * 78
            + "\n"
        )

        f.write(
            "END OF REPORT\n"
        )

        f.write(
            "=" * 78
            + "\n"
        )

    log(
        f"Summary report created: "
        f"{SUMMARY_REPORT}"
    )


# ==============================================================================
# FINAL CONSOLE SUMMARY
# ==============================================================================

def print_final_summary(
    results,
    sap_count,
    pdf_count,
    all_sap_rows,
    filtered_sap_rows
):

    matched = sum(
        1
        for r in results
        if r["Status"] == "MATCHED"
    )

    matched_duplicate = sum(
        1
        for r in results
        if r["Status"]
        == "MATCHED_DUPLICATE_PDF"
    )

    missing_pdf = sum(
        1
        for r in results
        if r["Status"] == "MISSING_PDF"
    )

    untracked_pdf = sum(
        1
        for r in results
        if r["Status"] == "UNTRACKED_PDF"
    )

    untracked_duplicate = sum(
        1
        for r in results
        if r["Status"]
        == "UNTRACKED_DUPLICATE_PDF"
    )

    accounted = (
        matched
        + matched_duplicate
    )

    match_rate = (
        accounted / sap_count * 100
        if sap_count
        else 0
    )

    print()
    print("=" * 78)
    print("KUBO RECONCILIATION RESULT")
    print("=" * 78)

    print(
        f"SAP rows loaded             : "
        f"{all_sap_rows}"
    )

    print(
        f"SAP rows after month filter : "
        f"{filtered_sap_rows}"
    )

    print(
        f"Unique SAP documents        : "
        f"{sap_count}"
    )

    print(
        f"Unique PDF documents        : "
        f"{pdf_count}"
    )

    print(
        f"Matched                     : "
        f"{matched}"
    )

    print(
        f"Matched with duplicate PDF  : "
        f"{matched_duplicate}"
    )

    print(
        f"Missing PDFs                : "
        f"{missing_pdf}"
    )

    print(
        f"Untracked PDFs              : "
        f"{untracked_pdf}"
    )

    print(
        f"Untracked duplicate PDFs    : "
        f"{untracked_duplicate}"
    )

    print(
        f"Operational match rate      : "
        f"{match_rate:.2f}%"
    )

    print("=" * 78)

    if missing_pdf > 0:

        print(
            f"[ACTION] {missing_pdf} SAP "
            f"document(s) have no acknowledged PDF."
        )

    if matched_duplicate > 0:

        print(
            f"[ACTION] {matched_duplicate} SAP "
            f"document(s) have duplicate PDFs."
        )

    if untracked_pdf > 0:

        print(
            f"[ACTION] {untracked_pdf} PDF(s) "
            f"were not found in SAP."
        )

    if (
        missing_pdf == 0
        and matched_duplicate == 0
        and untracked_pdf == 0
        and untracked_duplicate == 0
    ):

        print(
            "[OK] No reconciliation discrepancies detected."
        )

    print("=" * 78)


# ==============================================================================
# MAIN
# ==============================================================================

def main():

    print()

    print("=" * 78)
    print("KUBO DOCUMENT MATCHING ENGINE — V5.5")
    print("=" * 78)

    log(
        f"Year: {CURRENT_YEAR}"
    )

    log(
        f"Month: {CURRENT_MONTH}"
    )

    log(
        f"Base directory: {BASE_DIR}"
    )

    log(
        f"Delivery directory: {DELIVERIES_DIR}"
    )

    # ==========================================================================
    # STEP 1 — ENVIRONMENT
    # ==========================================================================

    log(
        "\n[1/4] Validating network environment..."
    )

    if not validate_environment():

        log(
            "[FATAL] Environment validation failed."
        )

        return 1

    # ==========================================================================
    # STEP 2 — SAP
    # ==========================================================================

    log(
        "\n[2/4] Extracting SAP document records..."
    )

    (
        sap_ids,
        invalid_sap_ids,
        sap_source,
        all_sap_rows,
        filtered_sap_rows,
        duplicate_sap_ids,
        filtering_method
    ) = get_sap_document_ids()

    if not sap_ids:

        log(
            f"[FATAL] No valid SAP document IDs "
            f"found for "
            f"{CURRENT_MONTH} "
            f"{CURRENT_YEAR}."
        )

        return 1

    log(
        f"[OK] {len(sap_ids)} unique SAP "
        f"documents loaded."
    )

    # ==========================================================================
    # STEP 3 — PDF DISCOVERY
    # ==========================================================================

    log(
        "\n[3/4] Scanning acknowledged delivery PDFs..."
    )

    (
        pdf_dict,
        duplicate_dict,
        invalid_pdf_files
    ) = find_delivery_pdfs()

    # ==========================================================================
    # STEP 4 — RECONCILIATION
    # ==========================================================================

    log(
        "\n[4/4] Executing reconciliation..."
    )

    results = perform_matching(
        sap_ids,
        pdf_dict,
        duplicate_dict
    )

    # ==========================================================================
    # REPORTING
    # ==========================================================================

    write_csv_report(
        results
    )

    write_summary_report(
        results=results,

        sap_count=len(sap_ids),

        pdf_count=len(pdf_dict),

        duplicate_count=len(
            duplicate_dict
        ),

        invalid_pdf_count=len(
            invalid_pdf_files
        ),

        sap_source=sap_source,

        invalid_sap_ids=invalid_sap_ids,

        all_sap_rows=all_sap_rows,

        filtered_sap_rows=filtered_sap_rows,

        duplicate_sap_ids=duplicate_sap_ids,

        filtering_method=filtering_method
    )

    # ==========================================================================
    # FINAL RESULT
    # ==========================================================================

    print_final_summary(
        results=results,

        sap_count=len(sap_ids),

        pdf_count=len(pdf_dict),

        all_sap_rows=all_sap_rows,

        filtered_sap_rows=filtered_sap_rows
    )

    log(
        "[OK] KUBO MATCHING PROCESS COMPLETE"
    )

    return 0


# ==============================================================================
# EXECUTION
# ==============================================================================

if __name__ == "__main__":

    sys.exit(
        main()
    )