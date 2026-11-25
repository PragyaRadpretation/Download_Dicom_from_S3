"""
Download DICOM study archives from Orthanc based on study IDs listed in an Excel file.

Output file structure:
    <ROOT_FOLDER>/<studyID>.zip

Behavior:
    - Reads study IDs from the Excel file (default: DX_studies.xlsx).
    - Downloads each study via GET <ORTHANC_URL>/studies/<studyID>/archive.
    - Saves the archive directly to <ROOT_FOLDER>/<studyID>.zip.
    - On success: deletes that row from the Excel file and saves immediately.
    - On error: logs the error in the terminal and in an 'Error' column on that row
      in the Excel file without deleting the row, then proceeds to the next row.
"""

from datetime import datetime
from pathlib import Path
from urllib.parse import quote
import os
import re
import sys
import zipfile

import openpyxl
import requests


# =============================================================================
# CONFIGURATION - Update these two values before running
# =============================================================================
ORTHANC_URL = "https://pacs-ayurveda.radpretation.ai/"          # Orthanc base URL (e.g., "http://192.168.1.10:8042")
ROOT_FOLDER = r"dicom_to_nifti/ayurveda_ct_mapped_dicom"        # Root folder where <studyID>.zip files will be saved

# Optional Orthanc Basic Auth (leave empty if authentication is not enabled)
ORTHANC_USERNAME = ""
ORTHANC_PASSWORD = ""

# Excel file settings (defaults to DX_studies.xlsx in the same directory as this script)
EXCEL_FILE = Path(__file__).resolve().parent / "CT_studies_dicom_download.xlsx"
SHEET_NAME = None  # None uses the active sheet (e.g., 'in')

# Network / download settings
CONNECT_TIMEOUT = 30          # Seconds to establish connection to Orthanc
READ_TIMEOUT = 600            # Seconds to wait between chunks for large studies
CHUNK_SIZE = 1024 * 1024      # 1 MB chunks
VERIFY_SSL = True             # Set to False if Orthanc uses a self-signed HTTPS certificate
# =============================================================================


INVALID_PATH_CHARS = re.compile(r'[<>:"/\\|?*]')


class ExcelSaveError(Exception):
    """Raised when the Excel workbook cannot be saved (e.g., file is open/locked)."""


def save_workbook_safely(wb: openpyxl.Workbook, excel_path: Path) -> None:
    """Save the Excel workbook atomically to prevent corruption on interruption."""
    temp_path = excel_path.with_name(f".{excel_path.stem}_tmp{excel_path.suffix}")
    try:
        wb.save(temp_path)
        os.replace(temp_path, excel_path)
    except PermissionError as exc:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)
        raise ExcelSaveError(
            f"Cannot save '{excel_path}'. Please close the file if it is open in Excel."
        ) from exc
    except Exception as exc:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)
        raise ExcelSaveError(f"Failed to save '{excel_path}': {exc}") from exc


def find_column_index(ws, target_name: str):
    """Find a column index (1-based) by case-insensitive header name in row 1."""
    target_lower = target_name.strip().lower()
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col_idx).value
        if val is not None and str(val).strip().lower() == target_lower:
            return col_idx
    return None


def get_or_create_error_column(ws) -> int:
    """Return the 'Error' column index, creating the header if it does not exist."""
    error_col = find_column_index(ws, "error")
    if error_col is None:
        error_col = ws.max_column + 1
        ws.cell(row=1, column=error_col, value="Error")
    return error_col


def download_study_archive(
    session: requests.Session,
    orthanc_url: str,
    study_id: str,
    root_folder: Path,
) -> Path:
    """
    Download a single study archive from Orthanc and save it directly to:
        <root_folder>/<study_id>.zip
    """
    if INVALID_PATH_CHARS.search(study_id):
        raise ValueError(f"Study ID contains invalid filesystem characters: '{study_id}'")

    root_folder.mkdir(parents=True, exist_ok=True)

    final_zip_path = root_folder / f"{study_id}.zip"
    temp_zip_path = root_folder / f"{study_id}.zip.part"

    url = f"{orthanc_url.rstrip('/')}/studies/{quote(study_id, safe='')}/archive"

    try:
        with session.get(
            url,
            stream=True,
            timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            verify=VERIFY_SSL,
        ) as response:
            response.raise_for_status()

            bytes_written = 0
            with open(temp_zip_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                    if chunk:
                        f.write(chunk)
                        bytes_written += len(chunk)

        if bytes_written == 0:
            raise ValueError("Downloaded archive is empty (0 bytes).")

        if not zipfile.is_zipfile(temp_zip_path):
            raise ValueError("Downloaded file is not a valid ZIP archive.")

        os.replace(temp_zip_path, final_zip_path)
        return final_zip_path

    except BaseException:
        # Clean up incomplete download
        if temp_zip_path.exists():
            temp_zip_path.unlink(missing_ok=True)
        raise


def main() -> int:
    orthanc_url = ORTHANC_URL.strip()
    root_folder_str = ROOT_FOLDER.strip()

    if not orthanc_url or not orthanc_url.startswith(("http://", "https://")):
        print(f"[ERROR] Invalid ORTHANC_URL: '{orthanc_url}'. Must start with http:// or https://")
        return 1

    if not root_folder_str or root_folder_str == r"C:\path\to\root_folder":
        print("[ERROR] Please set ROOT_FOLDER at the top of download_dicom.py before running.")
        return 1

    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        print(f"[ERROR] Excel file not found: {excel_path}")
        return 1

    root_folder = Path(root_folder_str)
    try:
        root_folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(f"[ERROR] Could not create or access ROOT_FOLDER '{root_folder}': {exc}")
        return 1

    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as exc:
        print(f"[ERROR] Failed to open Excel file '{excel_path}': {exc}")
        return 1

    ws = wb[SHEET_NAME] if SHEET_NAME else wb.active
    study_col = find_column_index(ws, "studyID")
    if study_col is None:
        study_col = 1

    total_rows = max(0, ws.max_row - 1)
    if total_rows == 0:
        print("No study IDs found in the Excel file. Nothing to download.")
        return 0

    print(f"Excel file   : {excel_path} (Sheet: '{ws.title}')")
    print(f"Orthanc URL  : {orthanc_url}")
    print(f"Root folder  : {root_folder}")
    print(f"Studies found: {total_rows}")
    print("-" * 70)

    session = requests.Session()
    if ORTHANC_USERNAME:
        session.auth = (ORTHANC_USERNAME, ORTHANC_PASSWORD)

    succeeded = 0
    failed = 0
    skipped = 0
    processed = 0
    row_idx = 2  # Row 1 is the header

    try:
        while row_idx <= ws.max_row:
            raw_val = ws.cell(row=row_idx, column=study_col).value
            if raw_val is None or not str(raw_val).strip():
                skipped += 1
                row_idx += 1
                continue

            study_id = str(raw_val).strip()
            processed += 1
            print(f"[{processed}/{total_rows}] Downloading study: {study_id} ...")

            try:
                saved_path = download_study_archive(session, orthanc_url, study_id, root_folder)
                size_mb = saved_path.stat().st_size / (1024 * 1024)

                # Delete the processed row and immediately persist Excel
                ws.delete_rows(row_idx, 1)
                save_workbook_safely(wb, excel_path)

                succeeded += 1
                print(f"    [OK] Saved ({size_mb:.2f} MB) -> {saved_path} (row deleted from Excel)")
                # Do NOT increment row_idx because the next row shifted up into row_idx

            except KeyboardInterrupt:
                raise
            except ExcelSaveError as exc:
                # Excel file is locked (e.g., open in Microsoft Excel); stop so status isn't lost
                print(f"    [FATAL] {exc}")
                return 1
            except Exception as exc:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                error_msg = f"[{timestamp}] {type(exc).__name__}: {exc}"
                print(f"    [FAILED] {study_id} -> {error_msg}")

                # Log the error in Excel on the same row and keep the row unprocessed
                error_col = get_or_create_error_column(ws)
                ws.cell(row=row_idx, column=error_col, value=error_msg)
                try:
                    save_workbook_safely(wb, excel_path)
                except ExcelSaveError as save_exc:
                    print(f"    [FATAL] {save_exc}")
                    return 1

                failed += 1
                row_idx += 1

    except KeyboardInterrupt:
        print("\n[INTERRUPTED] Script stopped by user. Progress up to the last completed study has been saved.")
    finally:
        session.close()
        wb.close()

    print("-" * 70)
    print(f"Summary: Succeeded (deleted) = {succeeded} | Failed (kept) = {failed} | Empty skipped = {skipped}")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
