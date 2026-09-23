#!/usr/bin/env python3
"""
pricing_uploader.py

Python replacement for the "Load_Files_To_Staging" VBA macro in
Pricing_uploader_multiFile.xlsm.

WHAT IT DOES:

  For every line in files.csv (Business Date, Pricing SMP Matrix File,
  SMP Matrix File, DC Prices File):
    1. Build the day's folder from the date:
         \\\\eh-fp-01.energia.local\\trading\\Day Folders\\YYYYMM\\DD\\
    2. Read the "SMP Matrix (Euro) INPUT" sheet out of the SMP Matrix file
       and the Pricing SMP Matrix file, and the "Euros" sheet out of the
       DC Prices file.
    3. DELETE FROM the three staging.* tables in SQL Server.
    4. Bulk-insert the three files into the staging tables (see UPLOAD_PLAN).
    5. EXEC dbo.Process_Pricing_Data.

  A failure on one day is logged and skipped; the run carries on with the
  next day, like the macro's On Error Resume Next per-row behaviour.

WHY THIS IS FASTER THAN THE MACRO:
  - The macro inserts one row at a time with cn.Execute "INSERT ... VALUES (...)".
    For a ~1900-row matrix file that's ~1900 separate network round trips to
    SQL Server, per file, per day. This script uses pyodbc's
    fast_executemany, which batches the same inserts into a handful of
    round trips.
  - Source workbooks are opened with openpyxl in read_only mode, which is
    much faster than a full Excel Workbooks.Open and never launches Excel.

WHAT IT DELIBERATELY PRESERVES FROM THE MACRO (do not "fix" without asking):
  - The crossed table mapping: the "SMP Matrix" file goes into
    staging.Pricing_SMP_Matrix_LBE and the "Pricing SMP Matrix" file goes
    into staging.SMP_Matrix_LBE. That's what the macro did on every row of
    the control sheet. If it's a long-standing bug, flip UPLOAD_PLAN.
  - The "LastRow" sentinel logic for matrix files (stop at the row where
    column A == 0, else fall back to the last non-blank cell in column A)
    and the fixed "AX" last-column boundary for matrix files.
  - The DC Prices file's fixed layout: data lives in B5:F<lastrow of col C>.

REQUIREMENTS (install on the work PC):
    pip install -r requirements.txt
    -> also needs "ODBC Driver 18 for SQL Server". Windows authentication
       is used, same as the macro's "Integrated Security=SSPI".

USAGE:
    python pricing_uploader.py                Upload every day in files.csv
    python pricing_uploader.py --dry-run      Read the files, touch no SQL
    python pricing_uploader.py --test-connection

    Run "python pricing_uploader.py --help" for the full list.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import datetime
import getpass
import logging
import os
import sys
from dataclasses import dataclass, field
from typing import Any, Optional

import openpyxl
from openpyxl.utils import column_index_from_string

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# SQL Server connection details (same server/database the macro used).
SQL_SERVER = "ENR-HVH-EGL-02"
SQL_DATABASE = "FODB"
ODBC_DRIVER = "ODBC Driver 18 for SQL Server"

# Each day's files live in <DAY_FOLDERS_ROOT>\YYYYMM\DD\ (the same folder
# the control sheet's column D formula built).
DAY_FOLDERS_ROOT = r"\\eh-fp-01.energia.local\trading\Day Folders"

# The list of days to upload, one line per day. Looked for next to this
# script unless --csv says otherwise.
DEFAULT_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files.csv")
CSV_DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")

# Staging tables cleared before every day's upload.
STAGING_TABLES_TO_CLEAR = [
    "staging.SMP_Matrix_LBE",
    "staging.Pricing_SMP_Matrix_LBE",
    "staging.DC_Prices",
]

STORED_PROCEDURE = "dbo.Process_Pricing_Data"

# Source sheet names inside the day-folder workbooks.
MATRIX_SOURCE_SHEET = "SMP Matrix (Euro) INPUT"
DC_PRICES_SOURCE_SHEET = "Euros"

# The macro's exact (crossed) mapping of source file -> staging table.
# See the docstring above before changing it.
UPLOAD_PLAN = [
    # (in-memory table name,     SQL table it is inserted into)
    ("SMP_Matrix_LBE", "staging.Pricing_SMP_Matrix_LBE"),
    ("Pricing_SMP_Matrix_LBE", "staging.SMP_Matrix_LBE"),
    ("DC_Prices", "staging.DC_Prices"),
]

AX_COLUMN_INDEX = column_index_from_string("AX")  # 50 -- matrix file last column

EXCEL_ERROR_STRINGS = {
    "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NULL!", "#NAME?", "#GETTING_DATA",
}


# --------------------------------------------------------------------------
# Data holders
# --------------------------------------------------------------------------

@dataclass
class LoadedTable:
    name: str
    rows: list[list[Any]] = field(default_factory=list)
    ncols: int = 0
    source_file: str = ""


@dataclass
class DayResult:
    business_date: datetime.date
    ok: bool
    message: str = ""


# --------------------------------------------------------------------------
# Excel reading helpers
# --------------------------------------------------------------------------

def win_basename(path: str) -> str:
    """Extract just the filename off the end of a path, regardless of which
    slash style it uses and regardless of which OS this script runs on.
    (os.path.basename is platform-dependent -- on Linux/Mac it does NOT
    split on backslashes, which would silently break parsing of Windows
    UNC paths like \\\\server\\share\\file.xlsx if this were ever run,
    tested, or debugged off Windows.)"""
    return path.replace("\\", "/").rsplit("/", 1)[-1]


def clean_cell_value(value: Any) -> Any:
    """Map an openpyxl cell value the same way the macro's IsError/IsEmpty
    checks did: Excel errors and empty strings both become NULL. Everything
    else is passed straight through as a native Python type (int/float/
    str/datetime), which pyodbc will bind correctly."""
    if value is None:
        return None
    if isinstance(value, str):
        if value == "" or value in EXCEL_ERROR_STRINGS:
            return None
        return value
    return value


def find_matrix_last_row(ws) -> int:
    """Replicates:
        For r = 2 To wsSource.Rows.Count
            If wsSource.Cells(r, 1).Value = 0 Then LastRow = r - 1: Exit For
        Next r
        If LastRow = 0 Then LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    """
    last_row = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v == 0:
            last_row = r - 1
            break
    if last_row == 0:
        # End(xlUp) from the bottom of column A
        for r in range(ws.max_row, 1, -1):
            v = ws.cell(row=r, column=1).value
            if v not in (None, ""):
                last_row = r
                break
        if last_row == 0:
            last_row = 1
    return last_row


def find_last_nonblank_row(ws, col_index: int) -> int:
    """Replicates Cells(Rows.Count, col).End(xlUp).Row for a given column."""
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(row=r, column=col_index).value
        if v not in (None, ""):
            return r
    return 0


def load_matrix_file(full_path: str, target_name: str, uploaded_by: str,
                      upload_date: datetime.datetime) -> LoadedTable:
    """Replicates LoadMatrixFile."""
    filename = win_basename(full_path)

    wb = openpyxl.load_workbook(full_path, read_only=True, data_only=True)
    try:
        ws = wb[MATRIX_SOURCE_SHEET]

        last_row = find_matrix_last_row(ws)
        last_col = AX_COLUMN_INDEX
        data_rows = last_row - 1  # rows 2..last_row inclusive

        rows: list[list[Any]] = []
        if data_rows > 0:
            for r in range(2, last_row + 1):
                row_vals = [clean_cell_value(ws.cell(row=r, column=c).value)
                            for c in range(1, last_col + 1)]
                row_vals.extend([uploaded_by, upload_date, filename])
                rows.append(row_vals)

        return LoadedTable(name=target_name, rows=rows, ncols=last_col + 3,
                            source_file=filename)
    finally:
        wb.close()


def load_dc_prices_file(full_path: str, target_name: str, uploaded_by: str,
                         upload_date: datetime.datetime) -> LoadedTable:
    """Replicates LoadDCPricesFile: data lives in B5:F<lastrow-of-col-C>."""
    filename = win_basename(full_path)

    wb = openpyxl.load_workbook(full_path, read_only=True, data_only=True)
    try:
        ws = wb[DC_PRICES_SOURCE_SHEET]

        last_row = find_last_nonblank_row(ws, column_index_from_string("C"))
        data_rows = last_row - 4  # data starts row 5

        rows: list[list[Any]] = []
        if data_rows > 0:
            col_b = column_index_from_string("B")
            col_f = column_index_from_string("F")
            for r in range(5, last_row + 1):
                row_vals = [clean_cell_value(ws.cell(row=r, column=c).value)
                            for c in range(col_b, col_f + 1)]
                row_vals.extend([uploaded_by, upload_date, filename])
                rows.append(row_vals)

        return LoadedTable(name=target_name, rows=rows, ncols=8, source_file=filename)
    finally:
        wb.close()


# --------------------------------------------------------------------------
# SQL Server helpers
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# SQL Server helpers
# --------------------------------------------------------------------------

def get_connection(sql_server: str = SQL_SERVER, sql_database: str = SQL_DATABASE):
    import pyodbc  # imported lazily so --dry-run works without pyodbc installed

    # Driver 18 encrypts by default and rejects the self-signed certificates
    # internal SQL Servers usually have. TrustServerCertificate keeps the
    # connection encrypted but skips the certificate check, which is what
    # Driver 17 and the macro's SQLOLEDB connection were already doing.
    conn_str = (
        f"DRIVER={{{ODBC_DRIVER}}};"
        f"SERVER={sql_server};"
        f"DATABASE={sql_database};"
        f"Trusted_Connection=yes;"
        f"Encrypt=yes;"
        f"TrustServerCertificate=yes;"
    )
    conn = pyodbc.connect(conn_str, autocommit=False)
    return conn


def clear_staging_tables(conn) -> None:
    cur = conn.cursor()
    for table in STAGING_TABLES_TO_CLEAR:
        cur.execute(f"DELETE FROM {table}")
    conn.commit()


def upload_table(conn, table_name: str, table: LoadedTable, batch_size: int = 1000) -> int:
    """Bulk insert with fast_executemany. Returns rows inserted.

    fast_executemany infers each column's SQL type from the first row of a
    batch. If that first row happens to have NULL in a column that later
    rows fill with text/numbers, pyodbc can occasionally raise a binding
    error ("String data, right truncation" being the classic one). That's
    a known pyodbc quirk, not a sign the data is wrong, so on that specific
    failure we retry the same chunk once with fast_executemany switched off
    (slower, but always correct) instead of aborting the whole row."""
    import pyodbc

    if not table.rows:
        return 0

    ncols = len(table.rows[0])
    placeholders = ",".join(["?"] * ncols)
    sql = f"INSERT INTO {table_name} VALUES ({placeholders})"

    total = 0
    for start in range(0, len(table.rows), batch_size):
        chunk = table.rows[start:start + batch_size]

        cur = conn.cursor()
        try:
            cur.fast_executemany = True
        except AttributeError:
            pass

        try:
            cur.executemany(sql, chunk)
        except pyodbc.Error:
            cur2 = conn.cursor()
            cur2.fast_executemany = False
            cur2.executemany(sql, chunk)

        total += len(chunk)
    conn.commit()
    return total


def run_stored_procedure(conn) -> None:
    cur = conn.cursor()
    cur.execute(f"EXEC {STORED_PROCEDURE}")
    conn.commit()


# --------------------------------------------------------------------------
# files.csv reading
# --------------------------------------------------------------------------

@dataclass
class DayFiles:
    line: int
    business_date: datetime.date
    pricing_file: str
    matrix_file: str
    dc_file: str

    @property
    def folder(self) -> str:
        return day_folder(DAY_FOLDERS_ROOT, self.business_date)


def day_folder(root: str, business_date: datetime.date) -> str:
    return os.path.join(root, business_date.strftime("%Y%m"), business_date.strftime("%d"))


def parse_date(text: str) -> Optional[datetime.date]:
    text = text.strip()
    for fmt in CSV_DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def read_csv_lines(csv_path: str) -> list[list[str]]:
    """Excel's plain "CSV (Comma delimited)" saves as Windows-1252, while
    "CSV UTF-8" adds a byte-order mark; handle both."""
    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            return list(csv.reader(f))
    except UnicodeDecodeError:
        with open(csv_path, newline="", encoding="cp1252") as f:
            return list(csv.reader(f))


def read_days(csv_path: str) -> list[DayFiles]:
    """Reads files.csv. Raises ValueError listing every bad line, so a typo
    stops the run before anything is uploaded."""
    days: list[DayFiles] = []
    errors: list[str] = []

    for line_no, cells in enumerate(read_csv_lines(csv_path), start=1):
        cells = [c.strip() for c in cells]
        if not any(cells):
            continue

        business_date = parse_date(cells[0])
        if business_date is None:
            if line_no == 1:
                continue  # header row
            errors.append(f"line {line_no}: {cells[0]!r} is not a date (use DD/MM/YYYY)")
            continue

        files = (cells + ["", "", ""])[1:4]
        missing = [name for name, value in zip(
            ("Pricing SMP Matrix File", "SMP Matrix File", "DC Prices File"), files) if not value]
        if missing:
            errors.append(f"line {line_no} ({cells[0]}): missing {', '.join(missing)}")
            continue

        days.append(DayFiles(line_no, business_date, *files))

    if errors:
        raise ValueError("Problems in " + csv_path + ":\n  " + "\n  ".join(errors))
    return days


# --------------------------------------------------------------------------
# Per-day processing
# --------------------------------------------------------------------------

def load_day_files(day: DayFiles, uploaded_by: str, upload_date: datetime.datetime) -> dict[str, LoadedTable]:
    """Reads the 3 source workbooks for one day. Pure I/O + parsing, no
    SQL -- safe to run in a worker thread."""
    return {
        "SMP_Matrix_LBE": load_matrix_file(
            os.path.join(day.folder, day.matrix_file),
            "SMP_Matrix_LBE", uploaded_by, upload_date),
        "Pricing_SMP_Matrix_LBE": load_matrix_file(
            os.path.join(day.folder, day.pricing_file),
            "Pricing_SMP_Matrix_LBE", uploaded_by, upload_date),
        "DC_Prices": load_dc_prices_file(
            os.path.join(day.folder, day.dc_file),
            "DC_Prices", uploaded_by, upload_date),
    }


def process_day(conn, day: DayFiles, tables: dict[str, LoadedTable],
                dry_run: bool, log: logging.Logger) -> DayResult:
    try:
        if dry_run:
            for name, dest in UPLOAD_PLAN:
                t = tables[name]
                log.info(f"  [dry-run] {len(t.rows)} rows from {t.source_file!r} -> {dest}")
            log.info(f"  [dry-run] would EXEC {STORED_PROCEDURE}")
            return DayResult(day.business_date, True, "dry-run")

        clear_staging_tables(conn)

        for name, dest in UPLOAD_PLAN:
            n = upload_table(conn, dest, tables[name])
            log.info(f"  inserted {n} rows into {dest} (from {tables[name].source_file!r})")

        run_stored_procedure(conn)

        return DayResult(day.business_date, True)
    except Exception as exc:  # mirrors the macro's On Error Resume Next per-row
        try:
            conn.rollback()
        except Exception:
            pass
        return DayResult(day.business_date, False, str(exc))


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def date_arg(text: str) -> datetime.date:
    d = parse_date(text)
    if d is None:
        raise argparse.ArgumentTypeError(f"{text!r} is not a date (use DD/MM/YYYY)")
    return d


def main() -> int:
    global DAY_FOLDERS_ROOT

    parser = argparse.ArgumentParser(description="Python replacement for Load_Files_To_Staging macro")
    parser.add_argument("--csv", default=DEFAULT_CSV,
                        help="List of days and file names (default: files.csv next to this script)")
    parser.add_argument("--dry-run", action="store_true", help="Read the files and print what would happen; no SQL")
    parser.add_argument("--date", type=date_arg, help="Only upload this one day (DD/MM/YYYY)")
    parser.add_argument("--from", dest="date_from", type=date_arg, help="Only upload days on or after this date")
    parser.add_argument("--to", dest="date_to", type=date_arg, help="Only upload days on or before this date")
    parser.add_argument("--workers", type=int, default=1, help="Parallel source-file readers (I/O bound; default 1)")
    parser.add_argument("--root", default=DAY_FOLDERS_ROOT, help=f"Day folders location (default {DAY_FOLDERS_ROOT})")
    parser.add_argument("--server", default=SQL_SERVER, help=f"SQL Server host (default {SQL_SERVER})")
    parser.add_argument("--database", default=SQL_DATABASE, help=f"SQL Server database (default {SQL_DATABASE})")
    parser.add_argument("--log-file", help="Also write log output to this file")
    parser.add_argument("--uploaded-by", default=None, help="Overrides Environ('USERNAME'); defaults to the current OS user")
    parser.add_argument("--test-connection", action="store_true",
                        help="Just open and close a connection to SQL Server, then exit")
    args = parser.parse_args()
    DAY_FOLDERS_ROOT = args.root

    log = logging.getLogger("pricing_uploader")
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%H:%M:%S")
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    log.addHandler(sh)
    if args.log_file:
        fh = logging.FileHandler(args.log_file)
        fh.setFormatter(fmt)
        log.addHandler(fh)

    def connect():
        try:
            return get_connection(args.server, args.database)
        except ImportError:
            log.error("pyodbc is not installed. Run: pip install -r requirements.txt")
        except Exception as exc:
            log.error(f"Could not connect to {args.server}/{args.database}: {exc}")
            if "IM002" in str(exc):
                log.error(f"'{ODBC_DRIVER}' is not installed. Install it from Microsoft "
                          "(search 'ODBC Driver 18 for SQL Server download') or ask IT.")
        return None

    if args.test_connection:
        log.info(f"Connecting to {args.server}/{args.database} using {ODBC_DRIVER} ...")
        conn = connect()
        if conn is None:
            return 1
        conn.close()
        log.info("Connected successfully.")
        return 0

    if not os.path.exists(args.csv):
        log.error(f"Can't find {args.csv}. Put files.csv next to the script, or pass --csv \"path\\to\\file.csv\".")
        return 1

    log.info(f"Reading days from {args.csv}")
    try:
        days = read_days(args.csv)
    except ValueError as exc:
        log.error(str(exc))
        log.error("Nothing was uploaded. Fix the lines above and run again.")
        return 1

    if args.date:
        days = [d for d in days if d.business_date == args.date]
    if args.date_from:
        days = [d for d in days if d.business_date >= args.date_from]
    if args.date_to:
        days = [d for d in days if d.business_date <= args.date_to]
    log.info(f"{len(days)} day(s) to process")

    if not days:
        log.warning("Nothing to do.")
        return 0

    conn = None
    if not args.dry_run:
        log.info(f"Connecting to {args.server}/{args.database} ...")
        conn = connect()
        if conn is None:
            log.error("Try: python pricing_uploader.py --test-connection")
            return 1

    uploaded_by = args.uploaded_by or os.environ.get("USERNAME") or getpass.getuser()
    results: list[DayResult] = []

    def do_load(day: DayFiles):
        return load_day_files(day, uploaded_by, datetime.datetime.now())

    def handle(day: DayFiles, loaded):
        log.info(f"{day.business_date:%d/%m/%Y} ({day.folder}):")
        if isinstance(loaded, Exception):
            log.error(f"  FAILED loading source files - {loaded}")
            results.append(DayResult(day.business_date, False, str(loaded)))
            return
        results.append(process_day(conn, day, loaded, args.dry_run, log))

    if args.workers > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = [pool.submit(do_load, day) for day in days]
            # Process in file order so SQL writes stay sequential/predictable.
            for day, fut in zip(days, futures):
                try:
                    loaded = fut.result()
                except Exception as exc:
                    loaded = exc
                handle(day, loaded)
    else:
        for day in days:
            try:
                loaded = do_load(day)
            except Exception as exc:
                loaded = exc
            handle(day, loaded)

    if conn is not None:
        conn.close()

    ok = sum(1 for r in results if r.ok)
    failed = [r for r in results if not r.ok]
    log.info(f"Done. {ok}/{len(results)} day(s) succeeded.")
    if failed:
        log.warning(f"{len(failed)} day(s) failed:")
        for r in failed:
            log.warning(f"  {r.business_date:%d/%m/%Y}: {r.message}")

    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
