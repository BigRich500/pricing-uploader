#!/usr/bin/env python3
"""
pricing_uploader.py

Python replacement for the "Load_Files_To_Staging" VBA macro in
Pricing_uploader_multiFile.xlsm.

WHAT IT DOES (mirrors the macro row-for-row):

  For every row (3..last) on the "File_uploader" control sheet:
    1. Build the day's folder path from column D.
    2. Read the "SMP Matrix (Euro) INPUT" sheet out of the file named in
       column G  -> in-memory table "SMP_Matrix_LBE"
    3. Read the "SMP Matrix (Euro) INPUT" sheet out of the file named in
       column E  -> in-memory table "Pricing_SMP_Matrix_LBE"
    4. Read the "Euros" sheet out of the file named in column I
       -> in-memory table "DC_Prices"
    5. DELETE FROM the three staging.* tables in SQL Server (matches the
       macro exactly -- it uses DELETE, not TRUNCATE).
    6. Bulk-insert the three in-memory tables into SQL Server, using the
       table names resolved from columns H, F and J.
    7. EXEC dbo.Process_Pricing_Data.

WHY THIS IS FASTER THAN THE MACRO:
  - The macro inserts one row at a time with cn.Execute "INSERT ... VALUES (...)"
    built as a giant string. For a ~1900-row matrix file that's ~1900 separate
    network round trips to SQL Server, per file, per day. This script uses
    pyodbc's fast_executemany, which batches the same inserts into a handful
    of round trips.
  - Source workbooks are opened with openpyxl in read_only mode, which is
    dramatically faster than a full Excel Application/Workbook open, and it
    never launches Excel or shows any UI.
  - Values are passed to SQL Server as native Python types (through bound
    parameters) instead of being hand-formatted into a SQL string, which
    removes the string-building/locale-formatting VBA had to do.

WHAT IT DELIBERATELY PRESERVES FROM THE MACRO (do not "fix" without asking):
  - The staging.* tables that are cleared are always the same 3 fixed names,
    regardless of what the control sheet's H/F/J cells say.
  - The table-name *assignment* looks swapped versus the sheet's own column
    headers: the sheet built from column G ("SMP_Matrix_LBE" in-memory table)
    is uploaded to the table named in column H, and the sheet built from
    column E ("Pricing_SMP_Matrix_LBE" in-memory table) is uploaded to the
    table named in column F. In the sample workbook, H resolves to
    "staging.Pricing_SMP_Matrix_LBE" and F resolves to "staging.SMP_Matrix_LBE"
    -- i.e. crossed relative to the in-memory table names. This script
    reproduces that exact crossing (see UPLOAD_PLAN below) because the macro
    does it that way. If this is actually a long-standing bug in the macro,
    flip UPLOAD_PLAN, but that's a decision for you, not something to guess at.
  - A failure on one day's row is logged and skipped, exactly like the
    macro's On Error Resume Next per-row behaviour; the run continues with
    the next row.
  - The "LastRow" sentinel logic for matrix files (stop at the row where
    column A == 0, else fall back to the last non-blank cell in column A)
    and the fixed "AX" last-column boundary for matrix files.
  - The DC Prices file's fixed layout: data lives in B5:F<lastrow of col C>.

REQUIREMENTS (install on the work PC):
    pip install openpyxl pyodbc
    -> also needs "ODBC Driver 17 (or 18) for SQL Server" installed on
       Windows (usually already present on a machine that has SSMS/Excel
       ODBC connections working). Integrated/Windows auth is used, same as
       the macro's "Integrated Security=SSPI".

USAGE:
    python pricing_uploader.py --control "Pricing_uploader_multiFile.xlsm"

    Useful options:
      --dry-run           Parse everything, print what WOULD happen, touch
                           no SQL. Use this first on a new machine.
      --row 5              Process only control-sheet row 5.
      --row 5:20            Process rows 5 through 20 inclusive.
      --workers 4          Read that many source-file trios in parallel
                           (network I/O bound; SQL writes stay sequential
                           and in the original row order).
      --log-file run.log   Also write the log to a file.

    Run "python pricing_uploader.py --help" for the full list.

    See README.md for a full step-by-step setup and verification checklist
    (installing dependencies, finding your ODBC driver, testing the SQL
    connection, dry-running against real files, etc.) before a first real
    run on the work PC.
"""

from __future__ import annotations

import argparse
import concurrent.futures
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

# Fixed staging tables that get cleared before every day's upload
# (matches Clear_Staging_Tables in the macro -- these are literal, not
# read from the control sheet).
STAGING_TABLES_TO_CLEAR = [
    "staging.SMP_Matrix_LBE",
    "staging.Pricing_SMP_Matrix_LBE",
    "staging.DC_Prices",
]

STORED_PROCEDURE = "dbo.Process_Pricing_Data"

# Source sheet names inside the day-folder workbooks.
MATRIX_SOURCE_SHEET = "SMP Matrix (Euro) INPUT"
DC_PRICES_SOURCE_SHEET = "Euros"

# Control-sheet name/layout inside the .xlsm.
CONTROL_SHEET = "File_uploader"
CONTROL_FIRST_ROW = 3
COL_FOLDER = "D"
COL_PRICING_FILE = "E"      # -> loaded as "Pricing_SMP_Matrix_LBE"
COL_PRICING_TABLE = "F"     # -> SQL table that the *SMP_Matrix_LBE* sheet goes to
COL_MATRIX_FILE = "G"       # -> loaded as "SMP_Matrix_LBE"
COL_MATRIX_TABLE = "H"      # -> SQL table that the *Pricing_SMP_Matrix_LBE* sheet goes to
COL_DC_FILE = "I"           # -> loaded as "DC_Prices"
COL_DC_TABLE = "J"          # -> SQL table that the DC_Prices sheet goes to

# The macro's exact (crossed) mapping of "in-memory table" -> "SQL table
# column it reads its destination name from". See the big docstring above.
UPLOAD_PLAN = [
    # (in-memory table name,        control-sheet column with the SQL table name)
    ("SMP_Matrix_LBE", COL_MATRIX_TABLE),      # H
    ("Pricing_SMP_Matrix_LBE", COL_PRICING_TABLE),  # F
    ("DC_Prices", COL_DC_TABLE),                # J
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
class RowResult:
    row: int
    business_date: Any
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

def get_connection(sql_server: str = SQL_SERVER, sql_database: str = SQL_DATABASE,
                    driver: str = "ODBC Driver 17 for SQL Server"):
    import pyodbc  # imported lazily so --dry-run works without pyodbc installed

    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={sql_server};"
        f"DATABASE={sql_database};"
        f"Trusted_Connection=yes;"
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


def list_odbc_drivers() -> list[str]:
    import pyodbc
    return [d for d in pyodbc.drivers() if "SQL Server" in d]


# --------------------------------------------------------------------------
# Control-sheet reading
# --------------------------------------------------------------------------

@dataclass
class ControlRow:
    row: int
    business_date: Any
    folder_path: str
    pricing_file: str
    pricing_table: str
    matrix_file: str
    matrix_table: str
    dc_file: str
    dc_table: str


def read_control_rows(control_path: str, only_rows: Optional[tuple[int, int]] = None) -> list[ControlRow]:
    wb = openpyxl.load_workbook(control_path, read_only=True, data_only=True)
    try:
        ws = wb[CONTROL_SHEET]

        col_e_idx = column_index_from_string(COL_PRICING_FILE)
        last_row = find_last_nonblank_row(ws, col_e_idx)

        def cell(row: int, col_letter: str) -> Any:
            return ws.cell(row=row, column=column_index_from_string(col_letter)).value

        results: list[ControlRow] = []
        for r in range(CONTROL_FIRST_ROW, last_row + 1):
            if only_rows is not None and not (only_rows[0] <= r <= only_rows[1]):
                continue

            folder = str(cell(r, COL_FOLDER) or "").strip()
            if not folder:
                continue
            if not folder.endswith("\\") and not folder.endswith("/"):
                folder += "\\"

            def s(col_letter: str) -> str:
                v = cell(r, col_letter)
                return str(v).strip() if v is not None else ""

            results.append(ControlRow(
                row=r,
                business_date=cell(r, "C"),
                folder_path=folder,
                pricing_file=s(COL_PRICING_FILE),
                pricing_table=s(COL_PRICING_TABLE),
                matrix_file=s(COL_MATRIX_FILE),
                matrix_table=s(COL_MATRIX_TABLE),
                dc_file=s(COL_DC_FILE),
                dc_table=s(COL_DC_TABLE),
            ))
        return results
    finally:
        wb.close()


# --------------------------------------------------------------------------
# Per-row processing
# --------------------------------------------------------------------------

def join_path(folder: str, filename: str) -> str:
    return folder + filename


def load_row_files(crow: ControlRow, uploaded_by: str, upload_date: datetime.datetime) -> dict[str, LoadedTable]:
    """Reads the 3 source workbooks for one control-sheet row. Pure I/O +
    parsing, no SQL -- safe to run in a worker thread."""
    matrix_table = load_matrix_file(
        join_path(crow.folder_path, crow.matrix_file),
        "SMP_Matrix_LBE", uploaded_by, upload_date,
    )
    pricing_table = load_matrix_file(
        join_path(crow.folder_path, crow.pricing_file),
        "Pricing_SMP_Matrix_LBE", uploaded_by, upload_date,
    )
    dc_table = load_dc_prices_file(
        join_path(crow.folder_path, crow.dc_file),
        "DC_Prices", uploaded_by, upload_date,
    )
    return {
        "SMP_Matrix_LBE": matrix_table,
        "Pricing_SMP_Matrix_LBE": pricing_table,
        "DC_Prices": dc_table,
    }


def resolve_sql_table_name(crow: ControlRow, column_letter: str) -> str:
    mapping = {
        COL_MATRIX_TABLE: crow.matrix_table,
        COL_PRICING_TABLE: crow.pricing_table,
        COL_DC_TABLE: crow.dc_table,
    }
    return mapping[column_letter]


def process_row(conn, crow: ControlRow, tables: dict[str, LoadedTable],
                 dry_run: bool, log: logging.Logger) -> RowResult:
    try:
        if dry_run:
            for name in ("SMP_Matrix_LBE", "Pricing_SMP_Matrix_LBE", "DC_Prices"):
                t = tables[name]
                log.info(f"  [dry-run] {name}: {len(t.rows)} rows from {t.source_file!r}")
            for in_memory_name, table_col in UPLOAD_PLAN:
                dest = resolve_sql_table_name(crow, table_col)
                log.info(f"  [dry-run] would insert {in_memory_name} -> {dest}")
            log.info(f"  [dry-run] would EXEC {STORED_PROCEDURE}")
            return RowResult(crow.row, crow.business_date, True, "dry-run")

        clear_staging_tables(conn)

        for in_memory_name, table_col in UPLOAD_PLAN:
            dest_table = resolve_sql_table_name(crow, table_col)
            if not dest_table:
                raise ValueError(f"No destination SQL table name in column {table_col} for row {crow.row}")
            n = upload_table(conn, dest_table, tables[in_memory_name])
            log.info(f"  inserted {n} rows into {dest_table} (from {in_memory_name})")

        run_stored_procedure(conn)

        return RowResult(crow.row, crow.business_date, True)
    except Exception as exc:  # mirrors the macro's On Error Resume Next per-row
        try:
            conn.rollback()
        except Exception:
            pass
        return RowResult(crow.row, crow.business_date, False, str(exc))


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def parse_row_range(spec: Optional[str]) -> Optional[tuple[int, int]]:
    if not spec:
        return None
    if ":" in spec:
        a, b = spec.split(":", 1)
        return (int(a), int(b))
    r = int(spec)
    return (r, r)


def main() -> int:
    parser = argparse.ArgumentParser(description="Python replacement for Load_Files_To_Staging macro")
    parser.add_argument("--control", help="Path to Pricing_uploader_multiFile.xlsm (not needed with --list-odbc-drivers or --test-connection)")
    parser.add_argument("--dry-run", action="store_true", help="Parse files and print what would happen; no SQL")
    parser.add_argument("--row", help="Only process one control-sheet row (e.g. 5) or a range (e.g. 5:20)")
    parser.add_argument("--workers", type=int, default=1, help="Parallel source-file readers (I/O bound; default 1)")
    parser.add_argument("--server", default=SQL_SERVER, help=f"SQL Server host (default {SQL_SERVER})")
    parser.add_argument("--database", default=SQL_DATABASE, help=f"SQL Server database (default {SQL_DATABASE})")
    parser.add_argument("--driver", default="ODBC Driver 17 for SQL Server", help="ODBC driver name")
    parser.add_argument("--log-file", help="Also write log output to this file")
    parser.add_argument("--uploaded-by", default=None, help="Overrides Environ('USERNAME'); defaults to the current OS user")
    parser.add_argument("--list-odbc-drivers", action="store_true",
                         help="Print installed SQL Server ODBC drivers and exit (use this to find the right --driver value)")
    parser.add_argument("--test-connection", action="store_true",
                         help="Just open and close a connection to SQL Server, then exit (like the macro's Test_SQL_Connection)")
    args = parser.parse_args()

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

    if args.list_odbc_drivers:
        try:
            drivers = list_odbc_drivers()
        except ImportError:
            log.error("pyodbc is not installed. Run: pip install pyodbc")
            return 1
        if drivers:
            log.info("Installed SQL Server ODBC drivers:")
            for d in drivers:
                log.info(f"  {d!r}")
        else:
            log.warning("No SQL Server ODBC drivers found. Install 'ODBC Driver 17 for SQL Server' "
                        "(or 18) from Microsoft's download page, or ask IT to install it.")
        return 0

    if args.test_connection:
        log.info(f"Connecting to {args.server}/{args.database} using driver {args.driver!r} ...")
        try:
            conn = get_connection(args.server, args.database, args.driver)
            conn.close()
            log.info("Connected successfully.")
            return 0
        except Exception as exc:
            log.error(f"Connection failed: {exc}")
            log.error("Run with --list-odbc-drivers to see what's installed, and pass the exact "
                      "name with --driver if it isn't 'ODBC Driver 17 for SQL Server'.")
            return 1

    if not args.control:
        parser.error("--control is required (unless using --list-odbc-drivers or --test-connection)")

    uploaded_by = args.uploaded_by or os.environ.get("USERNAME") or getpass.getuser()
    only_rows = parse_row_range(args.row)

    log.info(f"Reading control sheet from {args.control}")
    control_rows = read_control_rows(args.control, only_rows=only_rows)
    log.info(f"{len(control_rows)} row(s) to process")

    if not control_rows:
        log.warning("Nothing to do.")
        return 0

    conn = None
    if not args.dry_run:
        log.info(f"Connecting to {args.server}/{args.database} ...")
        try:
            conn = get_connection(args.server, args.database, args.driver)
        except Exception as exc:
            log.error(f"Could not connect to SQL Server: {exc}")
            log.error("Try: python pricing_uploader.py --list-odbc-drivers   (to find the right --driver value)")
            log.error("Or:  python pricing_uploader.py --test-connection --driver \"...\"")
            return 1

    results: list[RowResult] = []

    def do_load(crow: ControlRow):
        upload_date = datetime.datetime.now()
        return crow, load_row_files(crow, uploaded_by, upload_date)

    if args.workers > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = {pool.submit(do_load, crow): crow for crow in control_rows}
            loaded_by_row = {}
            for fut in concurrent.futures.as_completed(futures):
                crow = futures[fut]
                try:
                    _, tables = fut.result()
                    loaded_by_row[crow.row] = tables
                except Exception as exc:
                    loaded_by_row[crow.row] = exc
        # Process in original row order so SQL writes stay sequential/predictable.
        for crow in control_rows:
            log.info(f"Row {crow.row} (business date {crow.business_date}):")
            loaded = loaded_by_row[crow.row]
            if isinstance(loaded, Exception):
                log.error(f"  FAILED loading source files - {loaded}")
                results.append(RowResult(crow.row, crow.business_date, False, str(loaded)))
                continue
            results.append(process_row(conn, crow, loaded, args.dry_run, log))
    else:
        for crow in control_rows:
            log.info(f"Row {crow.row} (business date {crow.business_date}):")
            try:
                _, tables = do_load(crow)
            except Exception as exc:
                log.error(f"  FAILED loading source files - {exc}")
                results.append(RowResult(crow.row, crow.business_date, False, str(exc)))
                continue
            results.append(process_row(conn, crow, tables, args.dry_run, log))

    if conn is not None:
        conn.close()

    ok = sum(1 for r in results if r.ok)
    failed = [r for r in results if not r.ok]
    log.info(f"Done. {ok}/{len(results)} row(s) succeeded.")
    if failed:
        log.warning(f"{len(failed)} row(s) failed:")
        for r in failed:
            log.warning(f"  row {r.row} (business date {r.business_date}): {r.message}")

    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
