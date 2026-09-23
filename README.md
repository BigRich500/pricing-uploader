# Pricing uploader (Python replacement for the macro)

This replaces the `Load_Files_To_Staging` macro in
`Pricing_uploader_multiFile.xlsm`. It reads the same `File_uploader`
control sheet, loads the same three source files per day, clears the same
staging tables, uploads with the same table mapping, and runs the same
stored procedure — just much faster, because it batches the SQL inserts
instead of sending one `INSERT` per row.

**I built and unit-tested this against fake source files that copy the
exact layout of yours (same sheet names, same column boundaries, same
sentinel-row logic). I could not test it against your actual network
share or the `ENR-HVH-EGL-02` SQL Server, because those only exist on
your work PC.** The steps below are designed to catch a bad setup safely,
in order, before it ever touches real data.

## 1. Install Python packages

On your work PC, open a command prompt (or PowerShell) and run:

```
pip install openpyxl pyodbc
```

If `pip` isn't recognized, you need Python installed first (python.org,
or ask IT — plenty of trading-floor machines already have it via
Anaconda). Any Python 3.9 or newer works.

## 2. Make sure the SQL Server ODBC driver is installed

The script talks to SQL Server the same way the macro's ADODB connection
did (Windows/Integrated authentication — no password needed), but it goes
through an ODBC driver instead. Most Windows machines that can already
open the workbook and run the macro will have one, since Excel's own
`SQLOLEDB` provider is a different thing from a standalone ODBC driver, so
it's worth checking explicitly:

```
python pricing_uploader.py --list-odbc-drivers
```

- If you see `'ODBC Driver 17 for SQL Server'` or `'ODBC Driver 18 for SQL
  Server'` in the list, you're set — the script defaults to Driver 17,
  or pass `--driver "ODBC Driver 18 for SQL Server"` if that's what's
  listed instead.
- If nothing is listed, download and install "ODBC Driver for SQL
  Server" from Microsoft (search "Microsoft ODBC Driver for SQL Server
  download"), or ask IT to push it. It's a small, standard installer —
  no admin rights beyond a normal software install.

## 3. Test the SQL Server connection on its own

Before loading any real files, confirm the script can actually reach
`ENR-HVH-EGL-02` / `FODB` from your machine:

```
python pricing_uploader.py --test-connection
```

This just opens and closes a connection — it does not touch any tables.
You should see `Connected successfully.` If it fails, the error message
will tell you whether it's a driver name problem (see step 2) or a
network/permissions problem (same as if the macro's
`Test_SQL_Connection` button failed).

## 4. Dry-run against the real control sheet

This is the most important check. It reads your real
`Pricing_uploader_multiFile.xlsm`, opens every real source file it finds,
and prints exactly what it *would* insert and into which tables —
**without connecting to SQL Server at all**:

```
python pricing_uploader.py --control "Pricing_uploader_multiFile.xlsm" --dry-run
```

Pick one or two rows you already know the answer for (a day you've
already uploaded successfully with the macro) and check with `--row`:

```
python pricing_uploader.py --control "Pricing_uploader_multiFile.xlsm" --dry-run --row 5
```

Look for:
- Row counts for each of the three files that look sane (roughly what
  you'd expect for a day's matrix/pricing/DC-prices file).
- No `FAILED loading source files` errors (that means a filename or
  folder path didn't resolve — usually a typo in the control sheet, or a
  file that's been renamed/moved since the row was created).
- The three `would insert ... -> staging....` lines. Note that
  `SMP_Matrix_LBE` is deliberately shown going into
  `staging.Pricing_SMP_Matrix_LBE`, and `Pricing_SMP_Matrix_LBE` into
  `staging.SMP_Matrix_LBE` — that crossed mapping is copied straight from
  the macro's code (not the sheet's column headers, which suggest the
  opposite). **If that's actually not what you want, tell me and I'll
  flip it** — it's a two-line change (`UPLOAD_PLAN` near the top of the
  script), but I didn't want to silently "fix" something that might be
  intentional.

## 5. Run one real row

Once the dry run looks right, do a real run against a single row first —
ideally a day you can afford to re-run if something's off, or one you can
compare against what the macro already produced:

```
python pricing_uploader.py --control "Pricing_uploader_multiFile.xlsm" --row 5
```

Check the row counts it reports match what you expect, and spot-check the
staging tables in SQL Server (or however you normally verify after the
macro runs).

## 6. Run everything

```
python pricing_uploader.py --control "Pricing_uploader_multiFile.xlsm"
```

This processes every row on the control sheet, in order, exactly like
clicking the macro's button — including that a bad row (missing file,
bad table name, etc.) is logged and skipped rather than stopping the
whole run, matching the macro's original behaviour. At the end you get a
summary line and a list of any rows that failed, with the reason.

Add `--log-file run.log` if you want a saved copy of the output to check
later or send me if something looks wrong.

## Useful options

| Flag | What it does |
|---|---|
| `--dry-run` | Parse everything, print what would happen, touch no SQL |
| `--row 5` | Only process control-sheet row 5 |
| `--row 5:20` | Only process rows 5 through 20 |
| `--workers 4` | Read up to 4 days' source files in parallel over the network while SQL writes stay in order (biggest speed win if your files are on a slow network share) |
| `--server` / `--database` / `--driver` | Override the SQL Server connection details |
| `--log-file run.log` | Also write the log to a file |
| `--list-odbc-drivers` | Show installed SQL Server ODBC drivers |
| `--test-connection` | Just check SQL connectivity and exit |
| `--help` | Full list |

## Speed notes

The macro's slowest part was almost certainly the row-by-row
`cn.Execute "INSERT ... VALUES (...)"` for the matrix files — around 1,900
separate network round trips to SQL Server per file, per day. This script
batches those into chunks of 1,000 rows using `pyodbc`'s
`fast_executemany`, so a file that took hundreds of individual round
trips now takes a couple. Source files are also read with `openpyxl` in
read-only mode, which never opens a real Excel window and reads
noticeably faster than `Workbooks.Open`, especially over a network share.

If you have many days queued up and your files sit on a slow network
drive, `--workers 4` (or higher) will read several days' files at once
while still writing to SQL Server in the original row order — try it if
a full run still feels slow after the batching change alone.

## What's identical to the macro (on purpose)

- Same control sheet (`File_uploader`), same columns (D/E/F/G/H/I/J),
  same starting row (3), same "last row = last non-blank cell in column
  E" rule.
- Same source sheet names (`SMP Matrix (Euro) INPUT`, `Euros`).
- Same matrix-file row/column boundaries: data starts at row 2, stops at
  the row before a `0` found in column A (or the last non-blank row in
  column A if no `0` is found), and always stops at column `AX`.
- Same DC Prices layout: data lives in `B5:F<last row of column C>`.
- Same three audit columns appended after the data (uploaded-by,
  upload timestamp, source filename).
- Same fixed three staging tables cleared before every day's upload,
  regardless of what the control sheet's table-name columns say.
- Same crossed table-name mapping described in step 4 above.
- Same "skip this row and keep going" behaviour on a per-row failure.

## If something goes wrong

- **`Could not connect to SQL Server`** — run `--list-odbc-drivers` and
  `--test-connection` (steps 2–3 above) to narrow it down.
- **`FAILED loading source files - [Errno 2] No such file or directory`**
  — the folder path or filename on that control-sheet row doesn't match a
  real file. Same failure mode as the macro silently failing that row;
  double-check the row in Excel.
- **A `String data, right truncation` or similar pyodbc error during
  upload** — the script already retries that one automatically with a
  slower, more tolerant insert method, so you shouldn't normally see this
  surface at all; if it does, the row will show up in the final failed-row
  list with the real underlying reason.
- Anything else: send me the console output (or the `--log-file` file) and
  I'll take a look.
