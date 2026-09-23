# Pricing uploader (Python replacement for the macro)

This replaces the `Load_Files_To_Staging` macro in
`Pricing_uploader_multiFile.xlsm`. You list the days to upload in a CSV
file; for each day the script finds the three files in that day's folder,
clears the staging tables, bulk-uploads the files to SQL Server and runs
`dbo.Process_Pricing_Data`. It's much faster than the macro because it
batches the SQL inserts instead of sending one `INSERT` per row.

## 1. Set up (first time only)

Put `pricing_uploader.py`, `requirements.txt` and your `files.csv` (step 2)
in one folder, e.g. `C:\pricing-uploader\`. Open that folder in File
Explorer, click the address bar, type `cmd` and press Enter. Then:

```
pip install -r requirements.txt
```

If `pip` isn't recognised, try `py -m pip install -r requirements.txt`.

You also need **ODBC Driver 18 for SQL Server** installed (search
"Microsoft ODBC Driver 18 for SQL Server download", or ask IT).

## 2. Make files.csv

One line per day, four columns, in this order:

| Business Date | Pricing SMP Matrix File | SMP Matrix File | DC Prices File |
|---|---|---|---|
| 01/10/2025 | Pricing SMP Matrix - LBE Sept Update (01102025).xlsx | SMP Matrix - LBE Sept Update (01102025).xlsx | DC Prices - Closing.xlsx |

- Dates are `DD/MM/YYYY`.
- Just the file names, no folders. The script works out the folder from
  the date: `\\eh-fp-01.energia.local\trading\Day Folders\YYYYMM\DD\`.
- The header line is optional.

`files.example.csv` shows the format. The quickest way to make the real
one: copy columns C, E, G and I from the `File_uploader` sheet into a new
workbook and **Save As → CSV UTF-8 (Comma delimited)** named `files.csv`.

`files.csv` is in `.gitignore`, so it never gets pushed to GitHub.

## 3. Test the SQL connection

```
python pricing_uploader.py --test-connection
```

You should see `Connected successfully.`

## 4. Dry run (reads the files, touches no SQL)

```
python pricing_uploader.py --dry-run
```

For each day it prints the folder, and how many rows it read from each
file and which staging table they'd go into. Any file it can't find shows
as `FAILED loading source files`.

If `files.csv` has a typo (a bad date or a missing file name), the script
lists every bad line and stops before uploading anything.

## 5. Upload one day, then everything

```
python pricing_uploader.py --date 01/10/2025
python pricing_uploader.py
```

A day that fails (missing file, SQL error) is logged and skipped, and the
run carries on. The summary at the end lists the failed days and why.

## Options

| Option | What it does |
|---|---|
| `--dry-run` | Read the files and print what would happen; no SQL |
| `--date 01/10/2025` | Only upload that day |
| `--from 01/10/2025` / `--to 31/10/2025` | Only upload days in that range |
| `--csv "path\to\list.csv"` | Use a different CSV (default: `files.csv` next to the script) |
| `--workers 4` | Read 4 days' files at once over the network (SQL writes stay in order) |
| `--log-file run.log` | Also save the output to a file |
| `--test-connection` | Just check SQL connectivity and exit |
| `--root` / `--server` / `--database` | Override the day folders location or SQL Server details |

## Kept exactly as the macro did it

- The **SMP Matrix** file goes into `staging.Pricing_SMP_Matrix_LBE` and
  the **Pricing SMP Matrix** file into `staging.SMP_Matrix_LBE`. That looks
  swapped, but it's what the macro did on every row. If it's a bug, flip
  `UPLOAD_PLAN` near the top of the script.
- Matrix files: sheet `SMP Matrix (Euro) INPUT`, data from row 2 to the
  row before a `0` in column A, columns A to AX.
- DC Prices file: sheet `Euros`, data in `B5:F<last row of column C>`.
- The same three audit columns after the data (uploaded by, upload time,
  source file name).

## Connection details

Windows authentication, `ENR-HVH-EGL-02` / `FODB`, ODBC Driver 18. Driver
18 encrypts the connection by default and would reject the server's
internal certificate, so the script sets `TrustServerCertificate=yes`: still
encrypted, just without checking who issued the certificate, which is what
the macro's connection did.
