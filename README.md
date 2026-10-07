# Access to SQLite

A small Windows utility for taking a first look at data stored in an old Microsoft Access database.

Legacy data often lives inside a legacy application. Exporting through Excel adds another step and can run into worksheet limits. This tool extracts table data directly into SQLite so you can browse it, query it, and decide what to do next.

## Quick start

Requirements: Windows, Python 3.7+, pyodbc, and a Microsoft Access ODBC driver compatible with the Python process architecture. An MDB-only driver cannot open ACCDB files.

```bat
python -m pip install pyodbc
python access_to_sqlite.py "C:\examples\legacy.mdb"
```

Run without arguments to select a file through a dialog. The program pauses for Enter at the end; it is intended for interactive use.

The output is created beside the source: `legacy.db`. If that name exists, the program chooses `legacy_1.db`, then `legacy_2.db`, and so on. The source directory therefore needs write access for the output.

## What it does

- Reads user tables through Access ODBC and inserts rows in batches of 500.
- Obtains column names and types through an empty SELECT result rather than SQLColumns metadata, addressing a legacy driver decoding issue.
- Rolls back a failed table and continues with the remaining tables.
- Prints imported row counts and stores import information and errors in SQLite.

```sql
SELECT * FROM "__ACCESS_IMPORT_INFO__";
SELECT * FROM "__ACCESS_IMPORT_ERRORS__";
SELECT name FROM sqlite_master WHERE type = 'table';
-- Replace sample_table with an actual imported table name.
SELECT * FROM "sample_table" LIMIT 20;
```

## Scope

This is a data exploration utility. It does not rebuild Access relationships, indexes, primary keys, queries, forms, or VBA. Imported row counts are progress records, not a source-to-target completeness verification. Dates become text, booleans become integers, and NUMERIC values may change precision in SQLite.

The code issues no data or schema writes to Access, but the connection does not explicitly enforce read-only mode. Export a copy while the source is idle for a more stable inspection. Failed tables are omitted; review the summary before using the output. Extraction uses Windows Access ODBC; the resulting SQLite file can be inspected on other platforms.

## 中文說明

這是把老舊 Access 資料抽到 SQLite「先看看」的小工具。它省去 Excel 中轉，讓你先用 SQL 看懂資料，再決定如何整理或遷移。

安裝 pyodbc 與相符位元數的 Access ODBC 驅動後，可直接執行程式選檔，或在命令列指定 .mdb／.accdb。輸出位於來源旁邊，不覆蓋既有 .db；失敗資料表會回滾並記錄，其餘表繼續抽取。

它以資料探索為目的，不保證完整重建 Access 結構，也沒有做來源與目的資料的完整比對。輸出資料可用 SQLite 檢視工具開啟，或搬到 Linux 繼續分析。

## Step two: inspect the exported structure / 第二步：檢視匯出後的結構

[SQLite Structure Probe v0.2](tools/sqlite-structure-probe/) generates a text schema report and a CSV of relationship clues from the exported SQLite database. Read-only connection; actual data samples are disabled by default. Uses the Python standard library only.

先匯出資料，再把資料表、欄位與筆數攤開來看。新工具的使用方式與限制請見上方連結。它僅檢視 SQLite 現有結構，無法還原匯出時未保留的 Access 主鍵、索引或關聯。

## Validation and license

The source has been reviewed and parsed for Python syntax. End-to-end Access extraction has not been tested in this preparation environment, which has no Windows Access driver or sample Access database.

MIT License. Copyright (c) 2026 Elewater Huang. See [LICENSE](LICENSE).
