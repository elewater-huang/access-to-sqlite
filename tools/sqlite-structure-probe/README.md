# SQLite Structure Probe v0.2 / SQLite 結構探針

Step two after exporting Access data: inspect the SQLite structure without opening the Access application. Python 3.7+ standard library only; no Access driver is needed for this step.

## Usage / 使用方式

```bat
python tools\sqlite-structure-probe\sqlite_structure_probe.py "C:\examples\legacy.db"
```

Run without arguments to select a SQLite file through a dialog (requires tkinter and a graphical desktop). The program pauses for Enter when finished and is intended for interactive use.

Outputs beside the script (the directory must be writable):

- `schema_report_YYYYMMDD_HHMMSS.txt`: tables, views, CREATE SQL, columns, exact table row counts, indexes and declared foreign keys.
- `relationship_clues_YYYYMMDD_HHMMSS.csv`: columns matching editable keywords in `KEYWORDS`. Sample fields are blank by default.

Optional actual data samples, up to five distinct nonempty values per matching column:

```bat
python tools\sqlite-structure-probe\sqlite_structure_probe.py "C:\examples\legacy.db" --include-samples
```

The database connection enforces read-only mode and query-only access. Missing database paths are rejected. The report header includes only the database filename, not its full path. Samples are disabled unless explicitly requested. Table names, field names and CREATE SQL can still contain sensitive information; review reports before sharing. Exact counts and optional DISTINCT sampling may take time on large databases. Run against a stable copy for consistent results. Outputs use timestamps with one-second precision; a repeated run in the same second can overwrite that run's reports.

## Scope / 重要範圍

This tool reads SQLite, not MDB/ACCDB. It does not infer or validate relationships, identify unused tables, or inspect Access queries, forms or VBA. Keyword matches are clues, not candidate keys.

**The preceding Access exporter does not preserve Access primary keys, indexes or relationships.** An empty foreign-key/index section in its SQLite output does not mean those structures never existed in Access. This probe reports only metadata actually present in SQLite.

Access 適合個人或小公司快速建立管理工具，但經過多人接手與修改，資料表的用途可能逐漸不清楚。先將資料匯出到 SQLite，再用這支工具把資料表、欄位與筆數整理成文字報告，讓結構可以單獨閱讀、比較與討論。

預設不匯出實際資料範例；需要時才加上 `--include-samples`。程式以唯讀方式連線，報告只列資料庫檔名。關聯線索來自欄位名稱比對，不代表已確認關聯，也不能據此判定資料表無用。前一步匯出工具沒有保留 Access 原本的主鍵、索引與關聯，因此不能由這份報告還原完整的 Access 設計。

## Validation and license

Validated in Linux using synthetic SQLite data: table/view metadata, counts, indexes, foreign keys, default sample suppression, opt-in samples, filename-only header, Unicode/space/# filenames, rejected missing paths, and unchanged database content hash. SQLite read-only connections reject writes. Source parses with Python 3.7 grammar. Windows 7, Python 3.7 runtime and GUI behavior have not been tested for this version.

MIT License; see the repository's [LICENSE](../../LICENSE).
