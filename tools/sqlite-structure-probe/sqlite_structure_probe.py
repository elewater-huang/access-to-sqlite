# -*- coding: utf-8 -*-
"""
SQLite 結構探針 v0.2
Windows 7 / Python 3.7+
不需額外套件。

輸出：
1) schema_report_YYYYMMDD_HHMMSS.txt
2) relationship_clues_YYYYMMDD_HHMMSS.csv

用途：快速看清大型 SQLite 的表、欄位、索引、Foreign Key，
以及製令／流程／料盒／基板／Barcode 等疑似關聯欄位。
唯讀連線；資料範例預設關閉，使用 --include-samples 才輸出。
輸出固定放在本 .py 所在目錄。
"""

import os
import sys
import csv
import sqlite3
import datetime
import traceback
import argparse
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    tk = None

KEYWORDS = [
    "製令", "制令", "流程", "流程單", "料盒", "盒號",
    "基板", "板號", "基板號", "barcode", "bar_code", "條碼",
    "board", "panel", "lot", "batch", "order", "workorder", "work_order",
    "工單", "單號", "黑", "藍", "蓝", "綠", "绿", "紅", "红",
    "不良", "bad", "defect"
]


def qident(name):
    return '"' + str(name).replace('"', '""') + '"'


def output_folder():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def choose_db():
    if tk is None:
        return None
    root = tk.Tk()
    root.withdraw()
    root.update()
    path = filedialog.askopenfilename(
        title="選擇 SQLite 資料庫",
        filetypes=[
            ("SQLite Database", "*.db *.sqlite *.sqlite3"),
            ("DB", "*.db"),
            ("All files", "*.*")
        ]
    )
    root.destroy()
    return path


def safe_text(v, limit=240):
    if v is None:
        return ""
    if isinstance(v, bytes):
        return "<BLOB %d bytes>" % len(v)
    s = str(v).replace("\r", "\\r").replace("\n", "\\n")
    if len(s) > limit:
        s = s[:limit] + "...<truncated>"
    return s


def is_candidate(name):
    low = str(name).lower()
    return any(str(k).lower() in low for k in KEYWORDS)


def sample_values(conn, table, column, limit=5):
    sql = (
        "SELECT DISTINCT {c} FROM {t} "
        "WHERE {c} IS NOT NULL AND TRIM(CAST({c} AS TEXT)) <> '' "
        "LIMIT {n}"
    ).format(c=qident(column), t=qident(table), n=int(limit))
    out = []
    try:
        for row in conn.execute(sql):
            out.append(safe_text(row[0], 180))
    except Exception as e:
        out.append("<ERROR: %s>" % safe_text(e, 160))
    return out


def exact_count(conn, table):
    try:
        r = conn.execute("SELECT COUNT(*) FROM %s" % qident(table)).fetchone()
        return r[0] if r else ""
    except Exception:
        return ""


def probe(db_path, include_samples=False):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = output_folder()
    report_path = os.path.join(outdir, "schema_report_%s.txt" % stamp)
    keys_path = os.path.join(outdir, "relationship_clues_%s.csv" % stamp)

    db_file = Path(db_path).resolve()
    if not db_file.is_file():
        raise FileNotFoundError("找不到資料庫：%s" % db_file.name)
    conn = sqlite3.connect(db_file.as_uri() + "?mode=ro", uri=True)
    conn.execute("PRAGMA query_only = ON")
    try:

        objects = conn.execute("""
            SELECT type, name, tbl_name, sql
            FROM sqlite_master
            WHERE type IN ('table','view')
              AND name NOT LIKE 'sqlite_%'
            ORDER BY type, name
        """).fetchall()

        candidate_rows = []

        with open(report_path, "w", encoding="utf-8-sig") as f:
            f.write("SQLite 結構探針報告\n")
            f.write("=" * 90 + "\n")
            f.write("DB: %s\n" % db_file.name)
            f.write("Time: %s\n" % datetime.datetime.now().isoformat(sep=" "))
            f.write("Objects: %d\n" % len(objects))
            f.write("Data samples: %s\n" % ("included" if include_samples else "disabled"))
            f.write("=" * 90 + "\n\n")

            for i, (obj_type, name, tbl_name, create_sql) in enumerate(objects, 1):
                f.write("[%d/%d] %s: %s\n" % (i, len(objects), obj_type.upper(), name))
                f.write("-" * 90 + "\n")

                row_count = ""
                if obj_type == "table":
                    row_count = exact_count(conn, name)
                    f.write("Rows: %s\n" % row_count)

                f.write("\nCREATE SQL:\n%s\n\n" % (create_sql or "(none)"))

                cols = conn.execute("PRAGMA table_info(%s)" % qident(name)).fetchall()
                f.write("COLUMNS:\n")
                f.write("cid | name | type | notnull | default | pk\n")
                f.write("-" * 70 + "\n")

                for col in cols:
                    cid, col_name, col_type, notnull, default_val, pk = col
                    f.write("%s | %s | %s | %s | %s | %s\n" % (
                        cid, safe_text(col_name), safe_text(col_type), notnull,
                        safe_text(default_val), pk
                    ))

                    if is_candidate(col_name):
                        vals = sample_values(conn, name, col_name, 5) if include_samples else []
                        candidate_rows.append({
                            "object_type": obj_type,
                            "table_name": name,
                            "column_name": col_name,
                            "declared_type": col_type,
                            "is_pk": pk,
                            "row_count": row_count,
                            "sample_1": vals[0] if len(vals) > 0 else "",
                            "sample_2": vals[1] if len(vals) > 1 else "",
                            "sample_3": vals[2] if len(vals) > 2 else "",
                            "sample_4": vals[3] if len(vals) > 3 else "",
                            "sample_5": vals[4] if len(vals) > 4 else ""
                        })

                if obj_type == "table":
                    f.write("\nFOREIGN KEYS:\n")
                    try:
                        fks = conn.execute("PRAGMA foreign_key_list(%s)" % qident(name)).fetchall()
                    except Exception:
                        fks = []
                    if not fks:
                        f.write("(none)\n")
                    else:
                        for fk in fks:
                            f.write("from=%s -> %s.%s | update=%s | delete=%s\n" % (
                                safe_text(fk[3] if len(fk) > 3 else ""),
                                safe_text(fk[2] if len(fk) > 2 else ""),
                                safe_text(fk[4] if len(fk) > 4 else ""),
                                safe_text(fk[5] if len(fk) > 5 else ""),
                                safe_text(fk[6] if len(fk) > 6 else "")
                            ))

                    f.write("\nINDEXES:\n")
                    try:
                        indexes = conn.execute("PRAGMA index_list(%s)" % qident(name)).fetchall()
                    except Exception:
                        indexes = []
                    if not indexes:
                        f.write("(none)\n")
                    else:
                        for idx in indexes:
                            idx_name = idx[1] if len(idx) > 1 else ""
                            unique = idx[2] if len(idx) > 2 else ""
                            origin = idx[3] if len(idx) > 3 else ""
                            cols2 = []
                            if idx_name:
                                try:
                                    info = conn.execute("PRAGMA index_info(%s)" % qident(idx_name)).fetchall()
                                    cols2 = [x[2] for x in info if len(x) > 2]
                                except Exception:
                                    pass
                            f.write("%s | unique=%s | origin=%s | cols=%s\n" % (
                                safe_text(idx_name), unique, safe_text(origin),
                                ", ".join(safe_text(x) for x in cols2)
                            ))

                f.write("\n" + "=" * 90 + "\n\n")

        fields = [
            "object_type", "table_name", "column_name", "declared_type",
            "is_pk", "row_count", "sample_1", "sample_2", "sample_3",
            "sample_4", "sample_5"
        ]
        with open(keys_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            for row in candidate_rows:
                w.writerow(row)

        return report_path, keys_path, len(objects), len(candidate_rows)

    finally:
        conn.close()

def main():
    parser = argparse.ArgumentParser(description="SQLite 結構探針：唯讀檢視資料庫結構")
    parser.add_argument("database", nargs="?", help="SQLite 資料庫路徑")
    parser.add_argument("--include-samples", action="store_true",
                        help="匯出實際資料範例（預設關閉）")
    args = parser.parse_args()
    db_path = args.database or choose_db()
    if not db_path:
        print("未選擇資料庫。")
        return
    if not os.path.isfile(db_path):
        print("找不到資料庫：", db_path)
        input("\n按 Enter 結束...")
        return

    try:
        report, keys, obj_count, key_count = probe(db_path, include_samples=args.include_samples)
        print("=" * 72)
        print("完成")
        print("Table/View：", obj_count)
        print("疑似關聯欄位：", key_count)
        print("結構報告：", report)
        print("候選欄位：", keys)
        print("=" * 72)

        if tk is not None:
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo(
                    "SQLite 結構探針完成",
                    "Table/View：%d\n疑似關聯欄位：%d\n\n%s\n\n%s" %
                    (obj_count, key_count, report, keys)
                )
                root.destroy()
            except Exception:
                pass
    except Exception as e:
        print("\n失敗：", e)
        traceback.print_exc()

    input("\n按 Enter 結束...")


if __name__ == "__main__":
    main()
