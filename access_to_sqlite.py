# -*- coding: utf-8 -*-
"""
Access -> SQLite Converter v1.1

重點修正：
- 不再使用 pyodbc cursor.columns()，避開舊 Access/Jet ODBC metadata 的 UTF-16 解碼問題。
- 改由 SELECT * FROM [table] WHERE 1=0 + cursor.description 取得欄位資訊。
- 每張表使用 SAVEPOINT；單一資料表失敗時記錄後繼續，不讓整批轉換中止。
- 不覆蓋既有 .db，自動輸出 _1、_2 ...。

需求：
    pip install pyodbc

適用：Windows / Python 3.7+
"""

import os
import sys
import sqlite3
import datetime
import decimal
import uuid
import traceback

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    tk = None

try:
    import pyodbc
except ImportError:
    print("缺少 pyodbc。")
    print("請先執行：pip install pyodbc")
    input("\n按 Enter 結束...")
    sys.exit(1)


VERSION = "1.1"


def qident(name):
    return '"' + str(name).replace('"', '""') + '"'


def access_qident(name):
    return "[" + str(name).replace("]", "]]" ) + "]"


def sqlite_type_from_python(type_code):
    """pyodbc cursor.description 的 type_code -> SQLite affinity。"""
    if type_code is None:
        return "TEXT"

    # 常見 Python 型態
    try:
        if type_code is bool:
            return "INTEGER"
        if type_code is int:
            return "INTEGER"
        if type_code is float:
            return "REAL"
        if type_code is decimal.Decimal:
            return "NUMERIC"
        if type_code in (bytes, bytearray, memoryview):
            return "BLOB"
        if type_code in (datetime.datetime, datetime.date, datetime.time):
            return "TEXT"
        if type_code is str:
            return "TEXT"
    except Exception:
        pass

    # 某些 driver 回傳不同 class，退回 class 名稱判定
    name = getattr(type_code, "__name__", str(type_code)).lower()

    if any(x in name for x in ("bool", "int", "long", "short", "byte")):
        return "INTEGER"
    if any(x in name for x in ("float", "double", "real")):
        return "REAL"
    if any(x in name for x in ("decimal", "numeric", "money", "currency")):
        return "NUMERIC"
    if any(x in name for x in ("binary", "bytes", "blob", "image")):
        return "BLOB"

    return "TEXT"


def normalize_value(v):
    """轉成 sqlite3 可接受的值。"""
    if v is None:
        return None

    if isinstance(v, bool):
        return 1 if v else 0

    if isinstance(v, datetime.datetime):
        return v.isoformat(sep=" ")

    if isinstance(v, (datetime.date, datetime.time)):
        return v.isoformat()

    if isinstance(v, decimal.Decimal):
        return format(v, "f")

    if isinstance(v, uuid.UUID):
        return str(v)

    if isinstance(v, bytearray):
        return bytes(v)

    if isinstance(v, memoryview):
        return bytes(v)

    if isinstance(v, (str, int, float, bytes)):
        return v

    return str(v)


def get_access_driver():
    drivers = pyodbc.drivers()

    preferred = [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.mdb)",
    ]

    for d in preferred:
        if d in drivers:
            return d

    for d in drivers:
        if "Access Driver" in d:
            return d

    return None


def get_user_tables(cursor):
    tables = []

    for row in cursor.tables(tableType="TABLE"):
        name = row.table_name
        if not name:
            continue

        name = str(name)
        if name.startswith("MSys"):
            continue

        tables.append(name)

    # 去重保序
    seen = set()
    result = []
    for t in tables:
        if t not in seen:
            seen.add(t)
            result.append(t)

    return result


def get_columns_by_select(acc_connection, table_name):
    """
    關鍵修正：不用 cursor.columns()。
    Access 執行空結果 SELECT，從 cursor.description 取欄位名稱與 Python type。
    """
    cur = acc_connection.cursor()
    sql = "SELECT * FROM %s WHERE 1=0" % access_qident(table_name)
    cur.execute(sql)

    desc = cur.description
    if not desc:
        return []

    cols = []
    for d in desc:
        # DB-API description:
        # name, type_code, display_size, internal_size, precision, scale, null_ok
        col_name = d[0]
        type_code = d[1] if len(d) > 1 else None

        cols.append({
            "name": str(col_name),
            "sqlite_type": sqlite_type_from_python(type_code),
        })

    return cols


def choose_output_path(access_path):
    base_db_path = os.path.splitext(access_path)[0] + ".db"
    db_path = base_db_path

    if not os.path.exists(db_path):
        return db_path

    base, ext = os.path.splitext(base_db_path)
    n = 1
    while os.path.exists("%s_%d%s" % (base, n, ext)):
        n += 1

    return "%s_%d%s" % (base, n, ext)


def convert_access_to_sqlite(access_path, db_path=None, batch_size=500):
    access_path = os.path.abspath(access_path)

    if db_path is None:
        db_path = choose_output_path(access_path)
    else:
        db_path = os.path.abspath(db_path)
        if os.path.exists(db_path):
            raise RuntimeError("輸出 SQLite 已存在，不覆蓋：%s" % db_path)

    driver = get_access_driver()
    if not driver:
        raise RuntimeError(
            "找不到 Microsoft Access ODBC Driver。\n\n"
            "請確認已安裝 Access / Access Database Engine，且 Python 與 ODBC Driver 位元數相容。"
        )

    print("=" * 72)
    print("Access -> SQLite Converter v%s" % VERSION)
    print("來源：", access_path)
    print("輸出：", db_path)
    print("ODBC：", driver)
    print("=" * 72)

    conn_str = r"DRIVER={%s};DBQ=%s;" % (driver, access_path)
    acc = pyodbc.connect(conn_str)
    acc_cur = acc.cursor()

    sq = sqlite3.connect(db_path)
    sq.execute("PRAGMA journal_mode=WAL;")
    sq.execute("PRAGMA synchronous=NORMAL;")
    sq.execute("PRAGMA foreign_keys=OFF;")

    summary = []

    try:
        tables = get_user_tables(acc_cur)

        if not tables:
            raise RuntimeError("沒有找到可匯出的 Access TABLE。")

        print("找到 %d 張資料表。\n" % len(tables))

        for idx, table in enumerate(tables, 1):
            print("[%d/%d] %s" % (idx, len(tables), table))

            savepoint = "sp_table_%d" % idx
            sq.execute("SAVEPOINT %s" % savepoint)

            try:
                # 先直接 SELECT 取 schema，避開 SQLColumns metadata bug
                cols = get_columns_by_select(acc, table)

                if not cols:
                    raise RuntimeError("無法取得欄位資訊")

                col_defs = [
                    "%s %s" % (qident(c["name"]), c["sqlite_type"])
                    for c in cols
                ]

                create_sql = "CREATE TABLE %s (%s)" % (
                    qident(table),
                    ", ".join(col_defs)
                )
                sq.execute(create_sql)

                src = acc.cursor()
                select_sql = "SELECT * FROM %s" % access_qident(table)
                src.execute(select_sql)

                placeholders = ",".join(["?"] * len(cols))
                insert_sql = "INSERT INTO %s VALUES (%s)" % (
                    qident(table), placeholders
                )

                count = 0

                while True:
                    rows = src.fetchmany(batch_size)
                    if not rows:
                        break

                    converted = [
                        tuple(normalize_value(v) for v in row)
                        for row in rows
                    ]

                    sq.executemany(insert_sql, converted)
                    count += len(converted)

                    print("\r  已匯入 %d 筆" % count, end="")
                    sys.stdout.flush()

                sq.execute("RELEASE SAVEPOINT %s" % savepoint)
                sq.commit()

                print("\r  完成：%d 筆          " % count)
                summary.append((table, count, "OK", ""))

            except Exception as table_error:
                try:
                    sq.execute("ROLLBACK TO SAVEPOINT %s" % savepoint)
                    sq.execute("RELEASE SAVEPOINT %s" % savepoint)
                except Exception:
                    pass

                msg = "%s: %s" % (type(table_error).__name__, str(table_error))
                print("\n  !! 此表失敗，已略過：%s" % msg)
                summary.append((table, 0, "FAILED", msg))
                continue

        # 匯入資訊
        sq.execute('''
            CREATE TABLE IF NOT EXISTS "__ACCESS_IMPORT_INFO__" (
                key TEXT,
                value TEXT
            )
        ''')

        ok_count = sum(1 for x in summary if x[2] == "OK")
        fail_count = sum(1 for x in summary if x[2] == "FAILED")

        sq.executemany(
            'INSERT INTO "__ACCESS_IMPORT_INFO__" (key, value) VALUES (?, ?)',
            [
                ("converter_version", VERSION),
                ("source_access_file", access_path),
                ("import_time", datetime.datetime.now().isoformat(sep=" ")),
                ("odbc_driver", driver),
                ("table_count", str(len(tables))),
                ("success_table_count", str(ok_count)),
                ("failed_table_count", str(fail_count)),
            ]
        )

        # 失敗表也寫進 DB，方便追查
        sq.execute('''
            CREATE TABLE IF NOT EXISTS "__ACCESS_IMPORT_ERRORS__" (
                table_name TEXT,
                error_message TEXT
            )
        ''')

        for table, count, status, err in summary:
            if status == "FAILED":
                sq.execute(
                    'INSERT INTO "__ACCESS_IMPORT_ERRORS__" (table_name, error_message) VALUES (?, ?)',
                    (table, err)
                )

        sq.commit()

        print("\n" + "=" * 72)
        print("轉換完成")
        print("SQLite：", db_path)
        print("成功：%d 張表 / 失敗：%d 張表 / 共：%d 張表" % (
            ok_count, fail_count, len(summary)
        ))
        print("-" * 72)

        for table, count, status, err in summary:
            if status == "OK":
                print("[OK]   %-42s %10d 筆" % (table[:42], count))
            else:
                print("[FAIL] %-42s %s" % (table[:42], err))

        print("=" * 72)

        return db_path, summary

    except Exception:
        sq.rollback()
        raise

    finally:
        try:
            sq.close()
        except Exception:
            pass
        try:
            acc.close()
        except Exception:
            pass


def choose_access_file():
    if tk is None:
        return None

    root = tk.Tk()
    root.withdraw()
    root.update()

    path = filedialog.askopenfilename(
        title="選擇 Access 資料庫",
        filetypes=[
            ("Access Database", "*.mdb *.accdb"),
            ("MDB", "*.mdb"),
            ("ACCDB", "*.accdb"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()
    return path


def main():
    if len(sys.argv) >= 2:
        access_path = sys.argv[1]
    else:
        access_path = choose_access_file()

    if not access_path:
        print("未選擇檔案。")
        return

    if not os.path.isfile(access_path):
        print("找不到檔案：", access_path)
        input("\n按 Enter 結束...")
        return

    ext = os.path.splitext(access_path)[1].lower()
    if ext not in (".mdb", ".accdb"):
        print("這不是 .mdb / .accdb：", access_path)
        input("\n按 Enter 結束...")
        return

    try:
        db_path, summary = convert_access_to_sqlite(access_path)

        failed = [x for x in summary if x[2] == "FAILED"]
        msg = "轉換完成：\n\n%s\n\n成功 %d 張，失敗 %d 張" % (
            db_path,
            len(summary) - len(failed),
            len(failed),
        )

        if tk is not None:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Access -> SQLite", msg)
            root.destroy()

    except Exception as e:
        print("\n轉換失敗：")
        print(str(e))
        print("\n詳細錯誤：")
        traceback.print_exc()

        if tk is not None:
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Access -> SQLite 失敗", str(e))
                root.destroy()
            except Exception:
                pass

    input("\n按 Enter 結束...")


if __name__ == "__main__":
    main()
