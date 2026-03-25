# 🐜 Scout: 활동조사 B/C/D 섹션 데이터 로드 - 2026-03-23
"""
기후기술 활동조사 홈페이지 테이블 3개년 엑셀에서
B(b1~b4), C(c1~c8), D(d1~d3) 시트를 읽어
activity_survey_detail 테이블에 적재한다.
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")

import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).parent / "ctis_stats.db"
EXCEL_PATH = Path(__file__).parent.parent / "data" / "활동조사" / "기후기술_활동조사_홈페이지_테이블_3개년.xlsx"

SHEETS = [
    "b1", "b2", "b3", "b4",
    "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8",
    "d1", "d2", "d3",
]

# d1/d2/d3 use 점 (5-point scale), rest use %
UNIT_MAP = {
    "d1": "점", "d2": "점", "d3": "점",
}

META_COLS = ["조사연도", "대분류", "기술코드", "기술분야", "사례수(N)"]


def create_table(conn: sqlite3.Connection):
    conn.execute("DROP TABLE IF EXISTS activity_survey_detail")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS activity_survey_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            survey_year INTEGER NOT NULL,
            section TEXT NOT NULL,
            category_code TEXT,
            category_name TEXT,
            type_name TEXT,
            sample_size INTEGER,
            item_name TEXT NOT NULL,
            item_value REAL,
            item_unit TEXT DEFAULT '%'
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_asd_section_year
        ON activity_survey_detail(section, survey_year)
    """)
    conn.commit()


def load_sheet(conn: sqlite3.Connection, sheet_name: str, xl_path: str):
    df = pd.read_excel(xl_path, sheet_name=sheet_name, header=1)

    # Identify item columns (everything after the 5 meta cols)
    item_cols = [c for c in df.columns if c not in META_COLS]

    unit = UNIT_MAP.get(sheet_name, "%")
    rows = []

    for _, row in df.iterrows():
        year = int(row["조사연도"])
        type_name = str(row["대분류"])
        cat_code = str(row["기술코드"])
        cat_name = str(row["기술분야"])
        sample = int(row["사례수(N)"]) if pd.notna(row["사례수(N)"]) else 0

        for item_col in item_cols:
            val = row[item_col]
            if pd.isna(val):
                continue
            rows.append((
                year, sheet_name, cat_code, cat_name, type_name,
                sample, str(item_col), float(val), unit,
            ))

    conn.executemany("""
        INSERT INTO activity_survey_detail
            (survey_year, section, category_code, category_name, type_name,
             sample_size, item_name, item_value, item_unit)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()
    return len(rows)


def main():
    print(f"DB: {DB_PATH}")
    print(f"Excel: {EXCEL_PATH}")

    if not EXCEL_PATH.exists():
        print(f"ERROR: Excel file not found: {EXCEL_PATH}")
        return

    conn = sqlite3.connect(str(DB_PATH))
    create_table(conn)

    total = 0
    for sheet in SHEETS:
        n = load_sheet(conn, sheet, str(EXCEL_PATH))
        print(f"  {sheet}: {n} rows loaded")
        total += n

    print(f"\nTotal: {total} rows loaded into activity_survey_detail")

    # Verify
    cur = conn.execute("""
        SELECT section, COUNT(*), COUNT(DISTINCT survey_year)
        FROM activity_survey_detail
        GROUP BY section ORDER BY section
    """)
    print("\nVerification:")
    for row in cur:
        print(f"  {row[0]}: {row[1]} rows, {row[2]} years")

    conn.close()


if __name__ == "__main__":
    main()
