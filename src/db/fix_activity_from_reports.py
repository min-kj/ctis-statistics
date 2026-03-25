"""
Fix activity_survey DB data using official report values.

Sources:
  - 2022: HTML tables from CTis homepage (verified against 2022 report PDF)
  - 2023: PDF tables from 2023 report (PyMuPDF extraction)

Fixes metric_types: revenue_climate, rnd_expense_climate, employee_count, employee_count_climate
Keeps revenue_no_reason data unchanged.
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import re
import sqlite3
from bs4 import BeautifulSoup
import fitz  # PyMuPDF

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # src/
DB_PATH = os.path.join(BASE, 'db', 'ctis_stats.db')
HTML_PATH = os.path.join(
    BASE,
    '기술개발활동조사 _ 통계정보 _ 기후기술통계 _ '
    '국가기후기술정보시스템 CTis - 기후변화 대응과 기후기술 협력 플랫폼.html'
)
PDF_2023 = os.path.join(BASE, 'data', '활동조사',
                        '2023 기후기술개발_활동조사_공표_보고서_251121_2.pdf')

# org_size column order in both HTML and PDF tables
ORG_SIZES = ['100억미만', '100~500억', '500~1000억', '1000~2000억', '2000억이상', '공공기관등', '전체']

# category name -> DB category_id (T22, taxonomy_id=3)
CATEGORY_MAP = {
    '태양광·열': 83, '태양광.열': 83,
    '풍력': 84,
    '해양에너지 및 수력': 85, '해양 및 수력에너지': 85,
    '수열 및 지열': 86,
    '바이오에너지': 87, '바이오': 87,
    '수소암모니아 발전': 88, '수소암모니아': 88,
    '비재생에너지': 89,
    '수소·바이오매스': 90, '수소바이오매스': 90,
    '폐자원': 91,
    '발전효율': 92,
    '산업효율': 93,
    '수송효율': 94,
    '건물효율': 95,
    '온실가스 저장·흡수·활용': 96, '포집/저장/흡수/대체': 96,
    '전력·열 통합': 97, '전력통합': 97,
    '기후변화 모니터링': 98,
    '기후영향평가·진단': 99, '기후영향평가 및 진단': 99,
    '건강': 100,
    '물': 101,
    '농축수산': 102,
    '국토·연안, 산림·생태계, 산업·에너지': 103, '기타 탄력성 강화': 103,
    '적응정보·평가': 104, '적응기반': 104,
}

# Ordered list of 22 category names as they appear in tables
CATEGORY_ORDER_GAMCHUK = [
    '태양광·열', '풍력', '해양에너지 및 수력', '수열 및 지열',
    '바이오에너지', '수소암모니아 발전', '비재생에너지', '수소·바이오매스',
    '폐자원', '발전효율', '산업효율', '수송효율', '건물효율',
    '온실가스 저장·흡수·활용', '전력·열 통합',
]  # 15 categories
CATEGORY_ORDER_JEOKEUNG = [
    '기후변화 모니터링', '기후영향평가·진단', '건강', '물',
    '농축수산', '국토·연안, 산림·생태계, 산업·에너지', '적응정보·평가',
]  # 7 categories
CATEGORY_ORDER = CATEGORY_ORDER_GAMCHUK + CATEGORY_ORDER_JEOKEUNG  # 22 total


def parse_number(s):
    """Parse a number string like '6,761,802' or '-' into float or None."""
    if s is None:
        return None
    s = str(s).strip()
    if s in ('-', '', 'None', '.'):
        return None
    s = s.replace(',', '').replace('\n', '')
    try:
        return float(s)
    except ValueError:
        return None


# ──────────────────────────────────────────────────────────────────────
# 2022: Parse from HTML tables
# ──────────────────────────────────────────────────────────────────────

def parse_html_table(soup, table_id):
    """
    Parse an HTML table (a1, a3, a4, a5) into list of (category_id, org_size, value).
    Tables have structure:
      thead: 2 rows (headers)
      tbody: 25 rows (합계, 감축소계, 15 감축 categories, 적응소계, 7 적응 categories)
    Each row has: [group_col?, category_name, val_100, val_100_500, val_500_1000, val_1000_2000, val_2000, val_public, val_total]
    """
    table = soup.find('table', id=table_id)
    if not table:
        print(f"  ERROR: Table {table_id} not found in HTML")
        return []

    tbody = table.find('tbody')
    rows = tbody.find_all('tr')
    results = []

    for tr in rows:
        cells = [c.text.strip() for c in tr.find_all(['th', 'td'])]

        # Determine the category name and values
        # Row formats:
        #   합계: ['합계', v1, v2, v3, v4, v5, v6, v_total]  (8 cells)
        #   감축 소계: ['감축', '소계', v1, ..., v_total]  (9 cells)
        #   감축 category: ['01. 태양광·열', v1, ..., v_total]  (8 cells)
        #   적응 소계: ['적응', '소계', v1, ..., v_total]  (9 cells)
        #   적응 category: ['16. 기후변화 모니터링', v1, ..., v_total]  (8 cells)

        if len(cells) == 9:
            # Subtotal row (감축/적응 소계) - skip
            continue
        elif len(cells) == 8:
            name_raw = cells[0]
            values = cells[1:8]  # 7 values: 6 org_sizes + 전체
        else:
            continue

        # Clean category name
        name = re.sub(r'^\d{1,2}\.\s*', '', name_raw).strip()

        if name in ('합계', '소계'):
            # 합계 row - store with category_id=None
            for i, org in enumerate(ORG_SIZES):
                v = parse_number(values[i])
                if v is not None:
                    results.append((None, org, v))
            continue

        # Match category
        cat_id = CATEGORY_MAP.get(name)
        if cat_id is None:
            print(f"  WARNING: No match for category '{name}' in table {table_id}")
            continue

        for i, org in enumerate(ORG_SIZES):
            v = parse_number(values[i])
            if v is not None:
                results.append((cat_id, org, v))

    return results


def load_2022_from_html():
    """Load 2022 data from CTis HTML tables."""
    print(f"Loading 2022 HTML: {os.path.basename(HTML_PATH)}")
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    data = {}
    table_map = {
        'a1': ('revenue_climate', '백만원'),
        'a3': ('rnd_expense_climate', '백만원'),
        'a4': ('employee_count', '명'),
        'a5': ('employee_count_climate', '명'),
    }

    for tid, (metric, unit) in table_map.items():
        rows = parse_html_table(soup, tid)
        data[(2022, metric)] = [(cat_id, org, val, unit) for cat_id, org, val in rows]
        print(f"  {tid} ({metric}): {len(rows)} data points")

    return data


# ──────────────────────────────────────────────────────────────────────
# 2023: Parse from PDF tables
# ──────────────────────────────────────────────────────────────────────

def parse_pdf_table(doc, page_num, table_idx):
    """
    Parse a PDF table from the 2023 report.
    PyMuPDF packs merged cell values separated by \\n.
    Returns list of (category_id, org_size, value).
    """
    page = doc[page_num]
    tables = page.find_tables()
    if table_idx >= len(tables.tables):
        print(f"  ERROR: Page {page_num} has only {len(tables.tables)} tables, requested index {table_idx}")
        return []

    tab = tables.tables[table_idx]
    df = tab.to_pandas()
    results = []

    # Column indices (0-based): 0=대분류, 1=중분류, 2-7=org_sizes, 8=전체
    # But df columns might be named differently
    # The actual value columns are indices 2..8 (7 columns matching ORG_SIZES)

    # Row 0 is header row (대분류/중분류), skip it
    # Row 1 is 합계
    # Row 2 is 감축 소계
    # Row 3 is first 감축 category (but values for all 15 are packed with \n)
    # Rows 4-17 are empty (names only)
    # Row 18 is 적응 소계
    # Row 19 is first 적응 category (values for all 7 packed)
    # Rows 20-25 are empty

    def extract_values_row(row_idx, col_range):
        """Extract values from a single row."""
        vals = []
        for c in col_range:
            v = parse_number(str(df.iloc[row_idx, c]))
            vals.append(v)
        return vals

    def extract_packed_values(row_idx, col_range, n_categories):
        """Extract packed multi-row values from a merged cell row."""
        all_vals = []  # list of lists: [cat_idx][org_idx]
        for _ in range(n_categories):
            all_vals.append([])

        for c in col_range:
            cell_val = str(df.iloc[row_idx, c])
            parts = cell_val.split('\n')
            for i, part in enumerate(parts):
                if i < n_categories:
                    all_vals[i].append(parse_number(part))

        return all_vals

    col_range = range(2, 9)  # columns for 6 org_sizes + 전체

    # 합계 (row 1)
    vals = extract_values_row(1, col_range)
    for i, org in enumerate(ORG_SIZES):
        if vals[i] is not None:
            results.append((None, org, vals[i]))

    # 감축 categories (packed in row 3, 15 categories)
    gamchuk_vals = extract_packed_values(3, col_range, 15)
    for cat_idx, cat_name in enumerate(CATEGORY_ORDER_GAMCHUK):
        cat_id = CATEGORY_MAP[cat_name]
        if cat_idx < len(gamchuk_vals):
            for i, org in enumerate(ORG_SIZES):
                if i < len(gamchuk_vals[cat_idx]) and gamchuk_vals[cat_idx][i] is not None:
                    results.append((cat_id, org, gamchuk_vals[cat_idx][i]))

    # 적응 categories (packed in row 19, 7 categories)
    jeokeung_vals = extract_packed_values(19, col_range, 7)
    for cat_idx, cat_name in enumerate(CATEGORY_ORDER_JEOKEUNG):
        cat_id = CATEGORY_MAP[cat_name]
        if cat_idx < len(jeokeung_vals):
            for i, org in enumerate(ORG_SIZES):
                if i < len(jeokeung_vals[cat_idx]) and jeokeung_vals[cat_idx][i] is not None:
                    results.append((cat_id, org, jeokeung_vals[cat_idx][i]))

    return results


def load_2023_from_pdf():
    """Load 2023 data from PDF tables."""
    print(f"\nLoading 2023 PDF: {os.path.basename(PDF_2023)}")
    doc = fitz.open(PDF_2023)

    data = {}
    # Page 63: table 0 = 표1 (revenue_climate)
    # Page 65: table 0 = 표5 (R&D org - skip), table 1 = 표6 (rnd_expense_climate)
    # Page 66: table 0 = 표7 (employee_count), table 1 = 표8 (employee_count_climate)
    pdf_tables = [
        (63, 0, 'revenue_climate', '백만원'),
        (65, 1, 'rnd_expense_climate', '백만원'),
        (66, 0, 'employee_count', '명'),
        (66, 1, 'employee_count_climate', '명'),
    ]

    for page_num, table_idx, metric, unit in pdf_tables:
        rows = parse_pdf_table(doc, page_num, table_idx)
        data[(2023, metric)] = [(cat_id, org, val, unit) for cat_id, org, val in rows]
        print(f"  p{page_num} t{table_idx} ({metric}): {len(rows)} data points")

    doc.close()
    return data


# ──────────────────────────────────────────────────────────────────────
# Main: Update DB
# ──────────────────────────────────────────────────────────────────────

def main():
    # Load data from both sources
    data_2022 = load_2022_from_html()
    data_2023 = load_2023_from_pdf()

    # Merge
    all_data = {}
    all_data.update(data_2022)
    all_data.update(data_2023)

    # Connect to DB
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Metrics to fix
    metrics_to_fix = ['revenue_climate', 'rnd_expense_climate', 'employee_count', 'employee_count_climate']

    # Show current state
    print("\n=== Current DB state (to be replaced) ===")
    for mt in metrics_to_fix:
        cur.execute(
            "SELECT survey_year, COUNT(*), SUM(metric_value) FROM activity_survey "
            "WHERE metric_type=? GROUP BY survey_year ORDER BY survey_year", (mt,))
        for year, cnt, total in cur.fetchall():
            print(f"  {year} {mt}: {cnt} rows, sum={total:,.0f}")

    # Delete existing data for these metrics
    for mt in metrics_to_fix:
        cur.execute("DELETE FROM activity_survey WHERE metric_type=?", (mt,))
    deleted = cur.rowcount
    conn.commit()
    print(f"\nDeleted rows for {metrics_to_fix}")

    # Also delete revenue and rnd_expense (전체 매출/R&D) since they're not in report tables
    for mt in ['revenue', 'rnd_expense']:
        cur.execute("DELETE FROM activity_survey WHERE metric_type=?", (mt,))
    conn.commit()
    print("Also deleted revenue and rnd_expense (not in report tables)")

    # Insert new data
    inserts = []
    for (year, metric), rows in sorted(all_data.items()):
        for cat_id, org, val, unit in rows:
            inserts.append((year, cat_id, org, metric, val, unit, 1))

    cur.executemany(
        """INSERT INTO activity_survey
           (survey_year, category_id, org_size, metric_type, metric_value, metric_unit, is_weighted)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        inserts
    )
    conn.commit()
    print(f"\nInserted {len(inserts)} rows")

    # Verify
    print("\n=== Verification ===")
    for mt in metrics_to_fix:
        cur.execute(
            "SELECT survey_year, org_size, SUM(metric_value), COUNT(*) "
            "FROM activity_survey WHERE metric_type=? AND category_id IS NULL "
            "GROUP BY survey_year, org_size ORDER BY survey_year, org_size", (mt,))
        for year, org, total, cnt in cur.fetchall():
            print(f"  {year} {mt} {org}: {total:,.0f} ({cnt} rows)")

    # Specific check: 2022 revenue_climate 합계 전체 should be 223,238,123
    cur.execute(
        "SELECT metric_value FROM activity_survey "
        "WHERE survey_year=2022 AND metric_type='revenue_climate' "
        "AND category_id IS NULL AND org_size='전체'")
    row = cur.fetchone()
    expected = 223_238_123
    if row:
        actual = row[0]
        status = 'OK' if actual == expected else f'MISMATCH (expected {expected:,})'
        print(f"\n  CHECK 2022 revenue_climate 전체: {actual:,.0f} -> {status}")

    # Check: 2022 태양광·열 전체 = 30,404,081
    cur.execute(
        "SELECT metric_value FROM activity_survey "
        "WHERE survey_year=2022 AND metric_type='revenue_climate' "
        "AND category_id=83 AND org_size='전체'")
    row = cur.fetchone()
    expected2 = 30_404_081
    if row:
        actual2 = row[0]
        status2 = 'OK' if actual2 == expected2 else f'MISMATCH (expected {expected2:,})'
        print(f"  CHECK 2022 revenue_climate 태양광·열 전체: {actual2:,.0f} -> {status2}")

    # Check: 2023 revenue_climate 합계 전체 = 237,791,609
    cur.execute(
        "SELECT metric_value FROM activity_survey "
        "WHERE survey_year=2023 AND metric_type='revenue_climate' "
        "AND category_id IS NULL AND org_size='전체'")
    row = cur.fetchone()
    expected3 = 237_791_609
    if row:
        actual3 = row[0]
        status3 = 'OK' if actual3 == expected3 else f'MISMATCH (expected {expected3:,})'
        print(f"  CHECK 2023 revenue_climate 전체: {actual3:,.0f} -> {status3}")

    # Final summary
    print("\n=== Final DB summary ===")
    cur.execute("""
        SELECT survey_year, metric_type, COUNT(*), SUM(metric_value)
        FROM activity_survey
        GROUP BY survey_year, metric_type
        ORDER BY survey_year, metric_type
    """)
    for year, mt, cnt, total in cur.fetchall():
        total_str = f"{total:,.0f}" if total else "N/A"
        print(f"  {year} | {mt:45s} | {cnt:4d} rows | sum={total_str}")

    cur.execute("SELECT COUNT(*) FROM activity_survey")
    print(f"\nTotal rows: {cur.fetchone()[0]}")

    conn.close()
    print("\nDone!")


if __name__ == '__main__':
    main()
