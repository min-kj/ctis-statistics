# 🐜 Scout: PLAN-020 데이터 로더 - 2026-03-20
"""
CTis 기후기술통계 데이터 로더
- 2020년 수준조사 (44대/185개 세부기술)
- 2025년 수준조사 (38대/157개 세부기술)
- 44대↔38대 매핑 테이블

SQLite DB 기반 (fallback: Excel)
"""
import sqlite3
import warnings

import pandas as pd
import streamlit as st
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
MAPPING_DIR = Path(__file__).parent.parent / "mapping"
DB_PATH = Path(__file__).parent.parent / "db" / "ctis_stats.db"

COUNTRIES = ["한국", "중국", "일본", "미국", "EU"]
COUNTRY_CODES = ["kr", "cn", "jp", "us", "eu"]
COUNTRY_COLORS = {
    "한국": "#FF6B6B", "중국": "#4ECDC4", "일본": "#45B7D1",
    "미국": "#96CEB4", "EU": "#FECA57",
}
COUNTRY_FLAGS = {
    "한국": "🇰🇷", "중국": "🇨🇳", "일본": "🇯🇵", "미국": "🇺🇸", "EU": "🇪🇺",
}

# DB country_code (uppercase) → internal code (lowercase)
_DB_CC = {"KR": "kr", "CN": "cn", "JP": "jp", "US": "us", "EU": "eu"}
_CC_DB = {v: k for k, v in _DB_CC.items()}
# DB country_code → Korean display name
_DB_COUNTRY_NAME = {"KR": "한국", "CN": "중국", "JP": "일본", "US": "미국", "EU": "EU"}


def _get_db_connection():
    """Return a sqlite3 connection to the CTis stats DB."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def _db_available():
    """Check if the DB file exists."""
    return DB_PATH.exists()


def _pivot_survey_detail(rows_df, year):
    """
    Pivot long-format survey_result rows (one row per detail×country)
    into wide-format (one row per detail) matching the original Excel layout.

    Returns a DataFrame with columns like:
      type, category, detail, kr_level, kr_gap, kr_group, cn_level, ...
      leading_country, survey_year, taxonomy
    """
    if rows_df.empty:
        return pd.DataFrame()

    # Identify the leading country per detail
    leading = (
        rows_df[rows_df["is_leading"] == 1][["detail_id", "country_code"]]
        .drop_duplicates("detail_id")
        .rename(columns={"country_code": "leading_cc"})
    )

    # Base info (one row per detail)
    base_cols = ["detail_id", "detail_name", "detail_no", "category_name",
                 "type_name", "category_no"]
    base = rows_df[base_cols].drop_duplicates("detail_id")
    base = base.merge(leading, on="detail_id", how="left")
    base["leading_country"] = base["leading_cc"].map(_DB_COUNTRY_NAME).fillna("")

    # Pivot numeric columns per country
    for db_cc, code in _DB_CC.items():
        country_data = rows_df[rows_df["country_code"] == db_cc][
            ["detail_id", "tech_level", "tech_gap", "tech_group",
             "basic_research", "applied_research", "rd_trend"]
        ].rename(columns={
            "tech_level": f"{code}_level",
            "tech_gap": f"{code}_gap",
            "tech_group": f"{code}_group",
            "basic_research": f"{code}_basic_research",
            "applied_research": f"{code}_applied_research",
            "rd_trend": f"{code}_rd_trend",
        })
        base = base.merge(country_data, on="detail_id", how="left")

    # Rename to match original column names
    base = base.rename(columns={
        "type_name": "type",
        "category_name": "category",
        "detail_name": "detail",
        "detail_no": "detail_no",
        "category_no": "cat_no",
    })

    # For 2025: also provide category_38 (cleaned name without number prefix)
    if year == 2025:
        base["category_38_raw"] = base["category"]
        base["category_38"] = base["category"].str.replace(r"^\d+\.\s*", "", regex=True)
        base["taxonomy"] = "38대"
    else:
        base["taxonomy"] = "44대"

    base["survey_year"] = year

    # Ensure numeric columns
    for col in base.columns:
        if any(s in col for s in ["_level", "_gap", "_research"]):
            base[col] = pd.to_numeric(base[col], errors="coerce")

    return base


def _build_cat_stats(conn, year=2025):
    """
    Build 대분류별 통계 DataFrame matching the Excel '대분류별_통계' sheet.
    Columns: 대분류, 세부기술_수, 한국_평균수준, 한국_평균격차, ...
    """
    query = """
    SELECT category_name, country_code, avg_level, avg_gap, detail_count
    FROM v_category_summary
    WHERE survey_year = ?
    """
    df = pd.read_sql_query(query, conn, params=(year,))

    # Pivot to wide format
    categories = df["category_name"].unique()
    rows = []
    for cat in sorted(categories):
        cat_data = df[df["category_name"] == cat]
        row = {"대분류": cat}
        # detail_count is the same across countries for a category
        row["세부기술_수"] = int(cat_data["detail_count"].iloc[0])
        for _, r in cat_data.iterrows():
            name = _DB_COUNTRY_NAME.get(r["country_code"], r["country_code"])
            row[f"{name}_평균수준"] = r["avg_level"]
            row[f"{name}_평균격차"] = r["avg_gap"]
        rows.append(row)

    result = pd.DataFrame(rows)
    return result


def _build_country_stats(conn, year=2025):
    """
    Build 국가별 통계 DataFrame matching the Excel '국가별_통계' sheet.
    Columns: 국가, 평균_기술수준, 중앙_기술수준, 평균_기술격차, 중앙_기술격차, 최고국_횟수, 데이터_존재_수
    """
    # Get averages from view
    query = """
    SELECT country_code, avg_level, avg_gap, tech_count
    FROM v_country_summary
    WHERE survey_year = ?
    """
    summary = pd.read_sql_query(query, conn, params=(year,))

    # Get median and leading count from raw data
    detail_query = """
    SELECT country_code, tech_level, tech_gap, is_leading
    FROM survey_result
    WHERE survey_year = ?
    """
    raw = pd.read_sql_query(detail_query, conn, params=(year,))

    rows = []
    for _, s in summary.iterrows():
        cc = s["country_code"]
        name = _DB_COUNTRY_NAME.get(cc, cc)
        cc_raw = raw[raw["country_code"] == cc]
        rows.append({
            "국가": name,
            "평균_기술수준": round(s["avg_level"], 2),
            "중앙_기술수준": round(cc_raw["tech_level"].median(), 2),
            "평균_기술격차": round(s["avg_gap"], 2),
            "중앙_기술격차": round(cc_raw["tech_gap"].median(), 2),
            "최고국_횟수": int(cc_raw["is_leading"].sum()),
            "데이터_존재_수": int(s["tech_count"]),
        })

    # Order: 한국, 중국, 일본, 미국, EU
    result = pd.DataFrame(rows)
    order = {name: i for i, name in enumerate(COUNTRIES)}
    result["_order"] = result["국가"].map(order)
    result = result.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    return result


def _build_type_stats(conn, year=2025):
    """
    Build 감축/적응 통계 DataFrame matching the Excel '감축적응_통계' sheet.
    Columns: 분류, 세부기술_수, 대분류_수, 한국_평균수준, 한국_평균격차, ...
    """
    # Type summary from view
    query = """
    SELECT type_name, country_code, avg_level, avg_gap, tech_count
    FROM v_type_summary
    WHERE survey_year = ?
    """
    df = pd.read_sql_query(query, conn, params=(year,))

    # Count distinct categories per type
    cat_count_query = """
    SELECT tc.type_name, COUNT(DISTINCT tc.category_id) as cat_count
    FROM tech_category tc
    JOIN tech_detail td ON tc.category_id = td.category_id
    WHERE td.survey_year = ?
    GROUP BY tc.type_name
    """
    cat_counts = pd.read_sql_query(cat_count_query, conn, params=(year,))
    cat_count_map = dict(zip(cat_counts["type_name"], cat_counts["cat_count"]))

    # Leading counts per type × country
    leading_query = """
    SELECT tc.type_name, sr.country_code, SUM(sr.is_leading) as leading_count
    FROM survey_result sr
    JOIN tech_detail td ON sr.detail_id = td.detail_id
    JOIN tech_category tc ON td.category_id = tc.category_id
    WHERE sr.survey_year = ?
    GROUP BY tc.type_name, sr.country_code
    """
    leading_df = pd.read_sql_query(leading_query, conn, params=(year,))

    # Build rows: 전체, 감축, 적응
    type_names = ["전체", "감축", "적응"]
    rows = []
    for tn in type_names:
        row = {"분류": tn}
        if tn == "전체":
            subset = df
            row["세부기술_수"] = int(df.groupby("country_code")["tech_count"].first().iloc[0])
            # sum of all type cat counts
            row["대분류_수"] = sum(cat_count_map.values())
        else:
            subset = df[df["type_name"] == tn]
            if len(subset) == 0:
                continue
            row["세부기술_수"] = int(subset.groupby("country_code")["tech_count"].first().iloc[0])
            row["대분류_수"] = cat_count_map.get(tn, 0)

        for cc_db, name in _DB_COUNTRY_NAME.items():
            if tn == "전체":
                cc_data = df[df["country_code"] == cc_db]
                # Weighted average across types
                total_count = cc_data["tech_count"].sum()
                if total_count > 0:
                    row[f"{name}_평균수준"] = round(
                        (cc_data["avg_level"] * cc_data["tech_count"]).sum() / total_count, 2
                    )
                    row[f"{name}_평균격차"] = round(
                        (cc_data["avg_gap"] * cc_data["tech_count"]).sum() / total_count, 2
                    )
                else:
                    row[f"{name}_평균수준"] = 0
                    row[f"{name}_평균격차"] = 0
            else:
                cc_data = subset[subset["country_code"] == cc_db]
                if len(cc_data) > 0:
                    row[f"{name}_평균수준"] = round(cc_data["avg_level"].iloc[0], 2)
                    row[f"{name}_평균격차"] = round(cc_data["avg_gap"].iloc[0], 2)
                else:
                    row[f"{name}_평균수준"] = 0
                    row[f"{name}_평균격차"] = 0

            # Leading counts
            if tn == "전체":
                lc = leading_df[leading_df["country_code"] == cc_db]["leading_count"].sum()
            else:
                lc_sub = leading_df[
                    (leading_df["type_name"] == tn) & (leading_df["country_code"] == cc_db)
                ]
                lc = int(lc_sub["leading_count"].iloc[0]) if len(lc_sub) > 0 else 0
            row[f"{name}_최고국수"] = int(lc)

        rows.append(row)

    return pd.DataFrame(rows)


@st.cache_data(ttl=3600)
def load_2020_data():
    """2020년 수준조사 세부기술별 데이터 (44대 체계, 185개 세부기술)"""
    if _db_available():
        conn = _get_db_connection()
        try:
            query = """
            SELECT sr.*, td.detail_name, td.detail_no,
                   tc.category_name, tc.type_name, tc.category_no
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = 2020
            """
            rows_df = pd.read_sql_query(query, conn)
            df = _pivot_survey_detail(rows_df, 2020)
            return df
        finally:
            conn.close()

    # Fallback: Excel
    warnings.warn("DB not found, falling back to Excel for load_2020_data")
    path = DATA_DIR / "수준조사 데이터(2020)_250926.xlsx"
    df = pd.read_excel(path, sheet_name="세부기술별_데이터")

    col_map = {
        "대분류": "type",
        "중분류번호": "cat_no",
        "중분류": "category",
        "소분류": "subcategory",
        "세부기술번호 ": "detail_no",
        "세부기술": "detail",
        "최고 기술 보유국": "leading_country",
    }
    for country, code in zip(COUNTRIES, COUNTRY_CODES):
        col_map[f"{country}-기술 수준 (%)"] = f"{code}_level"
        col_map[f"{country}-기술 격차 (년)"] = f"{code}_gap"
        col_map[f"{country}-기술 수준 그룹"] = f"{code}_group"
    for country, code in zip(COUNTRIES, COUNTRY_CODES):
        col_map[f"{country}-연구 개발 활동 경향"] = f"{code}_rd_trend"
        col_map[f"{country}-기초 연구 역량(점)"] = f"{code}_basic_research"
        col_map[f"{country}-응용 개발 연구 역량(점)"] = f"{code}_applied_research"

    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    for col in df.columns:
        if any(s in col for s in ["_level", "_gap", "_research"]):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["survey_year"] = 2020
    df["taxonomy"] = "44대"
    return df


@st.cache_data(ttl=3600)
def load_2025_data():
    """2025년 수준조사 5개국 요약 (38대 체계, 157개 세부기술)

    Returns (df_detail, cat_stats, country_stats, type_stats) tuple.
    """
    if _db_available():
        conn = _get_db_connection()
        try:
            # Detail-level data
            query = """
            SELECT sr.*, td.detail_name, td.detail_no,
                   tc.category_name, tc.type_name, tc.category_no
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = 2025
            """
            rows_df = pd.read_sql_query(query, conn)
            df_detail = _pivot_survey_detail(rows_df, 2025)

            cat_stats = _build_cat_stats(conn, 2025)
            country_stats = _build_country_stats(conn, 2025)
            type_stats = _build_type_stats(conn, 2025)

            return df_detail, cat_stats, country_stats, type_stats
        finally:
            conn.close()

    # Fallback: Excel
    warnings.warn("DB not found, falling back to Excel for load_2025_data")
    path = DATA_DIR / "2차델파이_5개국_요약테이블.xlsx"

    df = pd.read_excel(path, sheet_name="5개국_요약")
    col_map = {
        "대분류": "category_38",
        "세부기술": "detail",
        "최고_기술_보유국": "leading_country",
        "분류": "type",
    }
    for country, code in zip(COUNTRIES, COUNTRY_CODES):
        col_map[f"{country}_기술수준"] = f"{code}_level"
        col_map[f"{country}_기술격차"] = f"{code}_gap"

    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    for col in df.columns:
        if any(s in col for s in ["_level", "_gap"]):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "category_38" in df.columns:
        df["category_38_raw"] = df["category_38"]
        df["category_38"] = df["category_38"].str.replace(
            r"^\d+\.\s*", "", regex=True
        )

    df["survey_year"] = 2025
    df["taxonomy"] = "38대"

    cat_stats = pd.read_excel(path, sheet_name="대분류별_통계")
    country_stats = pd.read_excel(path, sheet_name="국가별_통계")
    type_stats = pd.read_excel(path, sheet_name="감축적응_통계")

    return df, cat_stats, country_stats, type_stats


@st.cache_data(ttl=3600)
def load_2020_category_summary():
    """2020년 소분류(44대)별 집계 데이터"""
    if _db_available():
        conn = _get_db_connection()
        try:
            # Use v_category_summary for 2020
            query = """
            SELECT category_name, type_name, country_code,
                   avg_level, avg_gap, detail_count, leading_count
            FROM v_category_summary
            WHERE survey_year = 2020
            """
            df = pd.read_sql_query(query, conn)

            # Also get per-category aggregated research/group/rd_trend from survey_result
            extra_query = """
            SELECT tc.category_name, sr.country_code,
                   GROUP_CONCAT(DISTINCT sr.tech_group) as tech_group,
                   AVG(sr.basic_research) as basic_research,
                   AVG(sr.applied_research) as applied_research,
                   GROUP_CONCAT(DISTINCT sr.rd_trend) as rd_trend
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = 2020
            GROUP BY tc.category_name, sr.country_code
            """
            extra = pd.read_sql_query(extra_query, conn)

            # Merge
            df = df.merge(extra, on=["category_name", "country_code"], how="left")

            # Find leading country per category (most is_leading)
            leading_query = """
            SELECT tc.category_name, sr.country_code,
                   SUM(sr.is_leading) as lc
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = 2020 AND sr.is_leading = 1
            GROUP BY tc.category_name, sr.country_code
            ORDER BY tc.category_name, lc DESC
            """
            leading_df = pd.read_sql_query(leading_query, conn)
            # Pick the country with most leading techs per category
            leading_map = {}
            for _, row in leading_df.iterrows():
                cat = row["category_name"]
                if cat not in leading_map:
                    leading_map[cat] = _DB_COUNTRY_NAME.get(row["country_code"], row["country_code"])

            # Pivot to wide format (one row per category)
            categories = df["category_name"].unique()
            rows = []
            for cat in sorted(categories):
                cat_data = df[df["category_name"] == cat]
                row_dict = {
                    "type": cat_data["type_name"].iloc[0],
                    "category": cat,
                    "leading_country": leading_map.get(cat, ""),
                    "detail_count": int(cat_data["detail_count"].iloc[0]),
                }
                for _, r in cat_data.iterrows():
                    code = _DB_CC.get(r["country_code"], r["country_code"].lower())
                    row_dict[f"{code}_level"] = r["avg_level"]
                    row_dict[f"{code}_gap"] = r["avg_gap"]
                    row_dict[f"{code}_group"] = r.get("tech_group", "")
                    row_dict[f"{code}_basic_research"] = r.get("basic_research")
                    row_dict[f"{code}_applied_research"] = r.get("applied_research")
                    row_dict[f"{code}_rd_trend"] = r.get("rd_trend", "")
                rows.append(row_dict)

            result = pd.DataFrame(rows)

            # Ensure numeric columns
            for col in result.columns:
                if any(s in col for s in ["_level", "_gap", "_research"]):
                    result[col] = pd.to_numeric(result[col], errors="coerce")

            return result
        finally:
            conn.close()

    # Fallback: Excel
    warnings.warn("DB not found, falling back to Excel for load_2020_category_summary")
    path = DATA_DIR / "수준조사 데이터(2020)_250926.xlsx"
    df = pd.read_excel(path, sheet_name="소분류별_집계데이터")

    col_map = {
        "대분류": "type",
        "소분류": "category",
        "최고기술보유국": "leading_country",
        "세부기술수": "detail_count",
    }
    for country, code in zip(COUNTRIES, COUNTRY_CODES):
        col_map[f"{country}_기술수준"] = f"{code}_level"
        col_map[f"{country}_기술격차"] = f"{code}_gap"
        col_map[f"{country}_기술그룹"] = f"{code}_group"
    for code in COUNTRY_CODES:
        col_map[f"{code}_rd_trend"] = f"{code}_rd_trend"
        for metric in ["기초연구역량", "응용연구역량"]:
            for cn, cc in zip(COUNTRIES, COUNTRY_CODES):
                col_map[f"{cn}_{metric}"] = f"{cc}_{'basic' if '기초' in metric else 'applied'}_research"

    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    for col in df.columns:
        if any(s in col for s in ["_level", "_gap", "_research"]):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@st.cache_data(ttl=3600)
def load_mapping_44_to_38():
    """44대(2020)↔38대(2025) 세부기술 매핑 테이블"""
    if _db_available():
        conn = _get_db_connection()
        try:
            query = """
            SELECT
                td_src.detail_name AS detail_2020,
                sc.category_name   AS category_44,
                tc.category_name   AS category_38
            FROM taxonomy_mapping tm
            JOIN tech_category sc  ON tm.src_category_id = sc.category_id
            JOIN tech_category tc  ON tm.tgt_category_id = tc.category_id
            LEFT JOIN tech_detail td_src ON td_src.category_id = sc.category_id
            WHERE td_src.survey_year = 2020
            ORDER BY sc.category_name, td_src.detail_name
            """
            df = pd.read_sql_query(query, conn)
            return df
        finally:
            conn.close()

    # Fallback: Excel
    warnings.warn("DB not found, falling back to Excel for load_mapping_44_to_38")
    path = DATA_DIR / "기술연계표_맵핑표.xlsx"
    df = pd.read_excel(path, sheet_name="2020 세부185-44")
    df.columns = ["detail_2020", "category_44", "category_38"]
    return df


@st.cache_data(ttl=3600)
def load_timeseries_mapping():
    """시계열 맵핑표 — 44대↔38대 중분류 레벨 매핑 + 매핑 유형"""
    if _db_available():
        conn = _get_db_connection()
        try:
            query = """
            SELECT
                tm.mapping_type,
                sc.category_name   AS category_44,
                td_src.detail_name AS detail_2020,
                tc.category_name   AS category_38,
                td_tgt.detail_name AS detail_2025
            FROM taxonomy_mapping tm
            JOIN tech_category sc  ON tm.src_category_id = sc.category_id
            JOIN tech_category tc  ON tm.tgt_category_id = tc.category_id
            LEFT JOIN tech_detail td_src ON td_src.category_id = sc.category_id
                                         AND td_src.survey_year = 2020
            LEFT JOIN tech_detail td_tgt ON td_tgt.category_id = tc.category_id
                                         AND td_tgt.survey_year = 2025
            ORDER BY sc.category_name
            """
            df = pd.read_sql_query(query, conn)
            # Remove rows where both categories are null
            df = df.dropna(subset=["category_44", "category_38"], how="all")
            return df
        finally:
            conn.close()

    # Fallback: Excel
    warnings.warn("DB not found, falling back to Excel for load_timeseries_mapping")
    path = MAPPING_DIR / "2020-2025 시계열 맵핑표(검토요청).xlsx"
    df = pd.read_excel(path, sheet_name="맵핑표(38-45)", header=0)

    result = pd.DataFrame({
        "mapping_type": df.iloc[:, 0],
        "category_44": df.iloc[:, 1],
        "detail_2020": df.iloc[:, 3],
        "category_38": df.iloc[:, 5],
        "detail_2025": df.iloc[:, 7],
    })
    result = result.dropna(subset=["category_44", "category_38"], how="all")
    result = result[result["mapping_type"].apply(
        lambda x: isinstance(x, str) if pd.notna(x) else False
    )]
    return result


@st.cache_data(ttl=3600)
def load_timeseries_data():
    """Load mapped timeseries data for comparison between survey years.

    Dynamically determines which taxonomies were used for each survey year,
    finds mappings via taxonomy_mapping, filters to only mapped technologies,
    and normalizes gap units to months.

    Returns (df_old, df_new, mapping_info) where:
    - df_old: detail-level data for the older survey year (gap normalized to months)
    - df_new: detail-level data for the newer survey year (gap in months)
    - mapping_info: dict with metadata
    """
    if not _db_available():
        return pd.DataFrame(), pd.DataFrame(), {
            "old_year": None, "new_year": None,
            "old_taxonomy": None, "new_taxonomy": None,
            "mapped_count": 0, "unmapped_old": 0, "unmapped_new": 0,
        }

    conn = sqlite3.connect(str(DB_PATH))
    try:
        # 1. Find available survey years and their taxonomies
        year_tax = pd.read_sql_query("""
            SELECT DISTINCT sr.survey_year, t.code AS taxonomy_code,
                   t.taxonomy_id, sr.gap_unit
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            JOIN taxonomy t ON tc.taxonomy_id = t.taxonomy_id
            GROUP BY sr.survey_year
            ORDER BY sr.survey_year
        """, conn)

        if len(year_tax) < 2:
            return pd.DataFrame(), pd.DataFrame(), {
                "old_year": None, "new_year": None,
                "old_taxonomy": None, "new_taxonomy": None,
                "mapped_count": 0, "unmapped_old": 0, "unmapped_new": 0,
            }

        old_row = year_tax.iloc[0]
        new_row = year_tax.iloc[-1]
        old_year = int(old_row["survey_year"])
        new_year = int(new_row["survey_year"])
        old_tax_id = int(old_row["taxonomy_id"])
        new_tax_id = int(new_row["taxonomy_id"])
        old_tax_code = old_row["taxonomy_code"]
        new_tax_code = new_row["taxonomy_code"]
        old_gap_unit = old_row["gap_unit"]
        new_gap_unit = new_row["gap_unit"]

        # 2. Get taxonomy_mapping between the two taxonomies (category-level)
        mapping = pd.read_sql_query("""
            SELECT tm.src_category_id, tm.tgt_category_id, tm.mapping_type
            FROM taxonomy_mapping tm
            JOIN tech_category sc ON tm.src_category_id = sc.category_id
            JOIN tech_category tc ON tm.tgt_category_id = tc.category_id
            WHERE sc.taxonomy_id = ? AND tc.taxonomy_id = ?
        """, conn, params=(old_tax_id, new_tax_id))

        mapped_old_cat_ids = set(mapping["src_category_id"].tolist())
        mapped_new_cat_ids = set(mapping["tgt_category_id"].tolist())

        # 3. Load detail-level data for old year, filtered to mapped categories
        df_old = pd.read_sql_query("""
            SELECT sr.detail_id, sr.country_code, sr.tech_level, sr.tech_gap,
                   sr.gap_unit, sr.is_leading,
                   td.detail_name AS detail, td.category_id,
                   tc.category_name AS category, tc.type_name AS type
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = ?
        """, conn, params=(old_year,))

        df_new = pd.read_sql_query("""
            SELECT sr.detail_id, sr.country_code, sr.tech_level, sr.tech_gap,
                   sr.gap_unit, sr.is_leading,
                   td.detail_name AS detail, td.category_id,
                   tc.category_name AS category, tc.type_name AS type
            FROM survey_result sr
            JOIN tech_detail td ON sr.detail_id = td.detail_id
            JOIN tech_category tc ON td.category_id = tc.category_id
            WHERE sr.survey_year = ?
        """, conn, params=(new_year,))

        # Count total categories per taxonomy (for unmapped stats)
        total_old_cats = conn.execute(
            "SELECT COUNT(*) FROM tech_category WHERE taxonomy_id = ?",
            (old_tax_id,)
        ).fetchone()[0]
        total_new_cats = conn.execute(
            "SELECT COUNT(*) FROM tech_category WHERE taxonomy_id = ?",
            (new_tax_id,)
        ).fetchone()[0]

        # Filter to only mapped categories
        df_old_mapped = df_old[df_old["category_id"].isin(mapped_old_cat_ids)].copy()
        df_new_mapped = df_new[df_new["category_id"].isin(mapped_new_cat_ids)].copy()

        # 4. Normalize gap units to years (년)
        #    2020: 원본 year → 그대로
        #    2025: 원본 month → / 12 변환
        #    산식: 기술격차(년) = 기술격차(개월) / 12
        if new_gap_unit == "month":
            df_new_mapped["tech_gap"] = df_new_mapped["tech_gap"] / 12
            df_new_mapped["gap_unit"] = "year"

        # Pivot to wide format for each country
        def _pivot_to_wide(df_long):
            """Pivot long-format (one row per detail x country) to wide."""
            if df_long.empty:
                return pd.DataFrame()

            base = df_long[["detail_id", "detail", "category_id", "category", "type"]].drop_duplicates("detail_id")

            # Leading country
            leading = (
                df_long[df_long["is_leading"] == 1][["detail_id", "country_code"]]
                .drop_duplicates("detail_id")
            )
            leading["leading_country"] = leading["country_code"].map(_DB_COUNTRY_NAME).fillna("")
            base = base.merge(leading[["detail_id", "leading_country"]], on="detail_id", how="left")
            base["leading_country"] = base["leading_country"].fillna("")

            for db_cc, code in _DB_CC.items():
                cc_data = df_long[df_long["country_code"] == db_cc][
                    ["detail_id", "tech_level", "tech_gap"]
                ].rename(columns={
                    "tech_level": f"{code}_level",
                    "tech_gap": f"{code}_gap",
                })
                base = base.merge(cc_data, on="detail_id", how="left")

            for col in base.columns:
                if any(s in col for s in ["_level", "_gap"]):
                    base[col] = pd.to_numeric(base[col], errors="coerce")

            return base

        df_old_wide = _pivot_to_wide(df_old_mapped)
        df_new_wide = _pivot_to_wide(df_new_mapped)

        mapping_info = {
            "old_year": old_year,
            "new_year": new_year,
            "old_taxonomy": old_tax_code,
            "new_taxonomy": new_tax_code,
            "old_category_count": total_old_cats,
            "new_category_count": total_new_cats,
            "mapped_count": len(mapping),
            "unmapped_old": total_old_cats - len(mapped_old_cat_ids),
            "unmapped_new": total_new_cats - len(mapped_new_cat_ids),
            "old_gap_unit_original": old_gap_unit,
            "new_gap_unit_original": new_gap_unit,
            "gap_unit_normalized": "year",
        }

        return df_old_wide, df_new_wide, mapping_info
    finally:
        conn.close()


def get_country_averages(df, level_suffix="_level", gap_suffix="_gap"):
    """국가별 평균 기술수준/격차 계산"""
    result = {}
    for country, code in zip(COUNTRIES, COUNTRY_CODES):
        level_col = f"{code}{level_suffix}"
        gap_col = f"{code}{gap_suffix}"
        if level_col in df.columns:
            result[country] = {
                "level": df[level_col].mean(),
                "gap": df[gap_col].mean() if gap_col in df.columns else None,
            }
    return result


def aggregate_by_category(df, category_col="category", agg_cols=None):
    """중분류별 집계 (세부기술 → 중분류 평균)"""
    if agg_cols is None:
        agg_cols = [c for c in df.columns if any(s in c for s in ["_level", "_gap"])]

    agg_dict = {col: "mean" for col in agg_cols}
    agg_dict["detail"] = "count"
    if "type" in df.columns:
        agg_dict["type"] = "first"
    if "leading_country" in df.columns:
        agg_dict["leading_country"] = lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else "N/A"

    result = df.groupby(category_col).agg(agg_dict).reset_index()
    result = result.rename(columns={"detail": "detail_count"})
    return result
