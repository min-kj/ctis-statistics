# 🐜 Scout: 기술개발활동조사 페이지 - 2026-03-22
"""
CTis 기술개발활동조사 대시보드
- 22대 기후기술 분류 기준 주요 현황
- 2022/2023 시계열 비교
- 감축/적응 비중 분석
"""
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import sqlite3
import io
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))
from data_loader import _get_db_connection, DB_PATH


def _excel_download(df, filename, button_label="📥 Excel 다운로드", key=None):
    """Create Excel download button from DataFrame."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='데이터')
    st.download_button(
        label=button_label,
        data=buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )

# ── 22대 기후기술 색상 팔레트 ──
TECH_AREA_COLORS = [
    '#F26522', '#F7941D', '#FBB03B', '#FCE053', '#8DC63F',
    '#39B54A', '#00A651', '#00A99D', '#00BCD4', '#29ABE2',
    '#1E88E5', '#0066B3', '#0D47A1', '#1A237E', '#6639B7',
    '#8E24AA', '#AB47BC', '#9C27B0', '#C2185B', '#AD1457',
    '#880E4F', '#B73377',
]

TYPE_COLORS = {"감축": "#4a7c23", "적응": "#1E88E5"}

# 보고서 외부 링크
REPORT_LINKS = [
    ("2021 기후기술 산업통계", "https://nigt.re.kr/gtck/annualall.do?mode=view&articleNo=3097&article.offset=0&articleLimit=10&srCategoryId=26"),
    ("2022 기술개발 활동조사", "https://nigt.re.kr/gtck/annualall.do?mode=view&articleNo=4068&article.offset=0&articleLimit=10&srCategoryId=26"),
    ("2023 기술개발 활동조사", "https://www.msit.go.kr/bbs/view.do?sCode=user&mId=99&mPid=74&bbsSeqNo=79&nttSeqNo=3173692"),
]


def render():
    """기술개발활동조사 메인 렌더."""
    _render_overview()

    # 2단 메뉴: 기술개발 활동 규모 / 기술개발 활동 심층
    tab_scale, tab_deep = st.tabs(["기술개발 활동 규모", "기술개발 활동 심층"])

    with tab_scale:
        _render_tab_scale()

    with tab_deep:
        _render_tab_deep()


def _render_tab_scale():
    """C-1. 기술개발 활동 규모 — 매출/R&D/인력 (정량)"""

    col_y, col_t = st.columns(2)
    with col_y:
        year_option = st.selectbox(
            "분석 범위",
            ["2023년", "2022년", "시계열 비교 (2022 vs 2023)"],
            key="scale_year",
        )
    with col_t:
        type_filter = st.selectbox(
            "기술영역", ["전체", "감축", "적응"], key="scale_type",
        )
    if "시계열" in year_option:
        _render_timeseries()
    else:
        year = int(year_option.replace("년", ""))
        _render_single_year(year, type_filter)

    _render_evaluation(year_option)

    st.caption("※ 본 통계는 표본조사(가중치 적용) 결과이며, 모집단 추정치입니다. 세부 기술영역별 수치는 표본 크기에 따라 오차가 클 수 있습니다.")


def _render_tab_deep():
    """C-2. 기술개발 활동 심층 — B/C/D 섹션 (정성)"""

    col_y2, col_t2 = st.columns(2)
    with col_y2:
        year_deep = st.selectbox(
            "분석 범위",
            ["2023년", "2022년"],
            key="deep_year",
        )
    with col_t2:
        type_deep = st.selectbox(
            "기술영역", ["전체", "감축", "적응"], key="deep_type",
        )

    year = int(year_deep.replace("년", ""))

    _render_section_b(year, type_deep)
    st.markdown("---")
    _render_section_c(year, type_deep)
    st.markdown("---")
    _render_section_d(year, type_deep)

    st.caption("※ 본 통계는 표본조사(가중치 적용) 결과이며, 모집단 추정치입니다.")


# ────────────────────────────────────────────
# 통계 개요
# ────────────────────────────────────────────
def _render_overview():
    with st.expander("통계 개요", expanded=False):
        st.markdown("""
| 항목 | 내용 |
|------|------|
| **조사명** | 기후기술개발 활동조사 |
| **승인통계 번호** | 제440001호 |
| **작성기관** | 한국과학기술연구원 NIGT 데이터정보센터 |
| **조사 목적** | 국가 기후기술 R&D 및 사업화 촉진 정책수립 기초자료 |
| **통계 분야** | 과학·기술/환경 |
| **조사 대상** | 전국 기업체 약 2,752개 내외 |
| **분류체계** | 22대 기후기술 (감축 15 + 적응 7) |
| **조사 방식** | 표본조사 (가중치 적용) |
| **공표 주기** | 1년 (작성기준년 익익년 2월) |
| **모집단** | 2022: 13,574개소 / 2023: 13,926개소 |
| **조사 주체** | 국가녹색기술연구소 (NIGT) |
        """)

    st.markdown("##### 결과보고서")
    cols = st.columns(len(REPORT_LINKS))
    for col, (label, url) in zip(cols, REPORT_LINKS):
        with col:
            st.link_button(f"📄 {label}", url)



# ────────────────────────────────────────────
# 단일 연도 뷰
# ────────────────────────────────────────────
def _render_single_year(year: int, type_filter: str = "전체"):
    st.markdown(f"## {year}년 기술개발활동 현황")
    _render_section_a(year, type_filter)


def _load_section_a_data(year: int) -> pd.DataFrame:
    """Load revenue/rnd/employee data for a given year."""
    conn = _get_db_connection()
    query = """
        SELECT tc.category_name, tc.type_name, a.org_size,
               a.metric_type, a.metric_value, a.metric_unit
        FROM activity_survey a
        JOIN tech_category tc ON a.category_id = tc.category_id
        WHERE a.survey_year = ?
          AND a.metric_type IN (
              'revenue', 'revenue_climate',
              'rnd_expense', 'rnd_expense_climate',
              'employee_count', 'employee_count_climate'
          )
    """
    df = pd.read_sql_query(query, conn, params=(year,))
    conn.close()
    return df


def _load_total_data(year: int) -> pd.DataFrame:
    """Load total (category_id IS NULL) summary data."""
    conn = _get_db_connection()
    query = """
        SELECT org_size, metric_type, metric_value, metric_unit
        FROM activity_survey
        WHERE survey_year = ?
          AND category_id IS NULL
          AND metric_type IN (
              'revenue', 'revenue_climate',
              'rnd_expense', 'rnd_expense_climate',
              'employee_count', 'employee_count_climate'
          )
    """
    df = pd.read_sql_query(query, conn, params=(year,))
    conn.close()
    return df


def _millions_to_display(value_million: float) -> tuple[float, str]:
    """Convert million-won to readable unit. Returns (converted_value, unit_str)."""
    if abs(value_million) >= 1_000_000:
        return value_million / 1_000_000, "조원"
    elif abs(value_million) >= 100:
        return value_million / 100, "억원"
    else:
        return value_million, "백만원"


def _format_money(value_million: float) -> str:
    """Format million-won value for display."""
    val, unit = _millions_to_display(value_million)
    if val >= 100:
        return f"{val:,.0f}{unit}"
    elif val >= 10:
        return f"{val:,.1f}{unit}"
    else:
        return f"{val:,.2f}{unit}"


def _render_section_a(year: int, type_filter: str = "전체"):
    """섹션 A: 주요 현황."""
    st.markdown("### 주요 현황")

    df = _load_section_a_data(year)
    df_total = _load_total_data(year)

    if df.empty:
        st.warning(f"{year}년 주요 현황 데이터가 없습니다.")
        return

    # Apply type filter
    if type_filter != "전체":
        df = df[df["type_name"] == type_filter]

    if df.empty:
        st.warning(f"선택한 필터 조건에 해당하는 데이터가 없습니다.")
        return

    # ── 핵심지표 카드 ──
    _render_key_metrics(df, df_total, year)

    st.markdown("")

    # ── Row 1: 매출액 & R&D 투자 ──
    col_rev, col_rnd = st.columns(2)

    with col_rev:
        st.markdown("#### 기후기술 매출액")
        _render_donut_and_bar(df, df_total, "revenue_climate", "revenue_climate", "매출액")

    with col_rnd:
        st.markdown("#### 기후기술 R&D 투자")
        _render_donut_and_bar(df, df_total, "rnd_expense_climate", "rnd_expense_climate", "R&D 투자")

    # ── Row 2: 종사자 수 ──
    col_emp, col_emp_c = st.columns(2)

    with col_emp:
        st.markdown("#### 전체 종사자 수")
        _render_employee_bar(df, "employee_count")

    with col_emp_c:
        st.markdown("#### 기후기술분야 종사자 수")
        _render_employee_bar(df, "employee_count_climate")

    # Excel download - Section A 핵심지표 데이터
    tot = df_total[df_total["org_size"] == "전체"]
    metric_rows = []
    for mt in ["revenue_climate", "rnd_expense_climate",
                "employee_count", "employee_count_climate"]:
        row = tot[tot["metric_type"] == mt]
        val = row["metric_value"].iloc[0] if len(row) > 0 else 0
        unit = row["metric_unit"].iloc[0] if len(row) > 0 else ""
        metric_rows.append({"지표": mt, "값": val, "단위": unit})
    dl_metrics = pd.DataFrame(metric_rows)
    _excel_download(dl_metrics, f"활동조사_핵심지표_{year}.xlsx", key=f"dl_activity_metrics_{year}")


def _render_key_metrics(df: pd.DataFrame, df_total: pd.DataFrame, year: int):
    """핵심지표 카드 3개 렌더 (기후기술 매출/R&D/종사자)."""
    tot = df_total[df_total["org_size"] == "전체"]

    def _get_total(metric: str) -> float:
        row = tot[tot["metric_type"] == metric]
        return row["metric_value"].iloc[0] if len(row) > 0 else 0

    total_rev_c = _get_total("revenue_climate")
    total_rnd_c = _get_total("rnd_expense_climate")
    total_emp = _get_total("employee_count")
    total_emp_c = _get_total("employee_count_climate")

    c1, c2, c3 = st.columns(3)
    cards = [
        (c1, "기후기술 매출액", _format_money(total_rev_c), f"{year}년 기준"),
        (c2, "기후기술 R&D 투자", _format_money(total_rnd_c), f"{year}년 기준"),
        (c3, "기후기술 종사자", f"{total_emp_c:,.0f}명", f"전체 {total_emp:,.0f}명 중"),
    ]
    for col, title, value, sub in cards:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <h3>{title}</h3>
                <div class="value">{value}</div>
                <div class="sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)


def _render_donut_and_bar(df: pd.DataFrame, df_total: pd.DataFrame,
                          climate_metric: str, total_metric: str, label: str):
    """도넛(기후기술 분야별) + 바(규모별) 차트 렌더."""
    # ── Donut: 기후기술 분야별 (감축/적응 내 22대) ──
    climate_data = df[
        (df["metric_type"] == climate_metric) & (df["org_size"] == "전체")
    ].copy()

    if not climate_data.empty:
        climate_data = climate_data.sort_values("type_name")
        # Assign colors
        color_map = {}
        for i, name in enumerate(climate_data["category_name"].unique()):
            color_map[name] = TECH_AREA_COLORS[i % len(TECH_AREA_COLORS)]

        colors = [color_map[n] for n in climate_data["category_name"]]

        fig_donut = go.Figure(data=[go.Pie(
            labels=climate_data["category_name"],
            values=climate_data["metric_value"],
            hole=0.5,
            marker=dict(colors=colors),
            textinfo="label+percent",
            textposition="outside",
            textfont=dict(size=10),
            hovertemplate="%{label}<br>%{value:,.0f}백만원<br>%{percent}<extra></extra>",
        )])
        fig_donut.update_layout(
            title=dict(text=f"기후기술 {label} 분야별 비중", font=dict(size=14)),
            height=400,
            margin=dict(l=10, r=10, t=50, b=10),
            showlegend=False,
        )
        st.plotly_chart(fig_donut, key=f"donut_{climate_metric}", width="stretch")

    # ── Bar: 규모별 전체 vs 기후기술 ──
    size_order = ["100억미만", "100~500억", "500~1000억", "1000~2000억", "2000억이상", "공공기관등"]

    bar_total = df_total[
        (df_total["metric_type"] == total_metric) & (df_total["org_size"] != "전체")
    ].copy()
    bar_climate = df_total[
        (df_total["metric_type"] == climate_metric) & (df_total["org_size"] != "전체")
    ].copy()

    if not bar_total.empty:
        bar_total["org_size"] = pd.Categorical(bar_total["org_size"], categories=size_order, ordered=True)
        bar_total = bar_total.sort_values("org_size")
        bar_climate["org_size"] = pd.Categorical(bar_climate["org_size"], categories=size_order, ordered=True)
        bar_climate = bar_climate.sort_values("org_size")

        fig_bar = go.Figure()
        # Total
        vals_t = bar_total["metric_value"].values / 100  # 억원
        fig_bar.add_trace(go.Bar(
            name=f"전체 {label}",
            x=bar_total["org_size"].astype(str),
            y=vals_t,
            text=[f"{v:,.0f}" for v in vals_t],
            textposition="outside",
            textfont=dict(size=10),
            marker_color="rgba(74,124,35,0.3)",
            hovertemplate="%{x}<br>전체: %{y:,.0f}억원<extra></extra>",
        ))
        # Climate
        vals_c = bar_climate["metric_value"].values / 100  # 억원
        fig_bar.add_trace(go.Bar(
            name=f"기후기술 {label}",
            x=bar_climate["org_size"].astype(str),
            y=vals_c,
            text=[f"{v:,.0f}" for v in vals_c],
            textposition="outside",
            textfont=dict(size=10),
            marker_color="#4a7c23",
            hovertemplate="%{x}<br>기후기술: %{y:,.0f}억원<extra></extra>",
        ))

        fig_bar.update_layout(
            title=dict(text=f"{label} 규모별 분포 (전체 vs 기후기술)", font=dict(size=14)),
            barmode="group",
            height=350,
            xaxis=dict(title="기업 규모"),
            yaxis=dict(title="억원"),
            margin=dict(l=50, r=10, t=50, b=50),
            legend=dict(orientation="h", y=1.12),
        )
        st.plotly_chart(fig_bar, key=f"bar_{total_metric}", width="stretch")


def _render_employee_bar(df: pd.DataFrame, metric_type: str):
    """종사자 수 수평 바차트 (22대 기술별)."""
    emp_data = df[
        (df["metric_type"] == metric_type) & (df["org_size"] == "전체")
    ].copy()

    if emp_data.empty:
        st.info("데이터가 없습니다.")
        return

    emp_data = emp_data.sort_values("metric_value", ascending=True)

    # Assign colors by type
    colors = [TYPE_COLORS.get(t, "#999") for t in emp_data["type_name"]]

    fig = go.Figure(data=[go.Bar(
        y=emp_data["category_name"],
        x=emp_data["metric_value"],
        orientation="h",
        marker_color=colors,
        text=[f"{v:,.0f}명" for v in emp_data["metric_value"]],
        textposition="outside",
        textfont=dict(size=10),
        hovertemplate="%{y}<br>%{x:,.0f}명<extra></extra>",
    )])

    is_climate = "기후기술" in metric_type or "climate" in metric_type
    title = "기후기술분야 종사자 수 (22대 기술별)" if is_climate else "전체 종사자 수 (22대 기술별)"

    fig.update_layout(
        title=dict(text=title, font=dict(size=14)),
        height=max(400, len(emp_data) * 25 + 80),
        xaxis=dict(title="명"),
        margin=dict(l=120, r=60, t=50, b=30),
    )
    st.plotly_chart(fig, key=f"emp_{metric_type}", width="stretch")


# ────────────────────────────────────────────
# B/C/D 공통 데이터 로드
# ────────────────────────────────────────────
def _load_detail(section: str, year: int, type_filter: str | None = None) -> pd.DataFrame:
    """activity_survey_detail 테이블에서 특정 섹션/연도 데이터를 로드."""
    conn = _get_db_connection()
    query = """
        SELECT section, category_code, category_name, type_name,
               sample_size, item_name, item_value, item_unit
        FROM activity_survey_detail
        WHERE section = ? AND survey_year = ?
    """
    params: list = [section, year]
    if type_filter:
        query += " AND type_name = ?"
        params.append(type_filter)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def _detail_to_download_df(section: str, year: int) -> pd.DataFrame:
    """섹션 전체 데이터를 피벗하여 다운로드용 DataFrame 생성."""
    conn = _get_db_connection()
    df = pd.read_sql_query(
        """SELECT survey_year, section, type_name, category_code, category_name,
                  sample_size, item_name, item_value, item_unit
           FROM activity_survey_detail
           WHERE section = ? AND survey_year = ?
           ORDER BY category_code, item_name""",
        conn, params=(section, year),
    )
    conn.close()
    return df


def _render_lifecycle_chart(section: str, title: str, year: int):
    """수명주기 100% 가로누적바 차트 (b1 또는 b2)."""
    st.markdown(f"#### {title}")
    df_lc = _load_detail(section, year)
    if df_lc.empty:
        st.info(f"{section} 데이터가 없습니다.")
        return

    sub = df_lc[df_lc["category_code"].isin(["감축", "적응"])].copy()
    if sub.empty:
        st.info("감축/적응 소계 데이터가 없습니다.")
        return

    lifecycle_items = ["기술도입기", "기술성장기", "기술성숙기", "기술개발기", "기술쇠퇴기"]
    lc_colors = ["#66BB6A", "#43A047", "#2E7D32", "#FFA726", "#EF5350"]
    pivot = sub.pivot_table(
        index="category_name", columns="item_name",
        values="item_value", aggfunc="first",
    ).reindex(columns=lifecycle_items).fillna(0)

    fig = go.Figure()
    for i, item in enumerate(lifecycle_items):
        if item in pivot.columns:
            fig.add_trace(go.Bar(
                name=item,
                y=pivot.index,
                x=pivot[item],
                orientation="h",
                marker_color=lc_colors[i % len(lc_colors)],
                text=[f"{v:.1f}%" for v in pivot[item]],
                textposition="inside",
                textfont=dict(size=10, color="white"),
                hovertemplate="%{y}<br>" + item + ": %{x:.1f}%<extra></extra>",
            ))
    fig.update_layout(
        barmode="stack",
        height=200,
        margin=dict(l=80, r=10, t=10, b=30),
        xaxis=dict(title="%", range=[0, 105]),
        legend=dict(orientation="h", y=-0.3, font=dict(size=10)),
    )
    st.plotly_chart(fig, key=f"{section}_stack_{year}", use_container_width=True)

    with st.expander(f"22대 기술분야별 수명주기 ({title})"):
        detail = df_lc[~df_lc["category_code"].isin(["00", "감축", "적응"])].copy()
        pivot2 = detail.pivot_table(
            index=["type_name", "category_name"], columns="item_name",
            values="item_value", aggfunc="first",
        ).reindex(columns=lifecycle_items).fillna(0)
        pivot2 = pivot2.sort_index()

        labels = [f"[{t}] {c}" for t, c in pivot2.index]
        fig2 = go.Figure()
        for i, item in enumerate(lifecycle_items):
            if item in pivot2.columns:
                fig2.add_trace(go.Bar(
                    name=item,
                    y=labels,
                    x=pivot2[item].values,
                    orientation="h",
                    marker_color=lc_colors[i % len(lc_colors)],
                    hovertemplate="%{y}<br>" + item + ": %{x:.1f}%<extra></extra>",
                ))
        fig2.update_layout(
            barmode="stack",
            height=max(500, len(labels) * 24),
            margin=dict(l=180, r=10, t=10, b=30),
            xaxis=dict(title="%", range=[0, 105]),
            legend=dict(orientation="h", y=-0.08, font=dict(size=10)),
        )
        st.plotly_chart(fig2, key=f"{section}_detail_{year}", use_container_width=True)


# ────────────────────────────────────────────
# 섹션 B: 기술역량 (b1~b4)
# ────────────────────────────────────────────
def _render_section_b(year: int, type_filter: str = "전체"):
    st.markdown("### 기술역량")

    # ── Row 1: b1/b2 기술수명주기 + b3/b4 ──
    col_left, col_right = st.columns(2)

    with col_left:
        # 주력기술(1순위) 수명주기 — section b1
        _render_lifecycle_chart("b1", "주력기술(1순위) 수명주기", year)

        # 부차기술(2순위) 수명주기 — section b2
        _render_lifecycle_chart("b2", "부차기술(2순위) 수명주기", year)

    with col_right:
        # b3: R&D 전담조직
        st.markdown("#### R&D 전담조직 보유율")
        df_b3 = _load_detail("b3", year)
        if not df_b3.empty:
            sub3 = df_b3[df_b3["category_code"].isin(["00", "감축", "적응"]) & (df_b3["item_name"] == "있음")]
            if not sub3.empty:
                fig3 = go.Figure(go.Bar(
                    y=sub3["category_name"],
                    x=sub3["item_value"],
                    orientation="h",
                    marker_color=["#78909C", "#4a7c23", "#1E88E5"],
                    text=[f"{v:.1f}%" for v in sub3["item_value"]],
                    textposition="outside",
                    hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
                ))
                fig3.update_layout(
                    height=160, margin=dict(l=60, r=40, t=10, b=20),
                    xaxis=dict(title="%", range=[0, 100]),
                )
                st.plotly_chart(fig3, key=f"b3_{year}", use_container_width=True)

            # 22대 분야별 드릴다운
            with st.expander("22대 기술분야별 R&D 전담조직 보유율"):
                detail3 = df_b3[
                    ~df_b3["category_code"].isin(["00", "감축", "적응"])
                    & (df_b3["item_name"] == "있음")
                ].copy()
                if not detail3.empty:
                    detail3 = detail3.sort_values(["type_name", "category_code"])
                    labels3 = [f"[{t}] {c}" for t, c in zip(detail3["type_name"], detail3["category_name"])]
                    colors3 = [TYPE_COLORS.get(t, "#78909C") for t in detail3["type_name"]]
                    fig3d = go.Figure(go.Bar(
                        y=labels3,
                        x=detail3["item_value"].values,
                        orientation="h",
                        marker_color=colors3,
                        text=[f"{v:.1f}%" for v in detail3["item_value"]],
                        textposition="outside",
                        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
                    ))
                    fig3d.update_layout(
                        height=max(400, len(labels3) * 22),
                        margin=dict(l=200, r=40, t=10, b=20),
                        xaxis=dict(title="%", range=[0, 105]),
                    )
                    st.plotly_chart(fig3d, key=f"b3_detail_{year}", use_container_width=True)

        # b4: 협력경험
        st.markdown("#### 국내외 협력 경험")
        df_b4 = _load_detail("b4", year)
        if not df_b4.empty:
            sub4 = df_b4[df_b4["category_code"].isin(["00", "감축", "적응"]) & (df_b4["item_name"] == "있다")]
            if not sub4.empty:
                fig4 = go.Figure(go.Bar(
                    y=sub4["category_name"],
                    x=sub4["item_value"],
                    orientation="h",
                    marker_color=["#78909C", "#4a7c23", "#1E88E5"],
                    text=[f"{v:.1f}%" for v in sub4["item_value"]],
                    textposition="outside",
                    hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
                ))
                fig4.update_layout(
                    height=160, margin=dict(l=60, r=40, t=10, b=20),
                    xaxis=dict(title="%", range=[0, 100]),
                )
                st.plotly_chart(fig4, key=f"b4_{year}", use_container_width=True)

            # 22대 분야별 드릴다운
            with st.expander("22대 기술분야별 국내외 협력 경험"):
                detail4 = df_b4[
                    ~df_b4["category_code"].isin(["00", "감축", "적응"])
                    & (df_b4["item_name"] == "있다")
                ].copy()
                if not detail4.empty:
                    detail4 = detail4.sort_values(["type_name", "category_code"])
                    labels4 = [f"[{t}] {c}" for t, c in zip(detail4["type_name"], detail4["category_name"])]
                    colors4 = [TYPE_COLORS.get(t, "#78909C") for t in detail4["type_name"]]
                    fig4d = go.Figure(go.Bar(
                        y=labels4,
                        x=detail4["item_value"].values,
                        orientation="h",
                        marker_color=colors4,
                        text=[f"{v:.1f}%" for v in detail4["item_value"]],
                        textposition="outside",
                        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
                    ))
                    fig4d.update_layout(
                        height=max(400, len(labels4) * 22),
                        margin=dict(l=200, r=40, t=10, b=20),
                        xaxis=dict(title="%", range=[0, 105]),
                    )
                    st.plotly_chart(fig4d, key=f"b4_detail_{year}", use_container_width=True)

    # Download
    dl_b = _detail_to_download_df("b1", year)
    for s in ["b2", "b3", "b4"]:
        dl_b = pd.concat([dl_b, _detail_to_download_df(s, year)], ignore_index=True)
    _excel_download(dl_b, f"활동조사_B_기술역량_{year}.xlsx", key=f"dl_b_{year}")


# ────────────────────────────────────────────
# 섹션 C: 성과·협력 (c1~c8)
# ────────────────────────────────────────────
C_SECTIONS = {
    "c1": "R&D 성과 유형",
    "c2": "기술 도입 경험",
    "c3": "기술 도입 장벽",
    "c4": "국내 협력 목적",
    "c5": "해외 협력 목적",
    "c6": "해외 협력 지역",
    "c7": "해외 미진출 사유",
    "c8": "해외 진출 애로",
}


def _render_section_c(year: int, type_filter: str = "전체"):
    st.markdown("### 성과·협력")

    selected_c = st.selectbox(
        "C 섹션 항목 선택",
        list(C_SECTIONS.keys()),
        format_func=lambda k: f"{k} {C_SECTIONS[k]}",
        key=f"c_select_{year}",
    )

    df_c = _load_detail(selected_c, year)
    if df_c.empty:
        st.info(f"{selected_c} 데이터가 없습니다.")
        return

    # Show 감축/적응 소계 bar chart
    sub = df_c[df_c["category_code"].isin(["감축", "적응"])].copy()
    if sub.empty:
        sub = df_c[df_c["category_code"] == "00"].copy()

    if not sub.empty:
        # Pivot: category_name × item_name
        items = sub["item_name"].unique()
        cats = sub["category_name"].unique()

        fig = go.Figure()
        bar_colors = ["#4a7c23", "#1E88E5", "#78909C", "#FFA726"]
        for i, cat in enumerate(cats):
            cat_data = sub[sub["category_name"] == cat]
            item_vals = []
            for item in items:
                v = cat_data[cat_data["item_name"] == item]["item_value"]
                item_vals.append(v.iloc[0] if len(v) > 0 else 0)
            fig.add_trace(go.Bar(
                name=cat,
                x=item_vals,
                y=items,
                orientation="h",
                marker_color=bar_colors[i % len(bar_colors)],
                text=[f"{v:.1f}" for v in item_vals],
                textposition="outside",
                textfont=dict(size=10),
                hovertemplate=cat + "<br>%{y}: %{x:.1f}%<extra></extra>",
            ))

        unit_label = sub["item_unit"].iloc[0] if "item_unit" in sub.columns else "%"
        fig.update_layout(
            barmode="group",
            height=max(300, len(items) * 40 + 80),
            margin=dict(l=200, r=40, t=10, b=30),
            xaxis=dict(title=f"({unit_label})"),
            legend=dict(orientation="h", y=-0.12, font=dict(size=11)),
        )
        st.plotly_chart(fig, key=f"c_{selected_c}_{year}", use_container_width=True)

    # Expandable: 22대 분야별 top items
    with st.expander("22대 기술분야별 상세"):
        detail_c = df_c[~df_c["category_code"].isin(["00", "감축", "적응"])].copy()
        if detail_c.empty:
            st.info("상세 데이터가 없습니다.")
        else:
            # Show as a pivot table
            pivot_c = detail_c.pivot_table(
                index=["type_name", "category_name"],
                columns="item_name",
                values="item_value",
                aggfunc="first",
            ).round(1)
            st.dataframe(pivot_c, use_container_width=True)

    # Download
    dl_c = _detail_to_download_df(selected_c, year)
    _excel_download(dl_c, f"활동조사_{selected_c}_{year}.xlsx", key=f"dl_c_{selected_c}_{year}")


# ────────────────────────────────────────────
# 섹션 D: 정책환경 (d1~d3)
# ────────────────────────────────────────────
D_SECTIONS = {
    "d1": "지원환경 중요도·만족도",
    "d2": "정부지원정책 중요도·만족도",
    "d3": "규제완화 중요도·만족도",
}


def _render_section_d(year: int, type_filter: str = "전체"):
    st.markdown("### 정책환경")

    selected_d = st.selectbox(
        "D 섹션 항목 선택",
        list(D_SECTIONS.keys()),
        format_func=lambda k: f"{k} {D_SECTIONS[k]}",
        key=f"d_select_{year}",
    )

    df_d = _load_detail(selected_d, year)
    if df_d.empty:
        st.info(f"{selected_d} 데이터가 없습니다.")
        return

    # Show 전체 (category_code='00') importance vs satisfaction
    total_d = df_d[df_d["category_code"] == "00"].copy()

    if not total_d.empty:
        # Parse item_name to extract base name and 중요도/만족도/GAP
        imp_items = total_d[total_d["item_name"].str.contains("중요도")].copy()
        sat_items = total_d[total_d["item_name"].str.contains("만족도")].copy()
        gap_items = total_d[total_d["item_name"].str.contains("GAP")].copy()

        # Extract base label (remove _중요도/_만족도/_GAP suffix)
        def _base_label(s):
            return s.replace("_중요도", "").replace("_만족도", "").replace("_GAP", "").strip()

        if not imp_items.empty and not sat_items.empty:
            imp_items = imp_items.copy()
            imp_items["base"] = imp_items["item_name"].apply(_base_label)
            sat_items = sat_items.copy()
            sat_items["base"] = sat_items["item_name"].apply(_base_label)

            bases = imp_items["base"].tolist()

            imp_vals = imp_items.set_index("base")["item_value"]
            sat_vals = sat_items.set_index("base")["item_value"]

            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="중요도",
                y=bases,
                x=[imp_vals.get(b, 0) for b in bases],
                orientation="h",
                marker_color="#1E88E5",
                text=[f"{imp_vals.get(b, 0):.2f}" for b in bases],
                textposition="outside",
                textfont=dict(size=10),
                hovertemplate="%{y}<br>중요도: %{x:.2f}<extra></extra>",
            ))
            fig.add_trace(go.Bar(
                name="만족도",
                y=bases,
                x=[sat_vals.get(b, 0) for b in bases],
                orientation="h",
                marker_color="#66BB6A",
                text=[f"{sat_vals.get(b, 0):.2f}" for b in bases],
                textposition="outside",
                textfont=dict(size=10),
                hovertemplate="%{y}<br>만족도: %{x:.2f}<extra></extra>",
            ))
            fig.update_layout(
                barmode="group",
                height=max(300, len(bases) * 50 + 80),
                margin=dict(l=250, r=40, t=10, b=30),
                xaxis=dict(title="점 (5점 척도)", range=[0, 5.5]),
                legend=dict(orientation="h", y=-0.1, font=dict(size=11)),
            )
            st.plotly_chart(fig, key=f"d_{selected_d}_{year}", use_container_width=True)

            # GAP display
            if not gap_items.empty:
                gap_items = gap_items.copy()
                gap_items["base"] = gap_items["item_name"].apply(_base_label)
                gap_vals = gap_items.set_index("base")["item_value"]

                fig_gap = go.Figure(go.Bar(
                    y=bases,
                    x=[gap_vals.get(b, 0) for b in bases],
                    orientation="h",
                    marker_color="#EF5350",
                    text=[f"{gap_vals.get(b, 0):.2f}" for b in bases],
                    textposition="outside",
                    hovertemplate="%{y}<br>GAP: %{x:.2f}<extra></extra>",
                ))
                fig_gap.update_layout(
                    title=dict(text="중요도-만족도 GAP", font=dict(size=13)),
                    height=max(250, len(bases) * 35 + 60),
                    margin=dict(l=250, r=40, t=35, b=20),
                    xaxis=dict(title="GAP (점)"),
                )
                st.plotly_chart(fig_gap, key=f"d_gap_{selected_d}_{year}", use_container_width=True)

    # Expandable: 감축/적응 비교
    with st.expander("감축/적응 비교"):
        sub_d = df_d[df_d["category_code"].isin(["감축", "적응"])].copy()
        if sub_d.empty:
            st.info("감축/적응 데이터가 없습니다.")
        else:
            pivot_d = sub_d.pivot_table(
                index=["category_name"],
                columns="item_name",
                values="item_value",
                aggfunc="first",
            ).round(2)
            st.dataframe(pivot_d, use_container_width=True)

    # Download
    dl_d = _detail_to_download_df(selected_d, year)
    _excel_download(dl_d, f"활동조사_{selected_d}_{year}.xlsx", key=f"dl_d_{selected_d}_{year}")


# ────────────────────────────────────────────
# 시계열 비교
# ────────────────────────────────────────────
def _render_timeseries():
    st.markdown("## 시계열 비교 (2022 vs 2023)")

    years = [2022, 2023]
    data_by_year = {}
    total_by_year = {}

    for y in years:
        data_by_year[y] = _load_section_a_data(y)
        total_by_year[y] = _load_total_data(y)

    if any(d.empty for d in data_by_year.values()):
        st.warning("시계열 비교에 필요한 데이터가 부족합니다.")
        return

    # ── 감축/적응별 연도 비교 (3대 지표) ──
    metrics = [
        ("revenue_climate", "기후기술 매출액", "백만원"),
        ("rnd_expense_climate", "기후기술 R&D 투자", "백만원"),
        ("employee_count_climate", "기후기술 종사자 수", "명"),
    ]

    for metric_type, metric_label, unit in metrics:
        st.markdown(f"#### {metric_label}")
        col_chart, col_change = st.columns([3, 1])

        with col_chart:
            fig = go.Figure()
            for y in years:
                df_y = data_by_year[y]
                agg = df_y[
                    (df_y["metric_type"] == metric_type) & (df_y["org_size"] == "전체")
                ].groupby("type_name", as_index=False)["metric_value"].sum()

                if unit == "백만원":
                    agg["display_value"] = agg["metric_value"] / 100
                    disp_unit = "억원"
                else:
                    agg["display_value"] = agg["metric_value"]
                    disp_unit = unit

                fig.add_trace(go.Bar(
                    name=str(y),
                    x=agg["type_name"],
                    y=agg["display_value"],
                    text=[f"{v:,.0f}" for v in agg["display_value"]],
                    textposition="outside",
                    textfont=dict(size=12),
                    marker_color="#4a7c23" if y == 2023 else "rgba(74,124,35,0.4)",
                ))

            fig.update_layout(
                barmode="group",
                height=350,
                yaxis=dict(title=disp_unit if unit == "백만원" else unit),
                margin=dict(l=50, r=10, t=30, b=30),
                legend=dict(orientation="h", y=1.1),
            )
            st.plotly_chart(fig, key=f"ts_{metric_type}", width="stretch")

        # 증감률 계산
        with col_change:
            st.markdown("**전년 대비 변화**")
            for type_name in ["감축", "적응"]:
                vals = []
                for y in years:
                    df_y = data_by_year[y]
                    v = df_y[
                        (df_y["metric_type"] == metric_type)
                        & (df_y["org_size"] == "전체")
                        & (df_y["type_name"] == type_name)
                    ]["metric_value"].sum()
                    vals.append(v)

                if len(vals) == 2 and vals[0] > 0:
                    change_pct = (vals[1] - vals[0]) / vals[0] * 100
                    arrow = "▲" if change_pct > 0 else "▼"
                    color = "#c62828" if change_pct > 0 else "#1565c0"
                    st.markdown(
                        f"<span style='color:{color}; font-size:1.1rem;'>"
                        f"{type_name}: {arrow} {abs(change_pct):.1f}%</span>",
                        unsafe_allow_html=True,
                    )

    # Excel download - 시계열 비교 데이터
    ts_rows = []
    for metric_type, metric_label, unit in metrics:
        for y in years:
            df_y = data_by_year[y]
            for type_name in ["감축", "적응"]:
                v = df_y[
                    (df_y["metric_type"] == metric_type)
                    & (df_y["org_size"] == "전체")
                    & (df_y["type_name"] == type_name)
                ]["metric_value"].sum()
                ts_rows.append({"연도": y, "분류": type_name, "지표": metric_label, "값": v, "단위": unit})
    ts_dl_df = pd.DataFrame(ts_rows)
    _excel_download(ts_dl_df, "활동조사_시계열비교.xlsx", key="dl_activity_timeseries")


# ────────────────────────────────────────────
# 종합 평가
# ────────────────────────────────────────────
def _render_evaluation(year_option: str):
    st.markdown("### 종합 평가")

    if "시계열" in year_option:
        years = [2022, 2023]
    else:
        years = [int(year_option.replace("년", ""))]

    # Gather summary data — only use climate-specific metrics
    summaries = {}
    for y in years:
        df_total = _load_total_data(y)
        tot = df_total[df_total["org_size"] == "전체"]
        df_cat = _load_section_a_data(y)

        # Check if any data exists for this year
        if tot.empty and df_cat.empty:
            continue

        def _get(mt):
            r = tot[tot["metric_type"] == mt]
            return r["metric_value"].iloc[0] if len(r) > 0 else 0

        rev_c = _get("revenue_climate")
        rnd_c = _get("rnd_expense_climate")
        emp_c = _get("employee_count_climate")

        # 감축/적응 비중
        mit_rev = df_cat[
            (df_cat["metric_type"] == "revenue_climate")
            & (df_cat["org_size"] == "전체")
            & (df_cat["type_name"] == "감축")
        ]["metric_value"].sum()
        adp_rev = df_cat[
            (df_cat["metric_type"] == "revenue_climate")
            & (df_cat["org_size"] == "전체")
            & (df_cat["type_name"] == "적응")
        ]["metric_value"].sum()
        total_ca = mit_rev + adp_rev
        mit_pct = (mit_rev / total_ca * 100) if total_ca > 0 else 0
        adp_pct = (adp_rev / total_ca * 100) if total_ca > 0 else 0

        summaries[y] = {
            "rev_c": rev_c, "rnd_c": rnd_c,
            "emp_c": emp_c,
            "mit_pct": mit_pct, "adp_pct": adp_pct,
        }

    # If no summary data available for selected year(s)
    if not summaries:
        st.info("선택한 연도의 섹션 A (주요 현황) 데이터가 적재되지 않아 종합 평가를 표시할 수 없습니다.")
        return

    # Build evaluation text
    if len(years) == 1:
        y = years[0]
        if y not in summaries:
            st.info(f"{y}년 섹션 A 데이터 미적재 — 종합 평가를 표시할 수 없습니다.")
            return
        s = summaries[y]

        # Check if data is effectively empty (all zeros)
        if s["rev_c"] == 0 and s["rnd_c"] == 0 and s["emp_c"] == 0:
            st.info(f"{y}년 섹션 A 데이터 미적재 — 종합 평가를 표시할 수 없습니다.")
            return

        eval_html = f"""
        <div class="eval-box">
            <h4>{y}년 기후기술개발 활동 종합 평가</h4>
            <ul>
                <li><b>기후기술 매출</b>: {_format_money(s['rev_c'])}</li>
                <li><b>기후기술 R&D</b>: {_format_money(s['rnd_c'])}</li>
                <li><b>기후기술 종사자</b>: {s['emp_c']:,.0f}명</li>
                <li><b>감축/적응 비중</b>: 기후기술 매출 기준
                    감축 {s['mit_pct']:.1f}% vs 적응 {s['adp_pct']:.1f}%
                    — {'감축 중심 구조' if s['mit_pct'] > s['adp_pct'] else '적응 중심 구조'}</li>
            </ul>
        </div>
        """
    else:
        # 시계열: only compare years with data
        available = [y for y in years if y in summaries]
        if len(available) < 2:
            st.info("시계열 비교에 필요한 두 개 연도의 섹션 A 데이터가 부족합니다.")
            return
        s22, s23 = summaries[available[0]], summaries[available[1]]

        def _chg(old, new):
            if old > 0:
                return (new - old) / old * 100
            return 0

        rev_chg = _chg(s22["rev_c"], s23["rev_c"])
        rnd_chg = _chg(s22["rnd_c"], s23["rnd_c"])
        emp_chg = _chg(s22["emp_c"], s23["emp_c"])

        eval_html = f"""
        <div class="eval-box">
            <h4>{available[0]}-{available[1]} 기후기술개발 활동 변화 종합 평가</h4>
            <ul>
                <li><b>기후기술 매출</b>: {_format_money(s22['rev_c'])} → {_format_money(s23['rev_c'])}
                    ({'▲' if rev_chg > 0 else '▼'} {abs(rev_chg):.1f}%)</li>
                <li><b>기후기술 R&D</b>: {_format_money(s22['rnd_c'])} → {_format_money(s23['rnd_c'])}
                    ({'▲' if rnd_chg > 0 else '▼'} {abs(rnd_chg):.1f}%)</li>
                <li><b>기후기술 종사자</b>: {s22['emp_c']:,.0f}명 → {s23['emp_c']:,.0f}명
                    ({'▲' if emp_chg > 0 else '▼'} {abs(emp_chg):.1f}%)</li>
                <li><b>감축/적응 비중 변화</b>:
                    감축 {s22['mit_pct']:.1f}% → {s23['mit_pct']:.1f}%,
                    적응 {s22['adp_pct']:.1f}% → {s23['adp_pct']:.1f}%</li>
            </ul>
        </div>
        """

    st.markdown(eval_html, unsafe_allow_html=True)
