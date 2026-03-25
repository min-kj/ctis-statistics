# 🐜 Scout: 기후기술 포지셔닝 매트릭스 - 2026-03-22
"""
22대/38대 기후기술의 전략적 위치를 2축 매트릭스로 시각화.
수준조사(38대)와 활동조사(22대)를 교차하여 포지셔닝 분석.
"""
import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import io

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from data_loader import _get_db_connection


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


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

def _get_quadrant_labels(x_metric_key: str) -> dict:
    """X축 지표에 따라 사분면 라벨을 동적으로 생성."""
    if x_metric_key == "gap":
        # 격차: 높을수록 나쁨 → 해석 반전
        return {
            "top_right": "투자 중이나 격차 큼\n(고격차·고투자)",
            "top_left": "핵심 주력\n(저격차·고투자)",
            "bottom_right": "방치 위험\n(고격차·저투자)",
            "bottom_left": "효율적 기술\n(저격차·저투자)",
        }
    else:
        # 수준: 높을수록 좋음
        return {
            "top_right": "핵심 주력 기술\n(고수준·고투자 — 유지·강화)",
            "top_left": "투자 효율 점검\n(저수준·고투자 — 성과 부진?)",
            "bottom_right": "효율적 기술\n(고수준·저투자 — 투자 확대 검토)",
            "bottom_left": "전략적 판단 필요\n(저수준·저투자 — 투자? 포기?)",
        }

COLOR_MAP = {"감축": "#4a7c23", "적응": "#2196F3"}

Y_METRIC_OPTIONS = {
    "기후기술 R&D 투자(억원)": "rnd_expense_climate",
    "기후기술 매출액(억원)": "revenue_climate",
    "기후기술 종사자수(명)": "employee_count_climate",
}

X_METRIC_OPTIONS = {
    "한국 기술수준(%)": "level",
    "한국 기술격차(년)": "gap",
}

MAPPING_NOTE = """\
**⚠️ 매핑 해석 주의**
- **1:1 매핑 (10개)**: 풍력, 바이오, 폐자원, 발전효율, 산업효율, 수송효율, 건물효율, 건강, 물, 농축수산
  → 수준조사↔활동조사 직접 비교 가능
- **1:N 매핑 (12개)**: 태양광·열(2), 해양+수력(2), 수열+지열(2), 수소암모니아(2), 비재생에너지(2), 수소바이오매스(2), 포집/저장(4), 전력통합(3), 모니터링(2), 영향평가(2), 기타탄력성(3), 적응기반(2)
  → 22대 1개 분야 안에 성격이 다른 38대 기술이 포함. 활동조사 값은 합산 기준이므로 해석 주의
"""


# ---------------------------------------------------------------------------
# Data loading (cached)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=3600)
def _load_version_a():
    """22대 기준: 수준조사 38대→22대 역집계 + 활동조사."""
    conn = _get_db_connection()
    try:
        # X축: 수준조사 38대→22대 역집계
        x_query = """
        SELECT t22.category_id,
               t22.category_name,
               t22.type_name,
               AVG(sr.tech_level) AS avg_level,
               AVG(sr.tech_gap) / 12.0 AS avg_gap_year,
               tm.mapping_type
        FROM taxonomy_mapping tm
        JOIN tech_category t22 ON tm.src_category_id = t22.category_id
        JOIN taxonomy tax22 ON t22.taxonomy_id = tax22.taxonomy_id AND tax22.code = 'T22'
        JOIN tech_category t38 ON tm.tgt_category_id = t38.category_id
        JOIN tech_detail td ON td.category_id = t38.category_id AND td.survey_year = 2025
        JOIN survey_result sr ON sr.detail_id = td.detail_id
                              AND sr.survey_year = 2025
                              AND sr.country_code = 'KR'
        GROUP BY t22.category_id, t22.category_name, t22.type_name
        """
        x_df = pd.read_sql_query(x_query, conn)

        # mapping_type per T22 category (detect 1:N)
        mapping_query = """
        SELECT t22.category_id,
               GROUP_CONCAT(DISTINCT tm.mapping_type) AS mapping_types,
               COUNT(DISTINCT t38.category_id) AS t38_count
        FROM taxonomy_mapping tm
        JOIN tech_category t22 ON tm.src_category_id = t22.category_id
        JOIN taxonomy tax22 ON t22.taxonomy_id = tax22.taxonomy_id AND tax22.code = 'T22'
        JOIN tech_category t38 ON tm.tgt_category_id = t38.category_id
        JOIN taxonomy tax38 ON t38.taxonomy_id = tax38.taxonomy_id AND tax38.code = 'T38'
        GROUP BY t22.category_id
        """
        map_df = pd.read_sql_query(mapping_query, conn)

        x_df = x_df.merge(map_df[["category_id", "t38_count"]], on="category_id", how="left")
        x_df["is_1n"] = x_df["t38_count"].fillna(1).astype(int) > 1

        # Y축: 활동조사 (2023 or latest)
        y_frames = {}
        for label, metric_type in Y_METRIC_OPTIONS.items():
            y_query = """
            SELECT tc.category_id,
                   tc.category_name,
                   a.metric_value,
                   a.metric_unit
            FROM activity_survey a
            JOIN tech_category tc ON a.category_id = tc.category_id
            WHERE a.survey_year = 2023
              AND a.metric_type = ?
              AND a.org_size = '전체'
            """
            ydf = pd.read_sql_query(y_query, conn, params=[metric_type])
            y_frames[metric_type] = ydf

        return x_df, y_frames
    finally:
        conn.close()


@st.cache_data(ttl=3600)
def _load_version_b():
    """38대 기준: 1:1 매핑 기술만."""
    conn = _get_db_connection()
    try:
        query = """
        SELECT t38.category_id AS t38_id,
               t38.category_name AS tech_name,
               t38.type_name,
               t22.category_id AS t22_id,
               t22.category_name AS t22_name,
               AVG(sr.tech_level) AS avg_level,
               AVG(sr.tech_gap) / 12.0 AS avg_gap_year
        FROM taxonomy_mapping tm
        JOIN tech_category t22 ON tm.src_category_id = t22.category_id
        JOIN taxonomy tax22 ON t22.taxonomy_id = tax22.taxonomy_id AND tax22.code = 'T22'
        JOIN tech_category t38 ON tm.tgt_category_id = t38.category_id
        JOIN taxonomy tax38 ON t38.taxonomy_id = tax38.taxonomy_id AND tax38.code = 'T38'
        JOIN tech_detail td ON td.category_id = t38.category_id AND td.survey_year = 2025
        JOIN survey_result sr ON sr.detail_id = td.detail_id
                              AND sr.survey_year = 2025
                              AND sr.country_code = 'KR'
        WHERE tm.mapping_type = '1:1'
        GROUP BY t38.category_id, t38.category_name, t38.type_name,
                 t22.category_id, t22.category_name
        """
        x_df = pd.read_sql_query(query, conn)

        # Y축: 활동조사 via t22_id
        y_frames = {}
        for label, metric_type in Y_METRIC_OPTIONS.items():
            y_query = """
            SELECT tc.category_id AS t22_id,
                   a.metric_value,
                   a.metric_unit
            FROM activity_survey a
            JOIN tech_category tc ON a.category_id = tc.category_id
            WHERE a.survey_year = 2023
              AND a.metric_type = ?
              AND a.org_size = '전체'
            """
            ydf = pd.read_sql_query(y_query, conn, params=[metric_type])
            y_frames[metric_type] = ydf

        return x_df, y_frames
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _convert_unit(value, metric_type):
    """백만원 → 억원 변환 (명은 그대로)."""
    if "count" in metric_type:
        return value
    return value / 100.0  # 백만원 → 억원


def _build_plot_df_a(x_df, y_frames, y_metric_key):
    """Version A: 22대 기준 DataFrame 구성."""
    ydf = y_frames.get(y_metric_key)
    if ydf is None or ydf.empty:
        return None

    merged = x_df.merge(
        ydf[["category_id", "metric_value", "metric_unit"]],
        on="category_id",
        how="inner",
    )
    if merged.empty:
        return None

    merged["y_value"] = merged.apply(
        lambda r: _convert_unit(r["metric_value"], y_metric_key), axis=1,
    )
    # Display name with * for 1:N
    merged["display_name"] = merged.apply(
        lambda r: r["category_name"] + "*" if r["is_1n"] else r["category_name"],
        axis=1,
    )
    merged["color"] = merged["type_name"].map(COLOR_MAP).fillna("#999")
    return merged


def _build_plot_df_b(x_df, y_frames, y_metric_key):
    """Version B: 38대 기준 DataFrame 구성."""
    ydf = y_frames.get(y_metric_key)
    if ydf is None or ydf.empty:
        return None

    merged = x_df.merge(
        ydf[["t22_id", "metric_value", "metric_unit"]],
        on="t22_id",
        how="inner",
    )
    if merged.empty:
        return None

    merged["y_value"] = merged.apply(
        lambda r: _convert_unit(r["metric_value"], y_metric_key), axis=1,
    )
    merged["display_name"] = merged["tech_name"].str.replace(r"^\d+\.\s*", "", regex=True)
    merged["color"] = merged["type_name"].map(COLOR_MAP).fillna("#999")
    return merged


def _draw_scatter(plot_df, x_col, x_label, y_label, title, x_metric_key="level"):
    """Draw positioning scatter plot."""
    x_mean = plot_df[x_col].mean()
    y_mean = plot_df["y_value"].mean()

    quadrant_labels = _get_quadrant_labels(x_metric_key)

    fig = go.Figure()

    for type_name in plot_df["type_name"].unique():
        subset = plot_df[plot_df["type_name"] == type_name]
        color = COLOR_MAP.get(type_name, "#999")
        fig.add_trace(go.Scatter(
            x=subset[x_col],
            y=subset["y_value"],
            mode="markers+text",
            name=type_name,
            text=subset["display_name"].str[:8],
            textposition="top center",
            textfont=dict(size=11),
            marker=dict(
                size=14,
                color=color,
                line=dict(width=1, color="white"),
                opacity=0.85,
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                f"{x_label}: %{{x:.1f}}<br>"
                f"{y_label}: %{{y:,.0f}}<br>"
                "분야: %{customdata[1]}<br>"
                "사분면: %{customdata[2]}<br>"
                "<extra></extra>"
            ),
            customdata=subset[["display_name", "type_name", "quadrant"]].values,
        ))

    # Quadrant lines
    fig.add_hline(y=y_mean, line_dash="dash", line_color="gray", opacity=0.5)
    fig.add_vline(x=x_mean, line_dash="dash", line_color="gray", opacity=0.5)

    # Quadrant labels
    x_range = plot_df[x_col].max() - plot_df[x_col].min()
    y_range = plot_df["y_value"].max() - plot_df["y_value"].min()

    if x_range > 0 and y_range > 0:
        positions = [
            (x_mean + x_range * 0.3, y_mean + y_range * 0.35, quadrant_labels["top_right"]),
            (x_mean - x_range * 0.3, y_mean + y_range * 0.35, quadrant_labels["top_left"]),
            (x_mean + x_range * 0.3, y_mean - y_range * 0.35, quadrant_labels["bottom_right"]),
            (x_mean - x_range * 0.3, y_mean - y_range * 0.35, quadrant_labels["bottom_left"]),
        ]
        for qx, qy, qlabel in positions:
            fig.add_annotation(
                x=qx, y=qy, text=qlabel,
                showarrow=False, font=dict(size=10, color="gray"),
                bgcolor="rgba(255,255,255,0.7)", borderpad=4,
            )

    fig.update_layout(
        title=title,
        xaxis_title=x_label,
        yaxis_title=y_label,
        height=650,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )

    return fig, x_mean, y_mean


def _show_quadrants(plot_df, x_col, x_label, y_label, x_mean, y_mean):
    """Show quadrant technology lists."""
    st.markdown("### 사분면별 기술 분포")

    x_short = x_label.split("(")[0].strip()[-2:]
    y_short = y_label.split("(")[0].strip()[-2:]

    col1, col2 = st.columns(2)

    with col1:
        q_tr = plot_df[(plot_df[x_col] >= x_mean) & (plot_df["y_value"] >= y_mean)]
        st.markdown(f"**고{x_short} · 고{y_short}** ({len(q_tr)}개)")
        if len(q_tr) > 0:
            st.write(", ".join(q_tr["display_name"].tolist()))

        q_bl = plot_df[(plot_df[x_col] < x_mean) & (plot_df["y_value"] < y_mean)]
        st.markdown(f"**저{x_short} · 저{y_short}** ({len(q_bl)}개)")
        if len(q_bl) > 0:
            st.write(", ".join(q_bl["display_name"].tolist()))

    with col2:
        q_tl = plot_df[(plot_df[x_col] < x_mean) & (plot_df["y_value"] >= y_mean)]
        st.markdown(f"**저{x_short} · 고{y_short}** ({len(q_tl)}개)")
        if len(q_tl) > 0:
            st.write(", ".join(q_tl["display_name"].tolist()))

        q_br = plot_df[(plot_df[x_col] >= x_mean) & (plot_df["y_value"] < y_mean)]
        st.markdown(f"**고{x_short} · 저{y_short}** ({len(q_br)}개)")
        if len(q_br) > 0:
            st.write(", ".join(q_br["display_name"].tolist()))


def _show_comparison(plot_df, x_col, x_label, y_label, x_mean, y_mean, x_metric_key="level"):
    """Show technology comparison multiselect with quadrant type."""
    st.markdown("### 기술 상세 비교")
    selected = st.multiselect(
        "비교할 기술 선택 (2~3개)",
        plot_df["display_name"].tolist(),
        max_selections=3,
        key="pos_compare",
    )

    if len(selected) >= 2:
        compare_df = plot_df[plot_df["display_name"].isin(selected)].copy()
        compare_df["사분면"] = compare_df["quadrant"]

        # 소수점 포맷
        compare_df[x_col] = compare_df[x_col].round(1)
        compare_df["y_value"] = compare_df["y_value"].round(1)

        show_cols = {
            "display_name": "기술명",
            x_col: x_label,
            "y_value": y_label,
            "type_name": "감축/적응",
            "사분면": "사분면 유형",
        }
        display = compare_df[list(show_cols.keys())].rename(columns=show_cols)
        st.dataframe(display.set_index("기술명"), width="stretch")


# ---------------------------------------------------------------------------
# Main render
# ---------------------------------------------------------------------------

def render():
    st.markdown("### 기후기술 포지셔닝 매트릭스")
    st.markdown("""
    <div class="eval-box">
        38대/22대 기후기술의 전략적 위치를 한 눈에 조망합니다.
    </div>
    """, unsafe_allow_html=True)

    # Version selection
    version = st.radio(
        "기준 선택",
        [
            "22대 기준 (수준조사+활동조사 교차)",
            "38대 기준 (1:1 매핑 10개)",
        ],
        horizontal=True,
        key="pos_version",
    )

    is_version_a = version.startswith("22대")

    # Axis selectors
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        x_label = st.selectbox(
            "X축 지표", list(X_METRIC_OPTIONS.keys()),
            index=0, key="pos_x",
        )
    with col2:
        y_label = st.selectbox(
            "Y축 지표", list(Y_METRIC_OPTIONS.keys()),
            index=0, key="pos_y",
        )
    with col3:
        st.selectbox(
            "색상 구분", ["감축/적응"],
            index=0, key="pos_color",
            disabled=True,
        )

    x_metric_key = X_METRIC_OPTIONS[x_label]
    y_metric_key = Y_METRIC_OPTIONS[y_label]
    x_col = "avg_level" if x_metric_key == "level" else "avg_gap_year"

    # --- Version A: 22대 ---
    if is_version_a:
        x_df, y_frames = _load_version_a()
        if x_df is None or x_df.empty:
            st.warning("수준조사 데이터를 로드할 수 없습니다.")
            return

        plot_df = _build_plot_df_a(x_df, y_frames, y_metric_key)
        if plot_df is None or plot_df.empty:
            st.warning("활동조사 데이터가 없는 분류가 있습니다. Y축 지표를 변경해 보세요.")
            return

        title = "22대 기후기술 포지셔닝 매트릭스"

    # --- Version B: 38대 (1:1 only) ---
    else:
        x_df, y_frames = _load_version_b()
        if x_df is None or x_df.empty:
            st.warning("1:1 매핑 수준조사 데이터를 로드할 수 없습니다.")
            return

        plot_df = _build_plot_df_b(x_df, y_frames, y_metric_key)
        if plot_df is None or plot_df.empty:
            st.warning("활동조사 데이터가 없습니다. Y축 지표를 변경해 보세요.")
            return

        st.info(
            "22개 분야 중 1:1 매핑 **10개**만 표시합니다. "
            "나머지 12개는 1:N 매핑으로 정확한 교차 비교가 불가합니다."
        )
        title = "38대 기후기술 포지셔닝 매트릭스 (1:1 매핑)"

    # Compute quadrant column on plot_df before drawing
    _x_mean_pre = plot_df[x_col].mean()
    _y_mean_pre = plot_df["y_value"].mean()

    def _compute_quadrant(row):
        hi_x = row[x_col] >= _x_mean_pre
        hi_y = row["y_value"] >= _y_mean_pre
        if x_metric_key == "gap":
            x_hi_label, x_lo_label = "고격차", "저격차"
        else:
            x_hi_label, x_lo_label = "고수준", "저수준"
        y_short = y_label.split("(")[0].strip()[-2:]
        if hi_x and hi_y:
            return f"{x_hi_label}·고{y_short}"
        elif not hi_x and hi_y:
            return f"{x_lo_label}·고{y_short}"
        elif hi_x and not hi_y:
            return f"{x_hi_label}·저{y_short}"
        else:
            return f"{x_lo_label}·저{y_short}"

    plot_df["quadrant"] = plot_df.apply(_compute_quadrant, axis=1)

    # Draw scatter
    fig, x_mean, y_mean = _draw_scatter(plot_df, x_col, x_label, y_label, title, x_metric_key)
    st.plotly_chart(fig, width="stretch")

    # Excel download - scatter plot data with quadrant (reuse pre-computed)
    scatter_dl = plot_df.copy()
    scatter_dl["사분면"] = scatter_dl["quadrant"]
    dl_cols = {
        "display_name": "기술명",
        x_col: x_label,
        "y_value": y_label,
        "type_name": "감축/적응",
        "사분면": "사분면",
    }
    scatter_export = scatter_dl[list(dl_cols.keys())].rename(columns=dl_cols)
    _excel_download(scatter_export, "포지셔닝_매트릭스.xlsx", key="dl_positioning")

    # Quadrant summary
    _show_quadrants(plot_df, x_col, x_label, y_label, x_mean, y_mean)

    # Technology comparison
    _show_comparison(plot_df, x_col, x_label, y_label, x_mean, y_mean, x_metric_key)

    # Bottom note
    st.markdown("---")
    st.markdown(MAPPING_NOTE)
