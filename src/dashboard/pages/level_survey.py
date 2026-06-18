# 🐜 Scout: 기후기술수준조사 페이지 - 2026-03-20
"""
CTis 기후기술수준조사 대시보드
- 전체 탭: 5개국 비교 + 시계열
- 세부기술수준 탭: 핵심지표 카드 + 상세 테이블
"""
import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import io

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from data_loader import (
    load_2020_data, load_2025_data, load_2020_category_summary,
    load_mapping_44_to_38, load_timeseries_mapping, load_timeseries_data,
    get_country_averages, aggregate_by_category,
    _get_db_connection,
    COUNTRIES, COUNTRY_CODES, COUNTRY_COLORS, COUNTRY_FLAGS,
)


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


def render():
    sub_tab1, sub_tab2 = st.tabs(["전체", "세부기술수준"])

    with sub_tab1:
        render_overview()

    with sub_tab2:
        render_detail()


def _render_evaluation(df_detail, country_stats, type_stats, year_option):
    """종합 평가 — DB 기반 동적 생성, 한국 고정 + 비교국가 선택"""
    st.markdown("### 종합 평가")

    # 비교국가 선택
    compare = st.selectbox(
        "비교 대상 국가",
        [c for c in COUNTRIES if c != "한국"],
        index=3,  # 기본: 미국
        key="eval_compare",
    )
    compare_code = COUNTRY_CODES[COUNTRIES.index(compare)]

    # 한국 기본 지표 계산
    kr_level = df_detail["kr_level"].mean() if "kr_level" in df_detail.columns else 0
    kr_gap_month = df_detail["kr_gap"].mean() if "kr_gap" in df_detail.columns else 0
    kr_gap_year = kr_gap_month / 12  # 개월→년

    # 비교국 지표
    cmp_lv_col = f"{compare_code}_level"
    cmp_gp_col = f"{compare_code}_gap"
    cmp_level = df_detail[cmp_lv_col].mean() if cmp_lv_col in df_detail.columns else 0
    cmp_gap_year = (df_detail[cmp_gp_col].mean() / 12) if cmp_gp_col in df_detail.columns else 0

    # 한국 순위 계산
    country_levels = {}
    for c, code in zip(COUNTRIES, COUNTRY_CODES):
        lv_col = f"{code}_level"
        if lv_col in df_detail.columns:
            country_levels[c] = df_detail[lv_col].mean()
    kr_rank = sorted(country_levels.values(), reverse=True).index(country_levels.get("한국", 0)) + 1

    # 선도기술 수 (한국이 최고보유국)
    leading_kr = 0
    if "leading_country" in df_detail.columns:
        leading_kr = len(df_detail[df_detail["leading_country"].str.contains("한국", na=False)])

    # 비교국 선도기술 수
    leading_cmp = 0
    if "leading_country" in df_detail.columns:
        leading_cmp = len(df_detail[df_detail["leading_country"].str.contains(compare, na=False)])

    # 감축/적응 격차 (개월→년)
    gap_mit = 0
    gap_adp = 0
    if "type" in df_detail.columns and "kr_gap" in df_detail.columns:
        mit_data = df_detail[df_detail["type"] == "감축"]["kr_gap"]
        adp_data = df_detail[df_detail["type"] == "적응"]["kr_gap"]
        gap_mit = mit_data.mean() / 12 if len(mit_data) > 0 else 0
        gap_adp = adp_data.mean() / 12 if len(adp_data) > 0 else 0

    # 수준 차이
    level_diff = kr_level - cmp_level
    gap_diff = kr_gap_year - cmp_gap_year

    # 한국 vs 비교국 — 2020/2025 시계열 비교 바차트
    # 2020 데이터 로드
    df_2020 = load_2020_data()
    kr_level_2020 = df_2020["kr_level"].mean() if "kr_level" in df_2020.columns else 0
    kr_gap_2020 = df_2020["kr_gap"].mean() if "kr_gap" in df_2020.columns else 0  # 년 단위
    cmp_lv_2020 = df_2020[cmp_lv_col].mean() if cmp_lv_col in df_2020.columns else 0
    cmp_gp_2020 = df_2020[cmp_gp_col].mean() if cmp_gp_col in df_2020.columns else 0

    eval_c1, eval_c2 = st.columns(2)

    # 왼쪽: 한국 2020→2025
    with eval_c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="기술수준", x=["2020", "2025"],
            y=[kr_level_2020, kr_level],
            marker_color=["rgba(255,107,107,0.5)", "#FF6B6B"],
            text=[f"{kr_level_2020:.1f}%", f"{kr_level:.1f}%"],
            textposition="outside", textfont=dict(size=14, color="#333"),
            width=0.55,
        ))
        fig.add_trace(go.Bar(
            name="기술격차(년)", x=["2020", "2025"],
            y=[kr_gap_2020, kr_gap_year],
            marker_color=["rgba(200,80,80,0.3)", "rgba(200,60,60,0.6)"],
            text=[f"{kr_gap_2020:.1f}년", f"{kr_gap_year:.1f}년"],
            textposition="outside", textfont=dict(size=14, color="#333"),
            width=0.55,
        ))
        fig.update_layout(
            title=dict(text="한국", font=dict(size=16)),
            barmode="group", height=300,
            yaxis=dict(range=[0, 110]),
            margin=dict(l=30, r=10, t=50, b=30),
            legend=dict(orientation="h", y=1.12, font=dict(size=11)),
        )
        st.plotly_chart(fig, width="stretch")

    # 오른쪽: 비교국 2020→2025
    with eval_c2:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="기술수준", x=["2020", "2025"],
            y=[cmp_lv_2020, cmp_level],
            marker_color=["rgba(150,150,150,0.4)", COUNTRY_COLORS.get(compare, "#999")],
            text=[f"{cmp_lv_2020:.1f}%", f"{cmp_level:.1f}%"],
            textposition="outside", textfont=dict(size=14, color="#333"),
            width=0.55,
        ))
        fig.add_trace(go.Bar(
            name="기술격차(년)", x=["2020", "2025"],
            y=[cmp_gp_2020, cmp_gap_year],
            marker_color=["rgba(200,80,80,0.3)", "rgba(200,60,60,0.6)"],
            text=[f"{cmp_gp_2020:.1f}년", f"{cmp_gap_year:.1f}년"],
            textposition="outside", textfont=dict(size=14, color="#333"),
            width=0.55,
        ))
        fig.update_layout(
            title=dict(text=compare, font=dict(size=16)),
            barmode="group", height=300,
            yaxis=dict(range=[0, 110]),
            margin=dict(l=30, r=10, t=50, b=30),
            legend=dict(orientation="h", y=1.12, font=dict(size=11)),
        )
        st.plotly_chart(fig, width="stretch")

    st.markdown(f"""
    <div class="eval-box">
        <h4>한국의 기후기술 수준 종합 평가</h4>
        <ul>
            <li><b>기술수준</b>: 한국 {kr_level:.1f}% (5개국 중 {kr_rank}위)
                — {compare} {cmp_level:.1f}% 대비 {abs(level_diff):.1f}%p {'높음' if level_diff > 0 else '낮음'}</li>
            <li><b>기술격차</b>: 최고기술국 대비 평균 {kr_gap_year:.1f}년
                — {compare}({cmp_gap_year:.1f}년) 대비 {abs(gap_diff):.1f}년 {'작음(우수)' if gap_diff < 0 else '큼(열위)'}</li>
            <li><b>분야별 격차</b>: 감축({gap_mit:.1f}년) {'<' if gap_mit < gap_adp else '>'} 적응({gap_adp:.1f}년)
                — {'적응 분야의 격차가 더 큼' if gap_adp > gap_mit else '감축 분야의 격차가 더 큼'}</li>
            <li><b>선도기술</b>: 한국 {leading_kr}건 vs {compare} {leading_cmp}건
                (최고기술보유국 기준)</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)


_CP_CSS = """
<style>
.cp-grid{display:grid;gap:14px;margin:2px 0 6px}
.cp4{grid-template-columns:repeat(4,1fr)}
.cp3{grid-template-columns:repeat(3,1fr)}
.cp-card{border:1px solid #e7ebf1;border-radius:12px;overflow:hidden;background:#fff;
  display:flex;flex-direction:column;min-height:196px}
.cp-hd{background:#f7f9fc;padding:12px 16px;min-height:50px}
.cp-t{font-size:13.5px;font-weight:800;color:#2a3f60;line-height:1.3}
.cp-t small{display:block;font-weight:600;color:#64748b;font-size:11px;margin-top:2px}
.cp-mid{flex:1;display:flex;align-items:center;justify-content:center;padding:8px;position:relative}
.cp-val{font-size:32px;font-weight:800;color:#2a3f60;line-height:1}
.cp-u{font-size:14px;font-weight:700;color:#64748b;margin-left:3px}
.cp-badge{position:absolute;top:0;left:0;font-size:10px;font-weight:800;color:#fff;
  padding:3px 9px;border-radius:0 0 8px 0}
.cp-rankbox{padding:0 16px 16px}
.cp-rl{text-align:center;font-size:12px;color:#64748b;margin-bottom:8px}
.cp-rl b{color:#2a3f60;font-weight:800}
.cp-bar{position:relative;height:8px;border-radius:6px;margin-top:20px;
  background:linear-gradient(90deg,#e8615f 0%,#f0a35e 22%,#f2d97b 44%,#bcd6a6 62%,#7fb6d6 80%,#3d6fb0 100%)}
.cp-dot{position:absolute;top:50%;width:10px;height:10px;border-radius:50%;background:#fff;
  border:2px solid #9aa7bb;transform:translate(-50%,-50%)}
.cp-tick{position:absolute;top:-4px;width:3px;height:16px;border-radius:2px;background:#16263f;
  transform:translateX(-50%);box-shadow:0 0 0 2px #fff}
.cp-me{position:absolute;top:-22px;transform:translateX(-50%);font-size:13px}
.cp-ends{display:flex;justify-content:space-between;font-size:10px;color:#9aa7bb;margin-top:6px}
.ctis-link{border:1px solid #e7ebf1;border-radius:10px;background:#f7f9fc;padding:10px 12px;margin-bottom:6px}
.ctis-link .t{font-weight:700;color:#2a3f60;font-size:13px}
.ctis-link .d{font-size:11.5px;color:#64748b;margin-top:2px;line-height:1.3}
</style>
"""


def _cp_rank(vals, code, better):
    arr = sorted(((c, v) for c, v in vals.items() if v is not None),
                 key=lambda o: o[1], reverse=(better == "high"))
    return [c for c, _ in arr].index(code) + 1


def _cp_bar(vals, code, better, flags, code2name):
    items = [(c, v) for c, v in vals.items() if v is not None]
    vs = [v for _, v in items]
    mn, mx = min(vs), max(vs)
    span = (mx - mn) or 1

    def pos(v):
        p = (v - mn) / span
        if better == "low":
            p = 1 - p
        return 6 + p * 88

    dots = "".join(
        f'<span class="cp-dot" style="left:{pos(v):.1f}%"></span>'
        for c, v in items if c != code
    )
    me = dict(items).get(code)
    mehtml = ""
    if me is not None:
        flag = flags.get(code2name.get(code, ""), "")
        mehtml = (f'<span class="cp-me" style="left:{pos(me):.1f}%">{flag}</span>'
                  f'<span class="cp-tick" style="left:{pos(me):.1f}%"></span>')
    return f'<div class="cp-bar">{dots}{mehtml}</div>'


def _cp_card(title, sub, value, unit, vals, code, better, code2name, flags, badge=None):
    rank = _cp_rank(vals, code, better)
    badge_html = ""
    if badge:
        bg = "#3d8b40" if badge == "강점" else "#d9534f"
        badge_html = f'<div class="cp-badge" style="background:{bg}">{badge}</div>'
    sub_html = f"<small>{sub}</small>" if sub else ""
    return (
        f'<div class="cp-card"><div class="cp-hd"><div class="cp-t">{title}{sub_html}</div></div>'
        f'<div class="cp-mid">{badge_html}<div class="cp-val">{value}<span class="cp-u">{unit}</span></div></div>'
        f'<div class="cp-rankbox"><div class="cp-rl">5개국 중 <b>{rank}위</b></div>'
        f'{_cp_bar(vals, code, better, flags, code2name)}'
        f'<div class="cp-ends"><span>하위</span><span>상위</span></div></div></div>'
    )


def _render_country_landing():
    """벤치마킹 반영: Climate Watch식 국가 경쟁력 프로파일 (mockups/country-profile-cards 이식)"""
    df, cat_stats, country_stats, type_stats = load_2025_data()
    cat_col = ("category" if "category" in df.columns
               else ("category_38" if "category_38" in df.columns else None))
    codes, names = COUNTRY_CODES, COUNTRIES
    code2name = dict(zip(codes, names))
    name2code = dict(zip(names, codes))

    # 카테고리별 5개국 평균 수준 → 분야 선도국 카운트
    cat_levels = {}
    if cat_col:
        for cat, g in df.groupby(cat_col):
            cat_levels[cat] = {cc: float(g[f"{cc}_level"].mean())
                               for cc in codes if f"{cc}_level" in df.columns}
    leading_cat = {cc: 0 for cc in codes}
    for cat, lv in cat_levels.items():
        if lv:
            leading_cat[max(lv, key=lv.get)] += 1

    summary = {}
    for cc in codes:
        lvcol, gpcol = f"{cc}_level", f"{cc}_gap"
        summary[cc] = {
            "avg_level": float(df[lvcol].mean()) if lvcol in df.columns else 0,
            "avg_gap": float(df[gpcol].mean()) if gpcol in df.columns else 0,
            "leading_count": (int(df["leading_country"].str.contains(code2name[cc], na=False).sum())
                              if "leading_country" in df.columns else 0),
            "leading_cat": leading_cat[cc],
        }

    sel_name = st.selectbox("국가 선택", names, index=0, key="cp_country")
    sel = name2code[sel_name]

    st.markdown(_CP_CSS, unsafe_allow_html=True)
    st.markdown(f"#### {COUNTRY_FLAGS[sel_name]} {sel_name}의 기후기술 경쟁력 지표")
    st.caption("에너지·산업·적응 등 38대 분야 / 157개 세부기술 대표 지표와 5개국(한·중·일·미·EU) 중 순위")

    # 종합 경쟁력 4 카드
    aggs = [
        ("평균 기술수준", "157개 세부기술 평균", "avg_level", "%", "high", lambda v: f"{v:.1f}"),
        ("선도기술 보유", "최고기술 보유 수", "leading_count", "개", "high", lambda v: f"{int(v)}"),
        ("선도 분야", "38대 분야 중 선도", "leading_cat", "개", "high", lambda v: f"{int(v)}"),
        ("최고국 대비 평균 격차", "작을수록 우수", "avg_gap", "개월", "low", lambda v: f"{v:.1f}"),
    ]
    html = ""
    for title, sub, key, unit, better, fmt in aggs:
        vals = {cc: summary[cc][key] for cc in codes}
        html += _cp_card(title, sub, fmt(vals[sel]), unit, vals, sel, better, code2name, COUNTRY_FLAGS)
    st.markdown(f'<div class="cp-grid cp4">{html}</div>', unsafe_allow_html=True)

    # 분야별 강점/약점
    if cat_col and cat_levels:
        sel_levels = {cat: lv.get(sel, 0) for cat, lv in cat_levels.items() if lv}
        ranked = sorted(sel_levels, key=sel_levels.get, reverse=True)
        strong, weak = ranked[:3], ranked[-3:][::-1]
        st.markdown(f"###### 분야별 경쟁력 · {sel_name} 강점 3 · 약점 3")
        fhtml = ""
        for cat in strong + weak:
            vals = {cc: cat_levels[cat].get(cc, 0) for cc in codes}
            kind = "강점" if cat in strong else "약점"
            fhtml += _cp_card(cat, "", f"{vals[sel]:.1f}", "%", vals, sel, "high",
                              code2name, COUNTRY_FLAGS, badge=kind)
        st.markdown(f'<div class="cp-grid cp3">{fhtml}</div>', unsafe_allow_html=True)

    st.caption("출처: CTis 기후기술 수준조사(2025) · 순위 막대는 5개국 분포, 굵은 표식=선택 국가(오른쪽=우수)")


def _render_ctis_links():
    """CTis 실제 기능 연계 바로가기 (ctos pages.md 기준 실제 메뉴)"""
    base = "https://www.ctis.re.kr/menu.es?mid="
    items = [
        ("📊 R&D 투자·성과", base + "a10201010100", "38대 기술별 국가 R&D 투자·성과 분석"),
        ("🧭 정책시계열", base + "a10401010000", "기술·부처별 정책 제·개정 이력"),
        ("📰 정책 보도자료", base + "a10402010000", "최신 기후기술 정책 보도자료"),
        ("🗂 기술인벤토리", base + "a10304010000", "38↔100↔45대 분류·기술 상세"),
    ]
    st.markdown("###### CTis 연계 — 관련 기능 바로가기")
    cols = st.columns(len(items))
    for col, (t, url, desc) in zip(cols, items):
        with col:
            st.markdown(
                f'<div class="ctis-link"><div class="t">{t}</div><div class="d">{desc}</div></div>',
                unsafe_allow_html=True,
            )
            st.link_button("바로가기 ↗", url, use_container_width=True)


def render_overview():
    """전체 탭 — 통계 개요 + 5개국 비교 + 시계열"""

    # ── 벤치마킹 반영: Climate Watch식 국가 경쟁력 프로파일 + CTis 연계 ──
    _render_country_landing()
    _render_ctis_links()
    st.markdown("---")

    # ── 통계 개요 ──
    with st.expander("통계 개요", expanded=False):
        st.markdown("""
| 항목 | 내용 |
|------|------|
| **조사명** | 기후기술 수준조사 |
| **조사 목적** | 주요 5개국(한·중·일·EU·미)의 기후기술 수준 및 격차 파악 |
| **조사 방법** | 전문가 델파이 조사 (2차 이상) |
| **분류체계** | 2020년: 44대 (185개 세부기술) → 2025년: 38대 (157개 세부기술) |
| **조사 주체** | 국가녹색기술연구소 (NIGT) |
| **주요 지표** | 기술수준(%), 기술격차(년), 최고기술보유국, R&D 역량 |
        """)

    # ── 보고서 다운로드 ──
    st.markdown("##### 결과보고서")
    report_col1, report_col2, report_col3 = st.columns(3)

    # 수준조사 보고서 — nigt.re.kr 외부 링크 (홈페이지 원본과 동일)
    reports = [
        ("2020년 기후기술 수준조사(총괄/감축)", "https://nigt.re.kr/gtck/gtcPublication.do?mode=view&articleNo=2485"),
        ("2020년 기후기술 수준평가(적응/융복합)", "https://nigt.re.kr/gtck/gtcPublication.do?mode=view&articleNo=2486"),
        ("2020년 기후기술 수준평가(요약)", "https://nigt.re.kr/gtck/gtcPublication.do?mode=view&articleNo=2487"),
    ]
    for col, (label, url) in zip([report_col1, report_col2, report_col3], reports):
        with col:
            st.link_button(f"📄 {label}", url)

    st.markdown("---")

    # 데이터 로드
    df_2025, cat_stats_2025, country_stats_2025, type_stats_2025 = load_2025_data()
    cat_summary_2020 = load_2020_category_summary()

    # 연도 선택
    col_filter, _ = st.columns([1, 2])
    with col_filter:
        year_option = st.selectbox(
            "분석 기간",
            ["2025년 (38대)", "2020년 (44대)", "2020-2025 시계열 비교"],
            key="overview_year",
        )

    if year_option == "2025년 (38대)":
        _render_country_comparison_2025(country_stats_2025, type_stats_2025)
    elif year_option == "2020년 (44대)":
        _render_country_comparison_2020(cat_summary_2020)
    else:
        _render_timeseries_comparison(cat_summary_2020, cat_stats_2025)

    # 종합 평가 (DB 기반 동적 생성)
    _render_evaluation(df_2025, country_stats_2025, type_stats_2025, year_option)


def _render_country_comparison_2025(country_stats, type_stats):
    """2025년 5개국 비교"""
    st.markdown("### 국가별 기술 현황")

    # 국가별 통계에서 데이터 추출
    levels = []
    gaps = []
    for _, row in country_stats.iterrows():
        name = row.get("국가", "")
        if name in COUNTRIES:
            levels.append({"국가": name, "기술수준": float(row["평균_기술수준"])})
            gaps.append({"국가": name, "기술격차": float(row["평균_기술격차"])})

    # 순서 맞추기
    level_map = {d["국가"]: d["기술수준"] for d in levels}
    gap_map = {d["국가"]: d["기술격차"] for d in gaps}

    ordered_levels = [level_map.get(c, 0) for c in COUNTRIES]
    ordered_gaps_month = [gap_map.get(c, 0) for c in COUNTRIES]
    ordered_gaps = [g / 12 for g in ordered_gaps_month]  # 개월→년 변환

    col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES,
            y=ordered_levels,
            marker_color=[COUNTRY_COLORS[c] for c in COUNTRIES],
            text=[f"{v:.1f}%" for v in ordered_levels],
            textposition="outside",
            textfont=dict(size=14),
        )])
        fig.update_layout(
            title="국가별 기술수준 (2025)", yaxis_title="기술수준 (%)",
            height=400, yaxis=dict(range=[0, 105]),
        )
        st.plotly_chart(fig, width="stretch")

    with col2:
        # 격차: 작을수록 좋음 → 빨강(큼)~초록(작음) 색상
        gap_max = max(ordered_gaps) if ordered_gaps else 1
        gap_colors = [
            f"rgb({min(255, int(200 * v / gap_max))}, {min(200, int(200 * (1 - v / gap_max)))}, 80)"
            for v in ordered_gaps
        ]
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES,
            y=ordered_gaps,
            marker_color=gap_colors,
            text=[f"{v:.1f}년" for v in ordered_gaps],
            textposition="outside",
            textfont=dict(size=14),
        )])
        fig.update_layout(
            title="국가별 기술격차 (2025) ↓작을수록 우수",
            yaxis_title="기술격차 (년)",
            height=400,
        )
        st.plotly_chart(fig, width="stretch")

    # 감축/적응 분류별 — 4열 레이아웃 (감축 수준|격차 | 적응 수준|격차)
    st.markdown("### 감축·적응 분야별 현황")
    if type_stats is not None and len(type_stats) > 1:
        c1, c2, c3, c4 = st.columns(4)

        for idx, (type_name, col_lv, col_gp) in enumerate([
            ("감축", c1, c2), ("적응", c3, c4)
        ]):
            row = type_stats[type_stats["분류"] == type_name]
            if len(row) == 0:
                continue
            row = row.iloc[0]

            # 수준 추출
            type_levels = []
            for c in COUNTRIES:
                col_name = f"{c}_평균수준"
                type_levels.append(float(row[col_name]) if col_name in row.index else 0)

            # 격차 추출 (개월→년 변환)
            type_gaps = []
            for c in COUNTRIES:
                col_name = f"{c}_평균격차"
                val = float(row[col_name]) if col_name in row.index else 0
                type_gaps.append(val / 12)  # 개월→년

            # 수준 차트
            with col_lv:
                fig = go.Figure(data=[go.Bar(
                    x=COUNTRIES, y=type_levels,
                    marker_color=[COUNTRY_COLORS[c] for c in COUNTRIES],
                    text=[f"{v:.1f}" for v in type_levels],
                    textposition="outside", textfont=dict(size=10),
                )])
                fig.update_layout(
                    title=f"[{type_name}] 기술수준 (%)",
                    yaxis=dict(range=[0, 105]), height=300,
                    margin=dict(l=30, r=10, t=40, b=30),
                    xaxis_tickfont=dict(size=9),
                )
                st.plotly_chart(fig, width="stretch")

            # 격차 차트
            with col_gp:
                gp_max = max(type_gaps) if type_gaps else 1
                gp_colors = [
                    f"rgb({min(255, int(200 * v / gp_max))}, {min(200, int(200 * (1 - v / gp_max)))}, 80)"
                    for v in type_gaps
                ]
                fig = go.Figure(data=[go.Bar(
                    x=COUNTRIES, y=type_gaps,
                    marker_color=gp_colors,
                    text=[f"{v:.1f}" for v in type_gaps],
                    textposition="outside", textfont=dict(size=10),
                )])
                fig.update_layout(
                    title=f"[{type_name}] 기술격차 (년) ↓",
                    height=300,
                    margin=dict(l=30, r=10, t=40, b=30),
                    xaxis_tickfont=dict(size=9),
                )
                st.plotly_chart(fig, width="stretch")

    # Excel download - 국가별 수준/격차 데이터
    dl_data = pd.DataFrame({
        "국가": COUNTRIES,
        "기술수준(%)": ordered_levels,
        "기술격차(년)": ordered_gaps,
    })
    _excel_download(dl_data, "국가별_기술현황_2025.xlsx", key="dl_country_2025")


def _render_country_comparison_2020(cat_summary):
    """2020년 5개국 비교 + 감축/적응"""
    st.markdown("### 국가별 기술 현황 (2020)")

    avgs = get_country_averages(cat_summary)
    ordered_levels = [avgs[c]["level"] for c in COUNTRIES]
    ordered_gaps = [avgs[c]["gap"] for c in COUNTRIES]

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES, y=ordered_levels,
            marker_color=[COUNTRY_COLORS[c] for c in COUNTRIES],
            text=[f"{v:.1f}%" for v in ordered_levels],
            textposition="outside", textfont=dict(size=14),
        )])
        fig.update_layout(
            title="국가별 기술수준 (2020)", yaxis_title="기술수준 (%)",
            height=400, yaxis=dict(range=[0, 105]),
        )
        st.plotly_chart(fig, width="stretch")

    with col2:
        gap_max = max(ordered_gaps) if ordered_gaps else 1
        gap_colors = [
            f"rgb({min(255, int(200 * v / gap_max))}, {min(200, int(200 * (1 - v / gap_max)))}, 80)"
            for v in ordered_gaps
        ]
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES, y=ordered_gaps,
            marker_color=gap_colors,
            text=[f"{v:.1f}년" for v in ordered_gaps],
            textposition="outside", textfont=dict(size=14),
        )])
        fig.update_layout(
            title="국가별 기술격차 (2020) ↓작을수록 우수",
            yaxis_title="기술격차 (년)",
            height=400,
        )
        st.plotly_chart(fig, width="stretch")

    # 감축/적응 분류별 (2020 — cat_summary의 type 컬럼 기반)
    st.markdown("### 감축·적응 분야별 현황 (2020)")
    if "type" in cat_summary.columns:
        c1, c2, c3, c4 = st.columns(4)
        for type_name, col_lv, col_gp in [("감축", c1, c2), ("적응", c3, c4)]:
            type_df = cat_summary[cat_summary["type"] == type_name]
            if len(type_df) == 0:
                continue
            type_avgs = get_country_averages(type_df)
            t_levels = [type_avgs.get(c, {}).get("level", 0) for c in COUNTRIES]
            t_gaps = [type_avgs.get(c, {}).get("gap", 0) for c in COUNTRIES]

            with col_lv:
                fig = go.Figure(data=[go.Bar(
                    x=COUNTRIES, y=t_levels,
                    marker_color=[COUNTRY_COLORS[c] for c in COUNTRIES],
                    text=[f"{v:.1f}" for v in t_levels],
                    textposition="outside", textfont=dict(size=10),
                )])
                fig.update_layout(
                    title=f"[{type_name}] 기술수준 (%)",
                    yaxis=dict(range=[0, 105]), height=300,
                    margin=dict(l=30, r=10, t=40, b=30),
                    xaxis_tickfont=dict(size=9),
                )
                st.plotly_chart(fig, width="stretch")

            with col_gp:
                gp_max = max(t_gaps) if t_gaps else 1
                gp_colors = [
                    f"rgb({min(255, int(200 * v / gp_max))}, {min(200, int(200 * (1 - v / gp_max)))}, 80)"
                    for v in t_gaps
                ]
                fig = go.Figure(data=[go.Bar(
                    x=COUNTRIES, y=t_gaps,
                    marker_color=gp_colors,
                    text=[f"{v:.1f}" for v in t_gaps],
                    textposition="outside", textfont=dict(size=10),
                )])
                fig.update_layout(
                    title=f"[{type_name}] 기술격차 (년) ↓",
                    height=300,
                    margin=dict(l=30, r=10, t=40, b=30),
                    xaxis_tickfont=dict(size=9),
                )
                st.plotly_chart(fig, width="stretch")

    # Excel download - 2020 국가별 수준/격차 데이터
    dl_data_2020 = pd.DataFrame({
        "국가": COUNTRIES,
        "기술수준(%)": ordered_levels,
        "기술격차(년)": ordered_gaps,
    })
    _excel_download(dl_data_2020, "국가별_기술현황_2020.xlsx", key="dl_country_2020")


def _render_timeseries_comparison(cat_2020, cat_stats_2025):
    """2020↔2025 시계열 비교 — 매핑 기반, 단위 정규화, 세부기술 직접 평균"""
    df_old, df_new, info = load_timeseries_data()

    if info["old_year"] is None or df_old.empty or df_new.empty:
        st.warning("시계열 비교 데이터를 불러올 수 없습니다.")
        return

    old_year = info["old_year"]
    new_year = info["new_year"]
    old_tax = info["old_taxonomy"]
    new_tax = info["new_taxonomy"]
    old_cat_count = info["old_category_count"]
    new_cat_count = info["new_category_count"]

    st.markdown(f"### 시계열 비교 ({old_year} → {new_year})")

    # Mapping statistics
    st.info(
        f"{old_tax} {old_cat_count}개 중 {old_cat_count - info['unmapped_old']}개 → "
        f"{new_tax} 매핑 완료, {info['unmapped_old']}개 매핑 불가  |  "
        f"{new_tax} {new_cat_count}개 중 {info['unmapped_new']}개는 신규 (매핑 대상 없음)  \n"
        f"매핑된 기술 대분류 {info['mapped_count']}건 기준으로 비교합니다. "
        f"기술격차는 모두 **년** 단위로 정규화되었습니다 "
        f"({new_year}년: {info['new_gap_unit_original']}(개월) → year 변환, 산식: 개월÷12)."
    )

    # Compute country averages directly from detail-level data (not category averages)
    levels_old = []
    gaps_old = []
    levels_new = []
    gaps_new = []
    for code in COUNTRY_CODES:
        lv_col = f"{code}_level"
        gp_col = f"{code}_gap"
        levels_old.append(df_old[lv_col].mean() if lv_col in df_old.columns else 0)
        gaps_old.append(df_old[gp_col].mean() if gp_col in df_old.columns else 0)
        levels_new.append(df_new[lv_col].mean() if lv_col in df_new.columns else 0)
        gaps_new.append(df_new[gp_col].mean() if gp_col in df_new.columns else 0)

    col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name=str(old_year), x=COUNTRIES, y=levels_old,
            marker_color="rgba(100,100,100,0.4)",
            text=[f"{v:.1f}" for v in levels_old], textposition="outside",
        ))
        fig.add_trace(go.Bar(
            name=str(new_year), x=COUNTRIES, y=levels_new,
            marker_color=[COUNTRY_COLORS[c] for c in COUNTRIES],
            text=[f"{v:.1f}" for v in levels_new], textposition="outside",
        ))
        fig.update_layout(
            title=f"기술수준 변화 ({old_year} → {new_year})",
            barmode="group", height=400, yaxis=dict(range=[0, 105]),
            yaxis_title="기술수준 (%)",
        )
        st.plotly_chart(fig, width="stretch")

    with col2:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name=str(old_year), x=COUNTRIES, y=gaps_old,
            marker_color="rgba(180,80,80,0.4)",
            text=[f"{v:.1f}" for v in gaps_old], textposition="outside",
        ))
        fig.add_trace(go.Bar(
            name=str(new_year), x=COUNTRIES, y=gaps_new,
            marker_color="rgba(200,60,60,0.7)",
            text=[f"{v:.1f}" for v in gaps_new], textposition="outside",
        ))
        fig.update_layout(
            title=f"기술격차 변화 ({old_year} → {new_year}) ↓작을수록 우수",
            barmode="group", height=400,
            yaxis_title="기술격차 (년)",
        )
        st.plotly_chart(fig, width="stretch")

    st.caption(
        f"* 매핑된 기술만 포함 ({old_year}: {len(df_old)}건, {new_year}: {len(df_new)}건 세부기술).  \n"
        f"* 국가별 평균은 세부기술 수준에서 직접 산출 (대분류 평균의 재평균이 아님).  \n"
        f"* 기술격차 정규화: {new_year}년 원본({info['new_gap_unit_original']}) → "
        f"{info['gap_unit_normalized']}(년)으로 변환 (÷12)."
    )

    # ── 섹션 2: 분야별 기술 변화 (38대 매핑 기준, 변화량 바차트) ──
    st.markdown(f"### 분야별 기술수준 변화 ({old_year} → {new_year})")

    # 38대 카테고리 기준 집계 준비
    conn = _get_db_connection()
    try:
        mapping_df = pd.read_sql_query("""
            SELECT tm.src_category_id, tc_new.category_name AS new_category
            FROM taxonomy_mapping tm
            JOIN tech_category tc_new ON tm.tgt_category_id = tc_new.category_id
        """, conn)
    finally:
        conn.close()

    # df_old → 38대 기준 재집계
    old_with_new_cat = df_old.merge(
        mapping_df, left_on="category_id", right_on="src_category_id", how="left"
    )

    # 한국 변화량 계산 (항상 표시)
    cat_new_kr = df_new.groupby("category").agg(
        kr_level=("kr_level", "mean"), type=("type", "first"),
    ).reset_index()
    cat_old_kr = old_with_new_cat.groupby("new_category").agg(
        kr_level=("kr_level", "mean"),
    ).reset_index().rename(columns={"new_category": "category"})

    merged = cat_new_kr.merge(cat_old_kr, on="category", suffixes=("_new", "_old"), how="inner")
    merged["kr_delta"] = merged["kr_level_new"] - merged["kr_level_old"]
    merged = merged.sort_values("kr_delta", ascending=True)

    # 비매핑 기술
    unmapped_new = cat_new_kr[~cat_new_kr["category"].isin(merged["category"])]

    # 비교국가 선택 (체크박스)
    show_compare = st.checkbox("비교 국가 표시", value=False, key="ts_show_compare")
    compare_country = None
    compare_code = None
    if show_compare:
        compare_country = st.selectbox(
            "비교 국가", [c for c in COUNTRIES if c != "한국"],
            index=3, key="ts_compare_country",
        )
        compare_code = COUNTRY_CODES[COUNTRIES.index(compare_country)]

        # 비교국 변화량 계산
        cmp_col = f"{compare_code}_level"
        cat_new_cmp = df_new.groupby("category")[cmp_col].mean().reset_index()
        cat_old_cmp = old_with_new_cat.groupby("new_category")[cmp_col].mean().reset_index().rename(
            columns={"new_category": "category"})
        cmp_merged = cat_new_cmp.merge(cat_old_cmp, on="category", suffixes=("_new", "_old"), how="inner")
        cmp_merged[f"{compare_code}_delta"] = cmp_merged[f"{cmp_col}_new"] - cmp_merged[f"{cmp_col}_old"]
        merged = merged.merge(cmp_merged[["category", f"{compare_code}_delta"]], on="category", how="left")

    # 차트
    fig = go.Figure()

    # 한국 (항상 표시)
    fig.add_trace(go.Bar(
        y=merged["category"], x=merged["kr_delta"],
        name="한국",
        orientation="h",
        marker_color=["#4a7c23" if v >= 0 else "#c0392b" for v in merged["kr_delta"]],
        text=[f"{v:+.1f}%p" for v in merged["kr_delta"]],
        textposition="outside", textfont=dict(size=12),
    ))

    # 비교국 (체크 시)
    if show_compare and compare_code and f"{compare_code}_delta" in merged.columns:
        fig.add_trace(go.Bar(
            y=merged["category"], x=merged[f"{compare_code}_delta"],
            name=compare_country,
            orientation="h",
            marker_color=["rgba(100,150,200,0.7)" if v >= 0 else "rgba(200,100,100,0.5)"
                          for v in merged[f"{compare_code}_delta"]],
            text=[f"{v:+.1f}" for v in merged[f"{compare_code}_delta"]],
            textposition="outside", textfont=dict(size=11),
        ))

    fig.add_vline(x=0, line_dash="solid", line_color="gray", opacity=0.5)
    fig.update_layout(
        title=f"한국 기술수준 변화량 ({old_year}→{new_year}, 38대 기술, %p)",
        barmode="group",
        height=max(600, len(merged) * 32),
        xaxis_title="기술수준 변화 (%p)",
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=180, r=60),
    )
    st.plotly_chart(fig, width="stretch")

    if len(unmapped_new) > 0:
        st.caption(
            f"* 비매핑 기술 ({len(unmapped_new)}개, {new_year}년 신규): "
            f"{', '.join(unmapped_new['category'].tolist()[:10])}"
            f"{'...' if len(unmapped_new) > 10 else ''}"
        )

    # Excel download - 시계열 비교 데이터
    ts_dl_data = pd.DataFrame({
        "국가": COUNTRIES,
        f"{old_year}_기술수준(%)": levels_old,
        f"{new_year}_기술수준(%)": levels_new,
        f"{old_year}_기술격차(년)": gaps_old,
        f"{new_year}_기술격차(년)": gaps_new,
    })
    _excel_download(ts_dl_data, f"시계열비교_{old_year}_{new_year}.xlsx", key="dl_timeseries")


def render_detail():
    """세부기술수준 탭"""

    df_2025, cat_stats, country_stats, _ = load_2025_data()

    # 필터
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
    with col1:
        year = st.selectbox("연도", ["2025", "2020"], key="detail_year")
    with col2:
        nation = st.selectbox("국가선택", COUNTRIES, key="detail_nation")
    with col3:
        type_filter = st.selectbox("구분", ["전체", "감축", "적응"], key="detail_type")
    with col4:
        if type_filter != "전체":
            # Get categories for selected type from data
            if year == "2025":
                available_cats = sorted(df_2025[df_2025["type"] == type_filter]["category_38"].dropna().unique()) if "category_38" in df_2025.columns else []
            else:
                _df_2020_tmp = load_2020_data()
                available_cats = sorted(_df_2020_tmp[_df_2020_tmp["type"] == type_filter]["category"].dropna().unique()) if "category" in _df_2020_tmp.columns else []

            detail_filter = st.selectbox(
                "세부기술 (38대/44대)",
                ["전체"] + list(available_cats),
                key="detail_tech",
            )
        else:
            detail_filter = "전체"
            st.selectbox("세부기술", ["전체"], disabled=True, key="detail_tech")

    # 데이터 선택
    if year == "2025":
        df = df_2025.copy()
        # 2025 격차: 개월→년 변환
        for code in COUNTRY_CODES:
            gap_col = f"{code}_gap"
            if gap_col in df.columns:
                df[gap_col] = df[gap_col] / 12
        if type_filter != "전체":
            df = df[df["type"] == type_filter]
        if detail_filter != "전체" and "category_38" in df.columns:
            df = df[df["category_38"] == detail_filter]
        gap_unit_label = "년"
    else:
        df = load_2020_data()
        if type_filter != "전체":
            df = df[df["type"] == type_filter]
        if detail_filter != "전체" and "category" in df.columns:
            df = df[df["category"] == detail_filter]
        gap_unit_label = "년"

    if len(df) == 0:
        st.warning("해당 조건의 데이터가 없습니다.")
        return

    # 핵심지표 카드 (선택 국가 기준)
    nation_code = COUNTRY_CODES[COUNTRIES.index(nation)]
    nation_lv_col = f"{nation_code}_level"
    nation_gp_col = f"{nation_code}_gap"
    nation_avg_level = df[nation_lv_col].mean() if nation_lv_col in df.columns else 0
    nation_avg_gap = df[nation_gp_col].mean() if nation_gp_col in df.columns else 0

    # 선도기술 수 (선택 국가가 최고기술보유국인 기술 수)
    leading_nation = 0
    if "leading_country" in df.columns:
        leading_nation = len(df[df["leading_country"].str.contains(nation, na=False)])

    # 최우수 분야
    best_tech = ""
    if nation_lv_col in df.columns and len(df) > 0:
        best_idx = df[nation_lv_col].idxmax()
        if pd.notna(best_idx):
            best_tech = df.loc[best_idx, "detail"] if "detail" in df.columns else ""

    c1, c2, c3, c4 = st.columns(4)
    for col, icon, label, value, sub in [
        (c1, "📊", "평균기술수준", f"{nation_avg_level:.1f}%", f"{nation} 기준"),
        (c2, "⏱️", "평균 기술격차", f"{nation_avg_gap:.1f}년", "최고기술국 대비 (작을수록 우수)"),
        (c3, "🏆", "선도기술 수", f"{leading_nation}개", f"{nation}이(가) 최고보유국"),
        (c4, "⭐", "최우수 분야", str(best_tech)[:15], f"{nation} 기술수준 최고"),
    ]:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <h3>{icon} {label}</h3>
                <div class="value">{value}</div>
                <div class="sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")

    # 선택국가 vs 주요국 비교
    st.markdown(f"### {nation} vs 주요국 기술수준 비교")

    col1, col2 = st.columns(2)
    with col1:
        levels = []
        for c, code in zip(COUNTRIES, COUNTRY_CODES):
            lv = df[f"{code}_level"].mean() if f"{code}_level" in df.columns else 0
            levels.append(lv)
        # Highlight selected nation
        bar_colors = []
        for c in COUNTRIES:
            if c == nation:
                bar_colors.append(COUNTRY_COLORS[c])
            else:
                base = COUNTRY_COLORS[c]
                bar_colors.append(f"rgba(180,180,180,0.4)")
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES, y=levels,
            marker_color=bar_colors,
            text=[f"{v:.1f}%" for v in levels], textposition="outside",
        )])
        fig.update_layout(title="기술수준 비교", height=350, yaxis=dict(range=[0, 105]))
        st.plotly_chart(fig, width="stretch")

    with col2:
        gaps = []
        for c, code in zip(COUNTRIES, COUNTRY_CODES):
            gp = df[f"{code}_gap"].mean() if f"{code}_gap" in df.columns else 0
            gaps.append(gp)
        gap_max = max(gaps) if gaps else 1
        gap_colors = []
        for i, c in enumerate(COUNTRIES):
            v = gaps[i]
            if c == nation:
                gap_colors.append(f"rgb({min(255, int(200 * v / gap_max))}, {min(200, int(200 * (1 - v / gap_max)))}, 80)")
            else:
                gap_colors.append("rgba(180,180,180,0.4)")
        fig = go.Figure(data=[go.Bar(
            x=COUNTRIES, y=gaps,
            marker_color=gap_colors,
            text=[f"{v:.1f}" for v in gaps], textposition="outside",
        )])
        fig.update_layout(title="기술격차 비교 ↓작을수록 우수", height=350)
        st.plotly_chart(fig, width="stretch")

    # 상세현황 테이블
    st.markdown("### 전체 기후기술 상세현황")
    st.caption(f"총 {len(df)}건의 세부기술이 등록되어 있습니다.")

    display_cols = []
    if "detail" in df.columns:
        display_cols.append("detail")
    if "type" in df.columns:
        display_cols.append("type")
    for code in COUNTRY_CODES:
        if f"{code}_level" in df.columns:
            display_cols.append(f"{code}_level")
        if f"{code}_gap" in df.columns:
            display_cols.append(f"{code}_gap")
    if "leading_country" in df.columns:
        display_cols.append("leading_country")

    rename = {
        "detail": "세부기술", "type": "구분",
        "kr_level": "한국수준(%)", "kr_gap": "한국격차",
        "cn_level": "중국수준(%)", "cn_gap": "중국격차",
        "jp_level": "일본수준(%)", "jp_gap": "일본격차",
        "us_level": "미국수준(%)", "us_gap": "미국격차",
        "eu_level": "EU수준(%)", "eu_gap": "EU격차",
        "leading_country": "최고보유국",
    }

    display_df = df[display_cols].rename(columns=rename).copy()
    st.dataframe(display_df, width="stretch", height=400)

    # Excel download - 상세 테이블
    _excel_download(display_df, f"기후기술_상세현황_{year}.xlsx", key="dl_detail_table")

    # ── 2개국 비교 분석 ──
    st.markdown("### 기술 비교 분석")
    compare_cols = st.columns(2)
    with compare_cols[0]:
        country1 = st.selectbox("국가 1", COUNTRIES, index=0, key="cmp_c1")
    with compare_cols[1]:
        country2_options = [c for c in COUNTRIES if c != country1]
        country2 = st.selectbox("국가 2", country2_options, index=0, key="cmp_c2")

    # Select a detail technology
    tech_list = df["detail"].dropna().unique().tolist() if "detail" in df.columns else []
    if tech_list:
        selected_tech = st.selectbox("세부기술 선택", tech_list, key="cmp_tech")

        if selected_tech:
            tech_rows = df[df["detail"] == selected_tech]
            if len(tech_rows) > 0:
                tech_row = tech_rows.iloc[0]
                code1 = COUNTRY_CODES[COUNTRIES.index(country1)]
                code2 = COUNTRY_CODES[COUNTRIES.index(country2)]

                lv1 = tech_row.get(f"{code1}_level", 0)
                lv2 = tech_row.get(f"{code2}_level", 0)
                gp1 = tech_row.get(f"{code1}_gap", 0)
                gp2 = tech_row.get(f"{code2}_gap", 0)

                # Gap values already converted to years for 2025 data above
                fig_cmp = go.Figure()
                fig_cmp.add_trace(go.Bar(
                    name=country1,
                    x=["기술수준(%)", "기술격차(년)"],
                    y=[lv1, gp1],
                    marker_color=COUNTRY_COLORS.get(country1, "#999"),
                    text=[f"{lv1:.1f}", f"{gp1:.1f}"],
                    textposition="outside",
                ))
                fig_cmp.add_trace(go.Bar(
                    name=country2,
                    x=["기술수준(%)", "기술격차(년)"],
                    y=[lv2, gp2],
                    marker_color=COUNTRY_COLORS.get(country2, "#999"),
                    text=[f"{lv2:.1f}", f"{gp2:.1f}"],
                    textposition="outside",
                ))
                fig_cmp.update_layout(
                    barmode="group", height=350,
                    title=f"{selected_tech} — {country1} vs {country2}",
                )
                st.plotly_chart(fig_cmp, width="stretch")
    else:
        st.info("비교할 세부기술 데이터가 없습니다.")
