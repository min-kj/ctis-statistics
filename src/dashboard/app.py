# 🐜 Scout: PLAN-020 CTis 기후기술통계 대시보드 - 2026-03-20
"""
CTis 기후기술 수준조사 대시보드 (Streamlit)
- 전체 탭: 5개국 비교 + 시계열(2020↔2025)
- 세부기술수준 탭: 핵심지표 + 상세 테이블
- 포지셔닝 매트릭스 탭: 38대 기술 Scatter Plot
"""
import streamlit as st

st.set_page_config(
    page_title="CTis 기후기술통계",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# CTis 스타일 CSS
st.markdown("""
<style>
    /* CTis 헤더 스타일 */
    .ctis-header {
        background: linear-gradient(135deg, #2d5016 0%, #4a7c23 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 0;
        margin: -1rem -1rem 2rem -1rem;
    }
    .ctis-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .ctis-header p { color: #d4e8c0; margin: 0.3rem 0 0 0; font-size: 0.9rem; }

    /* 핵심지표 카드 */
    .metric-card {
        background: linear-gradient(135deg, #2d5016 0%, #3d6b1e 100%);
        color: white;
        padding: 1.2rem;
        border-radius: 10px;
        text-align: center;
    }
    .metric-card h3 { font-size: 0.85rem; margin: 0; opacity: 0.9; }
    .metric-card .value { font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0; }
    .metric-card .sub { font-size: 0.75rem; opacity: 0.7; }

    /* 종합 평가 박스 */
    .eval-box {
        background: #f8faf5;
        border-left: 4px solid #4a7c23;
        padding: 1rem 1.5rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }

    /* 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] { gap: 0; }
    .stTabs [data-baseweb="tab"] {
        padding: 0.8rem 1.5rem;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# 페이지 import
from pages import level_survey, positioning, activity_survey


def main():
    # CTis 헤더
    st.markdown("""
    <div class="ctis-header">
        <h1>기후기술통계</h1>
        <p>홈 > 기후기술통계</p>
    </div>
    """, unsafe_allow_html=True)

    # 2단 메뉴 구조
    menu_level1, menu_level2 = st.tabs([
        "통계정보",
        "기후기술통계 분석",
    ])

    # ── 1단: 통계정보 ──
    with menu_level1:
        sub_tab1, sub_tab2 = st.tabs([
            "기후기술수준조사",
            "기술개발활동조사",
        ])
        with sub_tab1:
            level_survey.render()
        with sub_tab2:
            activity_survey.render()

    # ── 2단: 기후기술통계 분석 ──
    with menu_level2:
        positioning.render()


if __name__ == "__main__":
    main()
