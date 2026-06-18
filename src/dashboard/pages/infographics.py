# 🐜 인포그래픽 페이지 — 국가별 기후기술 선도 현황(HTML 목업 임베드)
"""
ctis-stats/src/dashboard/infographics/ 의 자체완결형 HTML 목업을
data.js(국기 base64 포함)만 인라인해 st.components.html 로 표시한다.
정적 서빙/상대경로 의존 없이 동작.
"""
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

# pages/infographics.py → dashboard/ → infographics/
DIR = Path(__file__).resolve().parent.parent / "infographics"

# 라벨 → (파일명, 임베드 높이)
MOCKUPS = {
    "국가별 선도 지형": ("country-leadership-landscape-v2.html", 1750),
    "국가 지표 카드": ("country-profile-cards.html", 1500),
    "2020→2025 변화": ("country-trend-2020-2025.html", 1650),
}

DESC = {
    "국가별 선도 지형": "38대 분야를 한국과의 기술격차(색)·선도국(국기)으로 표현. 분야 클릭 시 세부기술 비교.",
    "국가 지표 카드": "선택 국가의 평균수준·선도기술·격차 등 지표와 5개국 중 순위.",
    "2020→2025 변화": "5년간 기술수준 변화(슬로프그래프) + 한국 분야별 추격·후퇴.",
}


@st.cache_data(show_spinner=False)
def _inline(name: str) -> str:
    """HTML 의 <script src="data.js"> 를 data.js 내용으로 인라인해 자체완결 문서로."""
    html = (DIR / name).read_text(encoding="utf-8")
    datajs = (DIR / "data.js").read_text(encoding="utf-8")
    return html.replace('<script src="data.js"></script>', f"<script>{datajs}</script>")


def render():
    st.markdown("#### 국가별 기후기술 인포그래픽")
    label = st.radio(
        "보기 선택",
        list(MOCKUPS),
        horizontal=True,
        label_visibility="collapsed",
    )
    st.caption(DESC[label])

    name, height = MOCKUPS[label]
    path = DIR / name
    if not path.exists():
        st.error(
            f"목업 파일이 없습니다: `{path}`\n\n"
            "먼저 데이터를 생성하세요:\n"
            "`cd ctis-stats && uv run --no-project python src/dashboard/infographics/extract_data.py`"
        )
        return

    components.html(_inline(name), height=height, scrolling=True)
