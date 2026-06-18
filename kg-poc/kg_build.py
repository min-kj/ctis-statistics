# -*- coding: utf-8 -*-
"""
정책-데이터 Knowledge Graph PoC  (STEP3-A 구현)
- 정책문서/전략 → 데이터수요 → 법조항(근거) → 데이터셋(보유/외부) → 소관기관
- 그래프 질의로 "법정의무인데 CTis 미보유" 데이터 수요(=갭분석 표) 자동 도출
- 산출: policy_data_kg.html (인터랙티브) + gap_analysis.md
근거: STEP3_KG_타당성분석_정민경.md (기후기술법·기본계획·시행계획 검증)
실행: uv run --with pyvis --with networkx python kg_build.py
"""
import os
import networkx as nx
from pyvis.network import Network

OUT_HTML = os.path.join(os.path.dirname(__file__), "policy_data_kg.html")
OUT_GAP = os.path.join(os.path.dirname(__file__), "gap_analysis.md")

# ---------- 색상 ----------
COL = {
    "law_oblig": "#c0392b",   # 강행/반영의무 법조항
    "law_disc":  "#e67e22",   # 재량 조항
    "strategy":  "#2e6fb0",   # 전략/과제
    "doc":       "#34495e",   # 정책문서
    "d_now":     "#3d8b40",   # 즉시(보유)
    "d_part":    "#d4a017",   # 부분보유
    "d_new":     "#9b59b6",   # 신규(미보유)
    "d_ext":     "#7f8c8d",   # 외부연계
    "ds_ctis":   "#1b706b",   # CTis 데이터셋
    "ds_ext":    "#95a5a6",   # 외부 데이터셋
    "org":       "#16a085",   # 기관
}
STATUS_COL = {"즉시": "d_now", "부분": "d_part", "신규": "d_new", "외부연계": "d_ext"}

# ---------- 노드 정의 ----------
# 법조항 (id, label, 강제력)
laws = [
    ("L_7_1_2", "기후기술법 §7①2호\n인력 현황·수요 전망", "재량"),
    ("L_8_2",   "기후기술법 §8②\n산업계 기술수요·기술발전 예측", "반영의무"),
    ("L_8_3",   "기후기술법 §8③\n기술지도 작성", "강행의무"),
]
# 전략/과제
strategies = [
    ("S_3_4", "전략3-4 거버넌스·정책역량\n(CTis·활동조사·기술지도)"),
    ("S_3_2", "전략3-2 현장맞춤형 인재양성"),
    ("S_1",   "전략1 온실가스 감축"),
    ("S_2",   "전략2 기후변화 적응"),
]
docs = [
    ("DOC_BASIC", "제1차 기본계획('23~'32)\n3전략·15세부·85과제"),
    ("DOC_IMPL",  "시행계획('25)\n2.7조 예산·적응 2,758억"),
]
# 데이터셋(보유/외부)
datasets = [
    ("DS_LEVEL",  "CTis 기술수준조사(38대)", "ctis"),
    ("DS_ACT",    "CTis 활동조사(22대)", "ctis"),
    ("DS_RND",    "CTis R&D 투자·성과(38대)", "ctis"),
    ("DS_MAP",    "NIGT 분류 매핑표(38↔45↔100↔22)", "ctis"),
    ("DS_FIN",    "한신평 기업 재무데이터", "ext"),
    ("DS_GIR",    "환경부 온실가스 인벤토리(GIR)", "ext"),
]
orgs = [
    ("ORG_NIGT", "NIGT 데이터정보센터"),
    ("ORG_ENV",  "환경부/GIR"),
    ("ORG_FIN",  "한신평"),
]
# 데이터수요 (id, label, status, 근거 법조항, 충족 데이터셋들, 부분 데이터셋)
demands = [
    ("D_LEVEL", "기술수준(38대·5개국)",      "즉시",     None,      ["DS_LEVEL"], None),
    ("D_RND",   "R&D 투자(38대·부처/재원)",   "즉시",     None,      ["DS_RND"],   None),
    ("D_PERF",  "성과(논문·특허)",           "즉시",     None,      ["DS_RND"],   None),
    ("D_IND",   "산업현황(매출·종사자 22대)", "즉시",     None,      ["DS_ACT"],   None),
    ("D_HR",    "인력 현황·수요전망(성별·기능별)", "부분", "L_7_1_2", [],        ["DS_ACT"]),
    ("D_NEED",  "산업계 기술수요",            "신규",     "L_8_2",   [],           None),
    ("D_FORE",  "기술발전 예측결과",          "신규",     "L_8_2",   [],           None),
    ("D_MAP",   "기후기술 기술지도(로드맵)",   "신규",     "L_8_3",   [],           None),
    ("D_X15",   "38대×15세부전략 매핑",       "신규",     None,      [],           ["DS_MAP"]),
    ("D_FIN",   "기업 재무·산업 데이터",       "외부연계", None,      ["DS_FIN"],   None),
    ("D_GHG",   "온실가스 배출 인벤토리",      "외부연계", None,      ["DS_GIR"],   None),
    ("D_STAT",  "38대 국가승인통계",          "신규",     None,      [],           None),
]
# 전략 → 데이터수요 (요구)
requires = [
    ("S_3_4", "D_LEVEL"), ("S_3_4", "D_RND"), ("S_3_4", "D_PERF"), ("S_3_4", "D_IND"),
    ("S_3_4", "D_MAP"), ("S_3_4", "D_STAT"), ("S_3_4", "D_X15"),
    ("S_3_2", "D_HR"), ("S_3_4", "D_NEED"), ("S_3_4", "D_FORE"),
    ("S_1", "D_GHG"), ("S_1", "D_RND"), ("S_3_4", "D_FIN"),
]
doc_has = [("DOC_BASIC", "S_3_4"), ("DOC_BASIC", "S_3_2"), ("DOC_BASIC", "S_1"),
           ("DOC_BASIC", "S_2"), ("DOC_IMPL", "S_3_4")]
ds_org = [("DS_LEVEL", "ORG_NIGT"), ("DS_ACT", "ORG_NIGT"), ("DS_RND", "ORG_NIGT"),
          ("DS_MAP", "ORG_NIGT"), ("DS_GIR", "ORG_ENV"), ("DS_FIN", "ORG_FIN")]

# ---------- 그래프 구축 ----------
G = nx.DiGraph()
for nid, lab, force in laws:
    G.add_node(nid, label=lab, kind="law", force=force)
for nid, lab in strategies:
    G.add_node(nid, label=lab, kind="strategy")
for nid, lab in docs:
    G.add_node(nid, label=lab, kind="doc")
for nid, lab, kind in datasets:
    G.add_node(nid, label=lab, kind="dataset", dskind=kind)
for nid, lab in orgs:
    G.add_node(nid, label=lab, kind="org")
for nid, lab, status, basis, sat, part in demands:
    G.add_node(nid, label=lab, kind="demand", status=status, basis=basis, sat=sat, part=part)
    if basis:
        G.add_edge(nid, basis, label="근거")
    for ds in sat:
        G.add_edge(nid, ds, label="충족" if status != "외부연계" else "외부연계")
    if part:
        for ds in part:
            G.add_edge(nid, ds, label="부분")
for s, d in requires:
    G.add_edge(s, d, label="요구")
for a, b in doc_has:
    G.add_edge(a, b, label="포함")
for ds, og in ds_org:
    G.add_edge(ds, og, label="소관")

# ---------- 갭분석 질의 ----------
# 법정 근거(의무/재량)가 있고 CTis 충족 데이터셋이 없는 데이터수요
def is_ctis(ds_id):
    return G.nodes[ds_id].get("dskind") == "ctis"

gap_rows = []
for nid, lab, status, basis, sat, part in demands:
    has_ctis = any(is_ctis(d) for d in sat)
    if not has_ctis:
        force = G.nodes[basis]["force"] if basis else "-"
        basis_lab = G.nodes[basis]["label"].split("\n")[0] if basis else "-"
        part_lab = ", ".join(G.nodes[d]["label"] for d in (part or [])) or "-"
        gap_rows.append((lab, status, basis_lab, force, part_lab))

# 우선순위: 강행의무 > 반영의무 > 재량 > 무
order = {"강행의무": 0, "반영의무": 1, "재량": 2, "-": 3}
gap_rows.sort(key=lambda r: order.get(r[3], 9))

with open(OUT_GAP, "w", encoding="utf-8") as f:
    f.write("# 갭분석 (KG 질의 자동 도출) — 법정 근거 있으나 CTis 미보유\n\n")
    f.write("| 데이터 수요 | 상태 | 법적 근거 | 강제력 | 부분보유 |\n|---|---|---|---|---|\n")
    for lab, status, basis_lab, force, part_lab in gap_rows:
        f.write(f"| {lab} | {status} | {basis_lab} | {force} | {part_lab} |\n")
    f.write(f"\n> 총 {len(gap_rows)}건. 강행의무(§8③ 기술지도) 최우선.\n")

print("=== 갭분석 (CTis 미보유) ===")
for lab, status, basis_lab, force, part_lab in gap_rows:
    print(f"  [{force:5}] {lab}  ({status})  근거={basis_lab}")

# ---------- 시각화 (pyvis) ----------
net = Network(height="780px", width="100%", directed=True, bgcolor="#ffffff", font_color="#222", cdn_resources="in_line")
net.barnes_hut(gravity=-9000, central_gravity=0.3, spring_length=130)
for nid, data in G.nodes(data=True):
    kind = data["kind"]
    if kind == "law":
        color = COL["law_oblig"] if data["force"] in ("강행의무", "반영의무") else COL["law_disc"]
        shape, size = "box", 18
    elif kind == "strategy":
        color, shape, size = COL["strategy"], "box", 16
    elif kind == "doc":
        color, shape, size = COL["doc"], "box", 20
    elif kind == "dataset":
        color = COL["ds_ctis"] if data["dskind"] == "ctis" else COL["ds_ext"]
        shape, size = "ellipse", 16
    elif kind == "org":
        color, shape, size = COL["org"], "diamond", 14
    else:  # demand
        color, shape, size = COL[STATUS_COL[data["status"]]], "dot", 22
    title = f"{data['label']}" + (f"<br>상태: {data.get('status')}" if kind == "demand" else "")
    net.add_node(nid, label=data["label"], color=color, shape=shape, size=size, title=title)
for s, d, ed in G.edges(data=True):
    lab = ed.get("label", "")
    ecol = "#c0392b" if lab in ("근거",) else ("#bbb" if lab in ("소관", "포함") else "#888")
    dashes = lab in ("부분", "외부연계")
    net.add_edge(s, d, label=lab, color=ecol, dashes=dashes, font={"size": 9, "color": "#777"})

net.set_options('{"edges":{"smooth":{"type":"dynamic"},"arrows":{"to":{"enabled":true,"scaleFactor":0.6}}},"physics":{"stabilization":{"iterations":200}}}')
net.save_graph(OUT_HTML)
print(f"\nKG HTML: {OUT_HTML}")
print(f"갭분석:  {OUT_GAP}")
print(f"노드 {G.number_of_nodes()} · 엣지 {G.number_of_edges()} · 갭 {len(gap_rows)}건")
