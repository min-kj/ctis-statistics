# -*- coding: utf-8 -*-
"""KG PoC 인벤토리 덤프 — 실제 그래프 객체에서 노드/엣지 전부 출력"""
from collections import defaultdict
import kg_build as k  # import 시 그래프 G 빌드됨

G = k.G
order = ["doc", "strategy", "law", "demand", "dataset", "org"]
kname = {"doc": "정책문서", "strategy": "전략/과제", "law": "법조항",
         "demand": "데이터수요", "dataset": "데이터셋", "org": "기관"}

L = [f"# KG PoC 인벤토리 — 노드 {G.number_of_nodes()} · 엣지 {G.number_of_edges()}"]
for kind in order:
    ns = [(n, d) for n, d in G.nodes(data=True) if d["kind"] == kind]
    L.append(f"\n## {kname[kind]} ({len(ns)})")
    for n, d in ns:
        lab = d["label"].replace("\n", " / ")
        ex = ""
        if kind == "law":
            ex = f"  →  {d.get('force','')}"
        elif kind == "demand":
            ex = f"  →  [{d.get('status','')}]"
        elif kind == "dataset":
            ex = "  (CTis)" if d.get("dskind") == "ctis" else "  (외부)"
        L.append(f"- {lab}{ex}")

eg = defaultdict(list)
for s, t, d in G.edges(data=True):
    eg[d.get("label", "")].append(
        (G.nodes[s]["label"].split("\n")[0], G.nodes[t]["label"].split("\n")[0]))
L.append(f"\n## 엣지 ({G.number_of_edges()})")
for lab in ["포함", "요구", "근거", "충족", "외부연계", "부분", "소관"]:
    es = eg.get(lab, [])
    if not es:
        continue
    L.append(f"\n### [{lab}] {len(es)}건")
    for s, t in es:
        L.append(f"- {s}  →  {t}")

out = "\n".join(L)
with open("kg_inventory.md", "w", encoding="utf-8") as f:
    f.write(out)
print(out)
