# -*- coding: utf-8 -*-
"""
PLAN-020 국가별 선도 지형 목업 — 데이터 추출기
=================================================
src/db/ctis_stats.db (2025 수준조사, 38대/157개 세부기술 × 5개국)에서
country-leadership-landscape.html 목업이 쓰는 data.json 을 생성한다.

- 1차 소스: SQLite DB (survey_result + tech_detail + tech_category)
- 국가 색상/국기: src/dashboard/data_loader.py:21-29 의 COUNTRY_COLORS / COUNTRY_FLAGS
  를 그대로 미러링(단일 소스 원칙). 표준 라이브러리만 사용 → `uv run --no-project python` 으로 실행 가능.

실행:  cd ctis-stats && uv run --no-project python mockups/extract_data.py
출력:  mockups/data.json  (UTF-8, ensure_ascii=False)
"""
import base64
import json
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

# ── 경로 (스크립트 기준 절대경로 → 한글 경로/실행위치 무관) ─────────────────
# 위치: ctis-stats/src/dashboard/infographics/extract_data.py
ROOT = Path(__file__).resolve().parents[3]             # ctis-stats/
DB_PATH = ROOT / "src" / "db" / "ctis_stats.db"
HERE = Path(__file__).resolve().parent                 # infographics/
OUT_PATH = HERE / "data.json"
FLAG_DIR = HERE / "flags"

# ── 국가 메타 (data_loader.py COUNTRY_COLORS/FLAGS 미러) ────────────────────
COUNTRY_META = {
    "kr": {"name": "한국", "flag": "🇰🇷", "color": "#FF6B6B"},
    "cn": {"name": "중국", "flag": "🇨🇳", "color": "#4ECDC4"},
    "jp": {"name": "일본", "flag": "🇯🇵", "color": "#45B7D1"},
    "us": {"name": "미국", "flag": "🇺🇸", "color": "#96CEB4"},
    "eu": {"name": "EU",  "flag": "🇪🇺", "color": "#FECA57"},
}
CODES = list(COUNTRY_META)                     # [kr, cn, jp, us, eu]
DB2LC = {"KR": "kr", "CN": "cn", "JP": "jp", "US": "us", "EU": "eu"}

# 감축/적응 표식 색 (기존 페이지 관례)
TYPE_COLOR = {"감축": "#4a7c23", "적응": "#2196F3"}


def _clean_cat(name: str) -> str:
    """'01. 태양광·열' → '태양광·열' (번호 접두 제거)."""
    import re
    return re.sub(r"^\s*\d+[.\s]*", "", str(name)).strip()


def _round(x, n=1):
    return None if x is None else round(float(x), n)


def _flag_uri(code):
    """flags/{code}.png → base64 data URI (data.js 인라인만으로 국기까지 자체완결)."""
    p = FLAG_DIR / f"{code}.png"
    if not p.exists():
        return ""
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def build_trend(db_path):
    """2020 → 2025 변화: 국가별 평균수준/격차 추이 + 1:1 매핑 기반 한국 분야별 변화."""
    con = sqlite3.connect(db_path); con.row_factory = sqlite3.Row; cur = con.cursor()

    # 국가별 연도 평균 (격차는 년 단위로 정규화: 2020=년, 2025=개월)
    countries = {c: {} for c in CODES}
    for y in (2020, 2025):
        for r in cur.execute(
            "SELECT country_code, AVG(tech_level) lv, AVG(tech_gap) gp, MIN(gap_unit) gu "
            "FROM survey_result WHERE survey_year=? GROUP BY country_code", (y,)):
            lc = DB2LC.get(r["country_code"])
            if not lc:
                continue
            gp = r["gp"]
            gp_yr = gp if r["gu"] == "year" else (gp / 12.0 if gp is not None else None)
            countries[lc][str(y)] = {"level": _round(r["lv"]), "gap_yr": _round(gp_yr, 2)}
    for lc, d in countries.items():
        if "2020" in d and "2025" in d:
            d["d_level"] = _round(d["2025"]["level"] - d["2020"]["level"])
            d["d_gap"] = _round(d["2025"]["gap_yr"] - d["2020"]["gap_yr"], 2)

    # 한국 분야별 변화 (1:1 confirmed 매핑만 — 44대 src ↔ 38대 tgt 정합)
    krlvl = {}  # (year, category_id) → KR 평균수준
    for r in cur.execute(
        "SELECT sr.survey_year y, td.category_id cid, AVG(sr.tech_level) lv "
        "FROM survey_result sr JOIN tech_detail td ON sr.detail_id=td.detail_id "
        "WHERE sr.country_code='KR' GROUP BY sr.survey_year, td.category_id"):
        krlvl[(r["y"], r["cid"])] = r["lv"]
    catname = {r["category_id"]: r["category_name"]
               for r in cur.execute("SELECT category_id, category_name FROM tech_category")}
    movers = []
    for m in cur.execute("SELECT src_category_id s, tgt_category_id t FROM taxonomy_mapping "
                         "WHERE mapping_type='1:1' AND mapping_status='confirmed'"):
        l20, l25 = krlvl.get((2020, m["s"])), krlvl.get((2025, m["t"]))
        if l20 is None or l25 is None:
            continue
        movers.append({"name": _clean_cat(catname.get(m["t"], "")),
                       "y2020": _round(l20), "y2025": _round(l25),
                       "delta": _round(l25 - l20)})
    con.close()
    movers.sort(key=lambda x: x["delta"])  # 하락 → 상승
    return {"years": [2020, 2025], "countries": countries, "kr_movers": movers}


def main():
    if not DB_PATH.exists():
        sys.exit(f"[ERROR] DB 없음: {DB_PATH}\n  먼저 src/db/init_db.py 로 DB를 구축하세요.")

    con = sqlite3.connect(str(DB_PATH))
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    rows = cur.execute(
        """
        SELECT sr.detail_id, sr.country_code, sr.tech_level, sr.tech_gap,
               sr.gap_unit, sr.is_leading,
               td.detail_no, td.detail_name,
               tc.category_no, tc.category_name, tc.type_name
        FROM survey_result sr
        JOIN tech_detail   td ON sr.detail_id  = td.detail_id
        JOIN tech_category tc ON td.category_id = tc.category_id
        WHERE sr.survey_year = 2025
        ORDER BY tc.category_no, td.detail_no
        """
    ).fetchall()
    con.close()

    # ── 세부기술 단위로 5개국 묶기 ─────────────────────────────────────────
    details = {}   # detail_id → dict
    order = []     # category 등장 순서 유지
    cat_meta = {}  # cat_no → {name, type}
    gap_unit = "month"

    for r in rows:
        lc = DB2LC.get(r["country_code"])
        if lc is None:
            continue
        did = r["detail_id"]
        if did not in details:
            details[did] = {
                "detail_id": did,
                "no": r["detail_no"],
                "name": r["detail_name"],
                "cat_no": r["category_no"],
                "levels": {}, "gaps": {}, "leader": None,
            }
        d = details[did]
        d["levels"][lc] = _round(r["tech_level"])
        d["gaps"][lc] = _round(r["tech_gap"], 2)
        if r["is_leading"] == 1:
            d["leader"] = lc
        if r["gap_unit"]:
            gap_unit = r["gap_unit"]
        cn = r["category_no"]
        if cn not in cat_meta:
            cat_meta[cn] = {"name": r["category_name"], "type": r["type_name"]}
            order.append(cn)

    # leader 보정: is_leading 누락 시 최고 수준국으로
    for d in details.values():
        if d["leader"] is None and d["levels"]:
            d["leader"] = max(d["levels"], key=lambda c: d["levels"].get(c) or -1)

    # ── 분야(38대) 집계 ────────────────────────────────────────────────────
    by_cat = defaultdict(list)
    for d in details.values():
        by_cat[d["cat_no"]].append(d)

    categories = []
    for cn in order:
        ds = sorted(by_cat[cn], key=lambda x: x["no"])
        # 분야별 5개국 평균 수준
        avg_levels = {}
        for c in CODES:
            vals = [d["levels"].get(c) for d in ds if d["levels"].get(c) is not None]
            avg_levels[c] = _round(sum(vals) / len(vals)) if vals else None
        # 분야 선도국 = 세부기술 선도 최다국 (동률 → 평균수준 높은국)
        lead_cnt = defaultdict(int)
        for d in ds:
            if d["leader"]:
                lead_cnt[d["leader"]] += 1
        if lead_cnt:
            leader = max(lead_cnt, key=lambda c: (lead_cnt[c], avg_levels.get(c) or -1))
        else:
            leader = max(avg_levels, key=lambda c: avg_levels.get(c) or -1)
        categories.append({
            "no": cn,
            "name": _clean_cat(cat_meta[cn]["name"]),
            "name_full": cat_meta[cn]["name"],
            "type": cat_meta[cn]["type"],
            "leader": leader,
            "lead_counts": dict(lead_cnt),   # 분야 내 국가별 선도 세부기술 수
            "levels": avg_levels,
            "n_detail": len(ds),
            "details": [
                {"no": d["no"], "name": d["name"], "leader": d["leader"],
                 "levels": d["levels"], "gaps": d["gaps"]}
                for d in ds
            ],
        })

    # ── 국가별 요약 ────────────────────────────────────────────────────────
    country_summary = []
    for c in CODES:
        leading_count = sum(1 for d in details.values() if d["leader"] == c)
        leading_cats = sum(1 for cat in categories if cat["leader"] == c)
        all_lv = [d["levels"].get(c) for d in details.values() if d["levels"].get(c) is not None]
        avg_level = _round(sum(all_lv) / len(all_lv)) if all_lv else None
        all_gp = [d["gaps"].get(c) for d in details.values() if d["gaps"].get(c) is not None]
        avg_gap = _round(sum(all_gp) / len(all_gp), 1) if all_gp else None
        # 강·약점 분야 = 해당국 분야평균수준 상/하위 3
        ranked = sorted(
            [cat for cat in categories if cat["levels"].get(c) is not None],
            key=lambda cat: cat["levels"][c], reverse=True,
        )
        strong = [{"name": cat["name"], "level": cat["levels"][c]} for cat in ranked[:3]]
        weak = [{"name": cat["name"], "level": cat["levels"][c]} for cat in ranked[-3:][::-1]]
        country_summary.append({
            "code": c, **COUNTRY_META[c],
            "leading_count": leading_count,
            "leading_categories": leading_cats,
            "avg_level": avg_level,
            "avg_gap": avg_gap,
            "strong": strong, "weak": weak,
        })
    country_summary.sort(key=lambda x: x["leading_count"], reverse=True)

    out = {
        "meta": {
            "year": 2025,
            "title": "기후기술 국가별 선도 지형 2025",
            "subtitle": "기후기술수준조사 · 38대 분야 / 157개 세부기술 · 한·중·일·미·EU 비교",
            "source": "국가녹색기술연구소(NIGT) 2025 기후기술수준조사 (2차 델파이)",
            "n_detail": len(details),
            "n_category": len(categories),
            "gap_unit": gap_unit,
            "countries": [{"code": c, **COUNTRY_META[c], "img": _flag_uri(c)} for c in CODES],
            "type_color": TYPE_COLOR,
        },
        "categories": categories,
        "country_summary": country_summary,
        "trend": build_trend(str(DB_PATH)),
    }

    payload = json.dumps(out, ensure_ascii=False, indent=2)
    OUT_PATH.write_text(payload, encoding="utf-8")
    # file:// 더블클릭으로도 열리도록 data.js(인라인) 동시 출력 (fetch는 file://에서 CORS 차단됨)
    (OUT_PATH.parent / "data.js").write_text(
        "window.CTIS_DATA = " + payload + ";\n", encoding="utf-8")

    # ── 검증 리포트 ────────────────────────────────────────────────────────
    tot_lead = sum(cs["leading_count"] for cs in country_summary)
    print(f"✓ data.json 생성: {OUT_PATH}")
    print(f"  분야 {out['meta']['n_category']}개 / 세부기술 {out['meta']['n_detail']}개")
    print(f"  선도기술 합계: {tot_lead} (== {len(details)} 이어야 함)")
    print("  국가별 선도기술 수:", {cs["code"]: cs["leading_count"] for cs in country_summary})
    print("  국가별 선도분야 수:", {cs["code"]: cs["leading_categories"] for cs in country_summary})
    # 결측 점검
    miss = [(cat["no"], c) for cat in categories for c in CODES if cat["levels"].get(c) is None]
    print(f"  분야평균 결측: {len(miss)}건")
    assert out["meta"]["n_category"] == 38, "38대 분야 수 불일치"
    assert out["meta"]["n_detail"] == 157, "157개 세부기술 수 불일치"
    assert tot_lead == len(details), "선도기술 합계 != 세부기술 수"
    print("  ✅ 검증 통과")


if __name__ == "__main__":
    main()