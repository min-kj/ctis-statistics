# CTis 기후기술통계 DB 논리 설계(안)

<!-- 🐜 Scout: PLAN-020 DB 논리설계 - 2026-03-20 -->

## 1. 설계 원칙

| 원칙 | 설명 |
|------|------|
| **38대 기준 통합** | 38대 기후기술을 공통 참조 키(Common Key)로 사용 |
| **분류체계 독립** | 매핑 테이블로 44대↔38대↔22대↔100대 유연 전환 |
| **원본 보존** | 2025부터 세부기술(157개) 단위 원본 저장, 집계는 뷰/캐시로 |
| **시계열 확장** | survey_year 기준으로 2020, 2025, 2030… 자연 확장 |
| **응답 로데이터** | 2025부터 개별 응답자 로데이터 보존 (향후 분석 근거) |

---

## 2. 전체 구조 (ERD 개요)

```
┌─────────────────────────────────────────────────────────────────┐
│                        참조 테이블 (Reference)                    │
│                                                                 │
│  ┌──────────────┐  ┌───────────────┐  ┌──────────────────────┐ │
│  │ taxonomy     │  │ tech_category │  │ tech_detail          │ │
│  │ (분류체계)    │  │ (중분류)       │  │ (세부기술)            │ │
│  └──────┬───────┘  └───────┬───────┘  └──────────┬───────────┘ │
│         │                  │                      │             │
│  ┌──────┴──────────────────┴──────────────────────┘             │
│  │                                                              │
│  │  ┌─────────────────┐                                        │
│  └──│ taxonomy_mapping │ ← 44↔38↔22↔100 교차 매핑              │
│     └─────────────────┘                                        │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                     데이터 테이블 (Data)                          │
│                                                                 │
│  ┌─────────────────────────┐  ┌────────────────────────────┐   │
│  │ survey_result           │  │ survey_response_raw        │   │
│  │ (집계 결과: 157x5국가)   │  │ (로데이터: 응답자별 원시)    │   │
│  │ 2020: 185건 (44대)      │  │ 2025~: 14,856건            │   │
│  │ 2025: 157건 (38대)      │  │ (2020: 없음)               │   │
│  │                         │  │ 원본: 2차_최종DATA_(3차포함) │   │
│  └─────────────────────────┘  └────────────────────────────┘   │
│  ※ 데이터 계보: L0 Raw(1차/2차 개인응답) → L1 집계(통계_1119)    │
│    → L3 공표(★보고서버전_최종, 100점 환산) → L4 보고서PDF         │
│    상세: docs/2025-수준조사-데이터변환-검증보고서.md               │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │ activity_survey                                         │   │
│  │ (활동조사: 22대 × 기관규모 × 지표, 2022~2024 3개년)      │   │
│  │ ⚠️ 가중/미가중 이슈 확인 중                               │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                   집계 뷰 (Aggregation Views)                    │
│                                                                 │
│  v_category_summary    ← 중분류별 집계 (38대/44대)               │
│  v_country_summary     ← 국가별 전체 평균                        │
│  v_type_summary        ← 감축/적응 분류별 집계                    │
│  v_timeseries_compare  ← 2020↔2025 시계열 비교                   │
│  v_positioning_matrix  ← 포지셔닝 매트릭스용 축 데이터             │
└─────────────────────────────────────────────────────────────────┘
```

---

## 3. 테이블 상세 정의

### 3.1 taxonomy (분류체계)

분류체계의 버전을 관리. 연도별로 분류 기준이 바뀌어도 독립적으로 추적.

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `taxonomy_id` | PK, INT | 분류체계 ID | 1 |
| `code` | VARCHAR(20) | 분류체계 코드 | `T44`, `T38`, `T22`, `T100` |
| `name` | VARCHAR(100) | 분류체계 명칭 | `NIGT 44대 기후기술(2020)` |
| `category_count` | INT | 중분류 수 | 44, 38, 22, 100 |
| `detail_count` | INT | 세부기술 수 (있을 경우) | 185, 157, NULL |
| `effective_year` | INT | 적용 시작 연도 | 2020 |
| `source` | VARCHAR(200) | 근거 법령/기준 | `기후기술촉진법 시행규칙` |
| `is_current` | BOOLEAN | 현행 여부 | TRUE |

```
초기 데이터:
T44  | NIGT 44대 기후기술          | 44  | 185 | 2020 | NIGT 기술수준조사 2020
T38  | 38대 기후기술(기후기술법)     | 38  | 157 | 2025 | 기후기술촉진법 시행규칙 고시 2022
T22  | 22대 승인통계 분류           | 22  | -   | 2022 | 국가승인통계 기술개발활동조사
T100 | 탄소중립 100대 핵심기술      | 100 | -   | 2022 | 탄소중립 100대 핵심기술(2024 개정)
```

---

### 3.2 tech_category (중분류)

각 분류체계의 중분류 기술 항목.

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `category_id` | PK, INT | 중분류 ID | 1 |
| `taxonomy_id` | FK → taxonomy | 분류체계 | 2 (T38) |
| `category_no` | VARCHAR(10) | 번호 | `01` |
| `category_name` | VARCHAR(200) | 기술명 | `태양광 기술` |
| `type_code` | VARCHAR(10) | 대분류 코드 | `MIT` / `ADP` |
| `type_name` | VARCHAR(20) | 대분류명 | `감축` / `적응` |
| `definition` | TEXT | 기술 정의 | ... |

```
예시 (T38):
01 | 태양광 기술      | MIT | 감축
02 | 태양열 기술      | MIT | 감축
...
33 | 기후변화 감시·예측 | ADP | 적응
```

---

### 3.3 tech_detail (세부기술)

가장 낮은 단위. 157개(2025) / 185개(2020).

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `detail_id` | PK, INT | 세부기술 ID | 1 |
| `category_id` | FK → tech_category | 소속 중분류 | 1 (태양광) |
| `detail_no` | VARCHAR(10) | 번호 | `01-01` |
| `detail_name` | VARCHAR(300) | 세부기술명 | `건물형 및 입지 다변형 태양광 기술` |
| `survey_year` | INT | 조사 연도 | 2025 |

---

### 3.4 taxonomy_mapping (분류체계 간 매핑)

**핵심 테이블**. 44대↔38대↔22대↔100대 교차 매핑을 유연하게 지원.

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `mapping_id` | PK, INT | 매핑 ID | 1 |
| `src_category_id` | FK → tech_category | 원본 중분류 | 5 (T44의 `태양에너지`) |
| `tgt_category_id` | FK → tech_category | 대상 중분류 | 1 (T38의 `태양광 기술`) |
| `mapping_type` | VARCHAR(20) | 매핑 유형 | `1:1`, `N:1`, `1:N`, `split` |
| `mapping_status` | VARCHAR(20) | 상태 | `confirmed`, `pending`, `deleted` |
| `note` | TEXT | 비고 | `태양에너지 → 태양광+태양열 분리` |

```
매핑 예시:
T44:태양에너지 → T38:태양광 기술  (1:N, confirmed)
T44:태양에너지 → T38:태양열 기술  (1:N, confirmed)
T38:태양광 기술 → T22:신재생에너지  (N:1, confirmed)
T38:태양광 기술 → T100:차세대태양전지 (1:N, confirmed)
```

> **설계 포인트**: 세부기술 단위 매핑이 필요한 경우 `detail_mapping` 테이블 추가 가능.
> 현재는 중분류 매핑이면 충분 (sfr02-update.md A-1 요구사항)

---

### 3.5 survey_result (수준조사 집계 결과) ★ 핵심

**세부기술(157개) × 국가(5개)** 단위의 최종 집계 데이터.
2020년은 이 테이블이 원본(로데이터 없음), 2025년부터는 로데이터에서 집계된 결과.

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `result_id` | PK, INT | 결과 ID | 1 |
| `survey_year` | INT | 조사 연도 | 2025 |
| `detail_id` | FK → tech_detail | 세부기술 | 1 |
| `country_code` | VARCHAR(5) | 국가 코드 | `KR`, `CN`, `JP`, `US`, `EU` |
| `tech_level` | DECIMAL(5,2) | 기술수준 (%) | 85.5 |
| `tech_gap` | DECIMAL(5,2) | 기술격차 | 3.2 (2020:년, 2025:개월) |
| `gap_unit` | VARCHAR(5) | 격차 단위 | `year` / `month` |
| `tech_group` | VARCHAR(10) | 기술그룹 | `선도`, `추격`, `후발` |
| `basic_research` | DECIMAL(3,1) | 기초연구 역량 (점) | 4.2 |
| `applied_research` | DECIMAL(3,1) | 응용개발연구 역량 (점) | 3.8 |
| `rd_trend` | VARCHAR(10) | R&D 동향 | `활발`, `보합`, `정체` |
| `is_leading` | BOOLEAN | 최고보유국 여부 | TRUE |

```
행 수 추정:
  2020: 185 세부기술 × 5 국가 = 925건
  2025: 157 세부기술 × 5 국가 = 785건
  합계: ~1,710건

인덱스: (survey_year, detail_id, country_code) UNIQUE
```

> **2020 데이터**: `수준조사 데이터(2020)_250926.xlsx` → `세부기술별_데이터` 시트에서 적재
> **2025 데이터**: `2차델파이_5개국_요약테이블.xlsx` → `5개국_요약` 시트에서 적재

---

### 3.6 survey_response_raw (로데이터) ★ 2025~

개별 응답자의 원시 응답. 2020년은 없으므로 2025년부터만 존재.

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `response_id` | PK, INT | 응답 ID | 1 |
| `survey_year` | INT | 조사 연도 | 2025 |
| `delphi_round` | INT | 델파이 라운드 | 1, 2, 3 |
| `respondent_id` | INT | 응답자 ID (익명) | 2 |
| `detail_id` | FK → tech_detail | 세부기술 | 1 |
| `country` | VARCHAR(30) | 평가 대상 국가 | `중국` |
| `tech_group` | INT | 기술수준 그룹 (1~5) | 1 |
| `tech_level_pct` | DECIMAL(5,2) | 기술수준 비율 (%) | 100.0 |
| `tech_gap_month` | DECIMAL(5,1) | 기술격차 (개월) | 0 |
| `basic_research` | INT | 기초연구 역량 (1~5) | 5 |
| `applied_research` | INT | 응용개발연구 역량 (1~5) | 5 |
| `rd_trend` | INT | 연구개발 동향 (1~5) | 5 |

```
행 수:
  2차+3차 로데이터: 14,856건 (537명 × 세부기술 × 국가)
  1차 STEP1~4: 2,576건 (585명 × 세부기술 × 국가순위)
  1차 STEP5: 2,576건

인덱스: (survey_year, delphi_round, respondent_id, detail_id, country)
```

> **원본**: `2025 수준조사 Rawdata/2차_최종DATA_(3차 포함).xlsx`
> **참고**: 2차 로데이터는 26개국 응답 포함 (5개국 요약은 한/중/일/미/EU만 집계)

---

### ~~3.7 rnd_investment~~ (삭제)

> R&D 투자 분석은 현재 업무범위에서 제외됨.

---

### 3.7 activity_survey (기술개발활동조사) ★ 22대 분류

22대 기후기술 × 기관규모별 핵심 지표. 원본은 WT테이블(가중 적용 완료).

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| `activity_id` | PK, INT | 활동 ID | 1 |
| `survey_year` | INT | 조사 연도 | 2022, 2023, 2024 |
| `category_id` | FK → tech_category | 22대 기술 (T22) | 1 |
| `org_size` | VARCHAR(50) | 기관규모 층 | `층1(100억미만)`, `층5(2000억이상)` 등 |
| `metric_type` | VARCHAR(50) | 지표 유형 (아래 참조) | `revenue` |
| `metric_value` | DECIMAL(15,2) | 지표 값 | 30404081 |
| `metric_unit` | VARCHAR(20) | 단위 | `백만원`, `명`, `%`, `점` |
| `is_weighted` | BOOLEAN | 가중 여부 | TRUE |

```
분류: 22대(감축15 + 적응7) × 기관규모(6층) × 지표 × 연도
→ taxonomy_mapping으로 38대에 연계

보유 데이터: 2022, 2023, 2024 (3개년)
원본 파일:
  22년도: WT테이블 (가중, 140개 표)
  23년도: TAB 가중 + TAB 미가중 (각 129개 표)
  24년도: WT테이블 (가중, 140개 표)
```

**지표 유형 (metric_type)**:

| 섹션 | metric_type | 단위 | 설명 |
|------|-------------|------|------|
| A | `revenue` | 백만원 | 매출액 (전체/기후기술분야) |
| A | `revenue_no_reason` | % | 매출 미발생 사유 |
| B | `rnd_expense` | 백만원 | 연구개발비 (전체/기후기술분야) |
| B | `rnd_org_yn` | % | R&D 전담조직 보유 여부 |
| B | `rnd_outcome` | % | R&D 성과 유형 |
| C | `employee_count` | 명 | 종사자 수 (전체/연구원/생산/행정) |
| C | `recruit_ease` | % | 인력 충원 원활성 |
| D | `tech_transfer` | % | 기술이전/도입 경험 |
| E | `cooperation` | % | 협력활동 여부/유형/목적/지역 |
| F | `market` | % | 주요 시장/수출 계획/애로사항 |
| G | `policy_importance` | 점 | 중요도 평가 (5점척도) |
| G | `policy_satisfaction` | 점 | 만족도 평가 (5점척도) |

> ⚠️ **가중/미가중 이슈**: 3개년 공표보고서는 모두 가중 TAB 수치와 일치 확인.
> 다만 항목 성격별 가중/미가중 구분 적용 여부는 추가 확인 필요.
> 상세: `docs/activity-survey-data-guide.md` 참조

---

## 4. 집계 뷰 (Views)

### v_category_summary (중분류별 집계)

```sql
-- 38대 또는 44대 중분류별 5개국 평균
CREATE VIEW v_category_summary AS
SELECT
    sr.survey_year,
    tc.category_name,
    tc.type_name,
    sr.country_code,
    AVG(sr.tech_level)       AS avg_level,
    AVG(sr.tech_gap)         AS avg_gap,
    COUNT(sr.detail_id)      AS detail_count,
    SUM(sr.is_leading::int)  AS leading_count
FROM survey_result sr
JOIN tech_detail td ON sr.detail_id = td.detail_id
JOIN tech_category tc ON td.category_id = tc.category_id
GROUP BY sr.survey_year, tc.category_name, tc.type_name, sr.country_code;
```

### v_country_summary (국가별 전체 평균)

```sql
CREATE VIEW v_country_summary AS
SELECT
    survey_year,
    country_code,
    AVG(tech_level)  AS avg_level,
    AVG(tech_gap)    AS avg_gap,
    COUNT(*)         AS tech_count
FROM survey_result
GROUP BY survey_year, country_code;
```

### v_type_summary (감축/적응별 집계)

```sql
CREATE VIEW v_type_summary AS
SELECT
    sr.survey_year,
    tc.type_name,
    sr.country_code,
    AVG(sr.tech_level)  AS avg_level,
    AVG(sr.tech_gap)    AS avg_gap,
    COUNT(*)            AS tech_count
FROM survey_result sr
JOIN tech_detail td ON sr.detail_id = td.detail_id
JOIN tech_category tc ON td.category_id = tc.category_id
GROUP BY sr.survey_year, tc.type_name, sr.country_code;
```

### v_timeseries_compare (시계열 비교)

```sql
-- 44대↔38대 매핑 기반 시계열 비교
CREATE VIEW v_timeseries_compare AS
SELECT
    tm.mapping_type,
    tc_old.category_name AS category_2020,
    tc_new.category_name AS category_2025,
    sr_old.country_code,
    sr_old.avg_level AS level_2020,
    sr_new.avg_level AS level_2025,
    sr_new.avg_level - sr_old.avg_level AS level_delta
FROM taxonomy_mapping tm
JOIN tech_category tc_old ON tm.src_category_id = tc_old.category_id
JOIN tech_category tc_new ON tm.tgt_category_id = tc_new.category_id
JOIN v_category_summary sr_old
    ON sr_old.category_name = tc_old.category_name AND sr_old.survey_year = 2020
JOIN v_category_summary sr_new
    ON sr_new.category_name = tc_new.category_name AND sr_new.survey_year = 2025
    AND sr_old.country_code = sr_new.country_code
WHERE tm.mapping_status = 'confirmed';
```

### v_positioning_matrix (포지셔닝 매트릭스)

```sql
-- 38대 기술별 한국 기준 포지셔닝 데이터
CREATE VIEW v_positioning_matrix AS
SELECT
    tc.category_name,
    tc.type_name,
    kr.tech_level  AS kr_level,
    kr.tech_gap    AS kr_gap,
    us.tech_level  AS us_level,
    eu.tech_level  AS eu_level,
    kr.detail_count
FROM v_category_summary kr
JOIN v_category_summary us ON kr.category_name = us.category_name
    AND us.country_code = 'US' AND us.survey_year = kr.survey_year
JOIN v_category_summary eu ON kr.category_name = eu.category_name
    AND eu.country_code = 'EU' AND eu.survey_year = kr.survey_year
WHERE kr.country_code = 'KR';
```

### v_activity_summary (활동조사 집계)

```sql
-- 22대 기술별 핵심 지표 시계열
CREATE VIEW v_activity_summary AS
SELECT
    a.survey_year,
    tc.category_name,
    tc.type_name,
    a.metric_type,
    a.metric_unit,
    SUM(CASE WHEN a.org_size = '합계' THEN a.metric_value END) AS total_value,
    COUNT(DISTINCT a.org_size) AS size_count
FROM activity_survey a
JOIN tech_category tc ON a.category_id = tc.category_id
WHERE a.is_weighted = TRUE
GROUP BY a.survey_year, tc.category_name, tc.type_name, a.metric_type, a.metric_unit;
```

---

## 5. 데이터 흐름 (Data Flow)

```
[원본 Excel]                    [DB 적재]                    [대시보드]

=== 수준조사 ===

2025 Rawdata ──────┐
 1차 STEP1~4       │         survey_response_raw
 1차 STEP5         ├───────→ (14,856건)
 2차+3차 최종DATA   │            │ 집계
                   │            ▼
                   │         survey_result ──────→ v_category_summary
5개국 요약테이블 ───┤         (157×5=785건)        v_country_summary
                   │                              v_type_summary
                   │                              v_positioning_matrix
2020 세부기술 ──────┘         survey_result                   │
                             (185×5=925건)                    │
                                │                            │
                                ├─ taxonomy_mapping ─────────┤
                                │  (44↔38↔22↔100)            │
                                ▼                            ▼
                             v_timeseries_compare     [Streamlit]
                                                     app.py
=== 활동조사 ===                                      ├ level_survey.py
                                                      ├ positioning.py
22년도 WT테이블 ────┐                                  └ activity.py
23년도 TAB(가중) ───┼───→ activity_survey ────→ v_activity_summary
24년도 WT테이블 ────┘     (22대×규모×지표×3년)         │
                              │                      │
                              ├─ taxonomy_mapping ────┘
                              │  (22↔38)
                              ▼
23년도 TAB(미가중) ───→ (검증용 별도 보관)
```

---

## 6. 적재 전략

### Phase 1: 즉시 적재 (현재 데이터)

| 순서 | 원본 | 대상 테이블 | 건수 |
|------|------|-----------|------|
| 1 | 분류체계 정의 | taxonomy | 4건 (T44, T38, T22, T100) |
| 2 | 38대 기술 목록 | tech_category (T38) | 38건 |
| 3 | 44대 기술 목록 | tech_category (T44) | 44건 |
| 4 | 22대 기술 목록 | tech_category (T22) | 22건 |
| 5 | 157개 세부기술 (2025) | tech_detail | 157건 |
| 6 | 185개 세부기술 (2020) | tech_detail | 185건 |
| 7 | `기술연계표_맵핑표.xlsx` | taxonomy_mapping | ~200건 |
| 8 | `5개국_요약` (2025 집계) | survey_result | 785건 |
| 9 | `세부기술별_데이터` (2020) | survey_result | 925건 |
| 10 | `2차_최종DATA` (2025 로데이터) | survey_response_raw | 14,856건 |

### Phase 2: 활동조사 적재

| 순서 | 원본 | 대상 테이블 | 비고 |
|------|------|-----------|------|
| 11 | `기후기술조사(WT테이블) -22년도.xlsx` | activity_survey | 가중, 140개 표 파싱 |
| 12 | `2023년도 TAB_기후기술_V1_가중.xlsx` | activity_survey | 가중, 129개 표 파싱 |
| 13 | `260130_2024년도 활동조사_테이블.xlsx` | activity_survey | 가중, 140개 표 파싱 |
| 14 | 22대↔38대 매핑 (신규 작성) | taxonomy_mapping | |
| 15 | `2023년도 TAB_V1_미가중.xlsx` | (별도 보관) | 검증용, 가중 이슈 확인 후 재결정 |

---

## 7. 현재 대시보드와의 대응

| 현재 (data_loader.py) | DB 전환 후 |
|----------------------|-----------|
| `load_2025_data()` → Excel 직접 읽기 | `SELECT * FROM survey_result WHERE survey_year = 2025` |
| `load_2020_data()` → Excel 직접 읽기 | `SELECT * FROM survey_result WHERE survey_year = 2020` |
| `load_mapping_44_to_38()` → Excel | `SELECT * FROM taxonomy_mapping WHERE src=T44 AND tgt=T38` |
| `get_country_averages()` → Pandas 집계 | `SELECT * FROM v_country_summary` |
| `aggregate_by_category()` → Pandas 집계 | `SELECT * FROM v_category_summary` |
| 포지셔닝 매트릭스 → cat_stats 직접 계산 | `SELECT * FROM v_positioning_matrix` |

---

## 8. 향후 확장 고려

### 8.1 서비스 간 연계 (SFR-02 A-2)

```
survey_result (수준조사: 38대/44대)
    ↕ tech_category.category_id (38대 공통 키)
activity_survey (활동조사: 22대)
    ↕ taxonomy_mapping (22↔38)
    ↕ taxonomy_mapping (38↔100)
carbon_neutral_tech (탄소중립 100대, 향후)
```

### 8.2 데이터 품질 검증 룰 (SFR-02 E)

| 규칙 | SQL 조건 | 심각도 |
|------|---------|--------|
| 기술수준 범위 | `tech_level BETWEEN 0 AND 100` | ERROR |
| 기술격차 비음수 | `tech_gap >= 0` | ERROR |
| 5개국 완전성 | 세부기술당 country_code 5개 존재 | WARNING |
| 최고보유국 정합성 | `is_leading=TRUE`인 국가의 `tech_level`이 최대 | WARNING |
| 시계열 이상 감지 | 전년 대비 ±20%p 이상 변동 | INFO |
