# 재사용 가능 자산 매니페스트

> 신규 국가비교 인포매틱스(및 향후 시각화)에서 **바로 재사용 가능한** 자산 목록.
> 원본은 삭제하지 않으며, 경로/재사용 포인트만 정리한다. 경로는 `ctis-stats/` 기준.

## 코드·상수

| 자산 | 경로 | 재사용 포인트 |
|---|---|---|
| 국가 컬러/국기 상수 | `src/dashboard/data_loader.py:21-29` (`COUNTRY_COLORS`, `COUNTRY_FLAGS`) | 격자·범례·차트 색상의 **단일 소스**. `extract_data.py`가 미러링 |
| 2025 데이터 로더 | `data_loader.py` `load_2025_data()` (line 338) | 157행 wide df (`leading_country`, `{cc}_level/_gap`, `category`, `type`) |
| long→wide 피벗 | `data_loader.py` `_pivot_survey_detail()` (line 50) | 세부×국가 → 분야별 행 변환 패턴 |
| 선도국 판별 | `_pivot_survey_detail():62-74` (`is_leading`→`leading_country`) | **분야별 선도국 집계 로직의 근거** |
| 카테고리/국가/유형 집계 | `_build_cat_stats`/`_build_country_stats`/`_build_type_stats` | 분야·국가·감축적응 요약 |
| 엑셀 다운로드 버튼 | `src/dashboard/pages/level_survey.py:24-35` (`_excel_download`) | 표 → xlsx 내보내기 |
| 한글경로/유니코드 처리 | `_find()`, `_normalize_name()` (data_loader) | 스크립트 경로·`·`/`?` 정규화 |

## 데이터

| 자산 | 경로 | 비고 |
|---|---|---|
| **SQLite DB** | `src/db/ctis_stats.db` | 추출 1차 소스 (`survey_result`·`tech_category`·`tech_detail`) |
| 검증된 5개국 요약 엑셀 | `src/data/2차델파이_5개국_요약테이블.xlsx` | DB 미가용 시 fallback 소스 (로더가 처리) |
| 2020 수준조사 | `src/data/수준조사 데이터(2020)_250926.xlsx` | 시계열 비교용 (44대) |
| 38대 표준 분류·매핑 | `src/mapping/참고_2025...매핑표_플랫폼탑재용_251106.xlsx` | 분야 표준명 |
| 2020↔2025 시계열 매핑 | `src/mapping/2020-2025 시계열 맵핑표(검토요청).xlsx` | 44↔38 (1:1, 1:N) |
| **추출 산출 (신규)** | `src/dashboard/infographics/data.json` · `data.js` | 분야 38·세부 157·5개국, 결측 0 (검증완료) |

## 디자인 토큰 (일관성 유지)

| 토큰 | 값 | 출처 |
|---|---|---|
| 한국 / 중국 / 일본 / 미국 / EU | `#FF6B6B` / `#4ECDC4` / `#45B7D1` / `#96CEB4` / `#FECA57` | `COUNTRY_COLORS` |
| CTis 브랜드 그린 (헤더/강조) | `#2d5016` → `#4a7c23` 그라디언트 | `src/dashboard/app.py` CSS |
| 감축 / 적응 | `#4a7c23` / `#2196F3` | 기존 페이지 관례 |

## 참고 (외부, 작년 코드)

```
C:\Users\정민경\jupyter\[파이썬코드] 2025 통합플랫폼 01. 수준조사 대시보드\
├── v4_ori.py      ← 최종 (트리맵·레이더·3메뉴)
├── dashboard.py   ← 고도화 (히트맵·게이지)
└── dashboard\dash_v3.py
```