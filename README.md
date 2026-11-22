# 임상미생물 주간 브리핑 대시보드 (v8.2.1.1)

Google Sheets(원본 DB) + GitHub Actions(Python v8.2.1 2-Pass 채점·요약 엔진) + **GitHub Pages(0.3초 고속 5탭 메인 대시보드)** 단일화 아키텍처입니다.

## v8.2.1 핵심 업그레이드
- **3단계 안전 중복 차단 (v8.2.1 신규):** UID·PMID·DOI 1차 차단 + 이종 저널 간 초록 지문 비교(단어 Jaccard ≥ 0.70 또는 3-gram 겹침 ≥ 0.55)로 동시 게재 지침·논문 자동 탐지 + AI 호출 0원 `duplicate` 비노출 마킹.
- **100점 고정 앵커 루브릭 (35 / 30 / 20 / 15):** 실무·검사실 변경력(35/26/16/6), 근거 강도(30/22/13/5), 국내 적용성(20/14/8/3), 신규성·파급력(15/11/6/2) 완전 고정 앵커 채점.
- **논문 유형별 상한선(Cap):** 단일기관 소규모·후향(82), 서술형 리뷰(78), 증례·Letter·Editorial(69), 순수 기초연구(65). 원점수와 Cap 이력을 투명 보존.
- **Critical Exception (E1~E4):** 필수 3조건 모두 충족 시 점수와 별도로 예외 필독 판정(원점수·Cap 보존) 및 2-Pass 심층 요약 강제 적용.
- **2-Pass 차등 요약 (안 B):** 1단계 전수 `gemini-4-hydrogen (thinking: low)` → 2단계 `Must-Read(85+)` 및 `Critical Exception` 전용 `gemini-4-hydrogen (thinking: high)` (503 폴백: `gemini-3.6-flash`).
- **상반기 소급 가드 (기본 활성):** `allow_h1_backfill=false` 기본 설정으로 2026-07-01 이후 하반기 논문에 집중하여 무료 크레딧을 보호.
- **GitHub Pages 5탭 단일화:** 주간 TOP 5에서 **Critical Exception 우선 노출 → 잔여 슬롯 점수순 노출**, 5번째 탭 `검사장비·시약` 플레이스홀더 추가.

## GitHub Actions 실행 모드
- `pilot_20`: 본 DB(`items`) 보호 — 2026-07 이후 층화 표본 20편을 별도 시트 탭(`pilot_v82`)에 검증 출력
- `collect`: 하반기 가드 적용 수집 및 대시보드 갱신 (`allow_h1_backfill` 체크 시 상반기 소급 해제)
- `resummarize_all`: 하반기(7월~) 기존 데이터를 v8.2.1 2-Pass 기준으로 일괄 정비
- `reclassify_only`: 시트 `config` 컷오프 변경 시 AI 비용 0원으로 등급 재계산
- `sync_sheet_only`: 시트 수정분(`hai_stats` 등)만 대시보드에 즉시 반영 (AI 비용 0원)
- `test_apis`: Google Sheets · Gemini 3개 모델 · NCBI · KDCA 연결 진단
