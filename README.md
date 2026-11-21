# 임상미생물 주간 브리핑 대시보드 (v8.1)

Google Sheets(원본 DB) + GitHub Actions(자동 수집·2-Tier Gemini 정밀 채점·503 자동 우회) + GitHub Pages(초고속 4탭 대시보드) 아키텍처입니다.

## v8.1 주요 업그레이드
- **503 과부하·타임아웃 자동 우회(Fallback):** 논문·가이드라인 배치를 2건 단위로 안정화하고, `gemini-4-hydrogen` 혼잡 시 자동으로 동급 안정 모델(`gemini-3.6-flash`)로 우회하여 끊김 없이 완료합니다.
- **4단계 등급 + 100점 정량 채점:** AI가 4개 항목(실무 변경력·근거 강도·국내 적용성·혁신성, 각 25점)을 채점하면 코드가 합계 점수로 `필독(Must-Read)` / `주목(High)` / `참고(Medium)` / `단신(Low)`을 기계적으로 판정합니다.
- **2026년 1월 전체 확대:** 4대 저널(JCM·CMI·CID·ACM), 가이드라인, KDCA 전수·긴급·**표본감시 주간소식지(1주차부터 다중 페이지 소급)**, WHO DONs를 모두 2026-01-01부터 통합 수집합니다.

## 수동 실행 방법 (GitHub Actions)
- `collect`: 2026년 1월 이후 미수집 논문·가이드라인·표본감시·전수감시·WHO 일괄 수집 및 대시보드 갱신
- `resummarize_all`: 기존 데이터를 v8.1 4단계 + 100점 정량 채점 기준으로 일괄 재요약
- `reclassify_only`: 시트 `config`의 점수 컷오프(`cutoff_must_read` 등) 변경 시 **AI 비용 0원**으로 전체 등급 재계산
- `sync_sheet_only`: 구글 시트 수정분(`hai_stats` 월별 통계 등)만 대시보드에 즉시 반영 (**AI 비용 0원**)
- `test_apis`: API 키 및 Gemini 3개 모델(상위·우회·경량) 연결 진단
