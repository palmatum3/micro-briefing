# 임상미생물 주간 브리핑 대시보드 (v8.0)

Google Sheets(원본 DB) + GitHub Actions(자동 수집·Gemini 4 정밀 요약) + GitHub Pages(초고속 4탭 웹 대시보드) 아키텍처입니다.

## 주요 기능
- **6분 제한 완전 해방:** GitHub Actions에서 최대 3시간까지 끊김 없이 논문·가이드라인·KDCA·WHO 정보를 수집하고 `gemini-4-hydrogen` 상위 모델로 정밀 요약합니다.
- **초고속 로딩 (0.3초대):** 수집 결과를 `docs/data/dashboard.json`에 미리 빌드하여 모바일·PC 어디서나 즉시 열립니다.
- **지인 공유 안전장치:** 대시보드에 수동 수집 버튼이 없어 API 요금 오작동을 차단하며, 북마크는 각자 기기(`localStorage`)에 독립 저장됩니다.

## 수동 실행 방법 (GitHub 웹 화면)
1. 상단 **`Actions`** 탭 클릭 → 좌측 **`임상미생물 브리핑 자동화 (v8.0)`** 클릭
2. 우측 **`Run workflow`** 버튼 클릭 후 원하는 작업 선택:
   - `collect`: 최신 논문·가이드라인·감염병 수집 및 대시보드 업데이트
   - `resummarize_all`: 기존 논문·가이드라인 전체 `gemini-4-hydrogen` 정밀 재요약 + 감염병 경보 일괄 정비
   - `sync_sheet_only`: 구글 시트 수정분(`hai_stats` 월별 통계 등)만 대시보드에 즉시 반영 (AI 호출 0원)
   - `test_apis`: API 키 및 Gemini 모델 연결 진단
