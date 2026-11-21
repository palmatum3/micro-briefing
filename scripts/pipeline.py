#!/usr/bin/env python3
"""
임상미생물 주간 브리핑 자동 수집 · 2-Tier Gemini 정밀 요약 · 대시보드 빌드 엔진 (v8.0)
- Google Sheets(원본 DB) + GitHub Actions(무제한 수집·요약) + GitHub Pages(초고속 4탭 대시보드)
- 지원 모드 (--mode):
  * collect         : 전수감시 API + KDCA + WHO + 가이드라인 + 4대 저널 수집 -> 2-Tier Gemini 요약 -> 시트 저장 -> dashboard.json 빌드
  * resummarize_all : 기존 시트 논문·가이드라인 원문 초록 재조회 후 상위 모델(gemini-4-hydrogen) 정밀 재요약 + 경보 일괄 정비
  * sync_sheet_only : 구글 시트 현재 데이터만 읽어 dashboard.json 즉시 재빌드 (AI 호출 0원)
  * test_apis       : Google Sheets · Gemini(경량/상위) · NCBI · KDCA API 연결 진단
  * build_from_csv  : 로컬 CSV 파일로부터 초기 dashboard.json 생성 (초기 이관용)
"""

import argparse
import base64
import csv
import datetime
import html
import json
import os
import re
import sys
import time
import urllib.parse
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

CODE_VERSION = "v8.0"
KST = datetime.timezone(datetime.timedelta(hours=9))

SHEET_ITEMS = "items"
SHEET_LOGS = "run_logs"
SHEET_CONFIG = "config"
SHEET_KDCA_SURV = "kdca_surveillance"
SHEET_HAI_STATS = "hai_stats"

GEMINI_DEEP_BATCH_SIZE = 3
GEMINI_LITE_BATCH_SIZE = 5
KDCA_EID_BASE_URL = "https://apis.data.go.kr/1790387/EIDAPIService01"
DEFAULT_PAPER_START_DATE = "2026-07-01"
DEFAULT_YEAR_START_DATE = "2026-01-01"

PUBMED_CLEAN_FILTER = (
    ' AND hasabstract[text] NOT "Photo Quiz"[Title] NOT "Editorial"[Publication Type]'
    ' NOT "Letter"[Publication Type] NOT "Comment"[Publication Type]'
    ' NOT "Published Erratum"[Publication Type] NOT "News"[Publication Type]'
)

KDCA_URGENT_INCLUDE_RE = re.compile(
    r"(유행주의보|경보 발령|경보 해제|제1급.*발생|1급 감염병|콜레라.*발생|국내.*첫 발생|"
    r"CRE.*지침|카바페네마제.*지침|의료관련감염.*지침|항균제 내성.*지침)"
)
KDCA_EXCLUDE_RE = re.compile(
    r"(보도설명자료|성과 공유|퇴치 인증|민간·공공|예방접종은 정상|합동 훈련|도상훈련|협력 논의|"
    r"평가대회|만성질환|만성콩팥병|희귀질환|특수식|온열질환|한랭질환|흡연|담배|캠페인|MOU|"
    r"업무협약|인사|채용|공모전|심포지엄|포럼|기념식|시상식)"
)

WHO_COUNTRY_MAP = [
    (r"Switzerland|스위스", "스위스"),
    (r"Qatar|카타르", "카타르"),
    (r"Saudi Arabia|사우디아라비아", "사우디아라비아"),
    (r"Democratic Republic of the Congo|콩고민주공화국|민주콩고", "콩고민주공화국"),
    (r"Uganda|우간다", "우간다"),
    (r"Kenya|케냐", "케냐"),
    (r"Nigeria|나이지리아", "나이지리아"),
    (r"France|프랑스", "프랑스"),
    (r"India|인도", "인도"),
    (r"Bangladesh|방글라데시", "방글라데시"),
    (r"Italy|이탈리아", "이탈리아"),
    (r"Ethiopia|에티오피아", "에티오피아"),
    (r"Russian Federation|Russia|러시아", "러시아"),
    (r"United States|미국", "미국"),
    (r"United Kingdom|영국", "영국"),
    (r"Japan|일본", "일본"),
    (r"China|중국", "중국"),
    (r"Vietnam|베트남", "베트남"),
    (r"Thailand|태국", "태국"),
    (r"Indonesia|인도네시아", "인도네시아"),
    (r"Brazil|브라질", "브라질"),
    (r"African Region|아프리카", "아프리카"),
    (r"Multi-country|Multi-locations|다국가", "다국가"),
    (r"Global|전 세계|글로벌", "글로벌"),
]

ROOT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_JSON_PATH = ROOT_DIR / "docs" / "data" / "dashboard.json"

HTTP = requests.Session()
HTTP.headers.update({"User-Agent": "ClinicalMicroBriefing/8.0 (Academic Surveillance Bot)"})


def now_kst() -> datetime.datetime:
    return datetime.datetime.now(KST)


def now_kst_str() -> str:
    return now_kst().strftime("%Y-%m-%d %H:%M:%S")


def today_kst_str() -> str:
    return now_kst().strftime("%Y-%m-%d")


def strip_html(text: str) -> str:
    if not text:
        return ""
    s = re.sub(r"<[^>]*>", " ", str(text))
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def extract_spreadsheet_id(raw: str) -> str:
    s = (raw or "").strip()
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", s)
    return m.group(1) if m else s


def parse_flexible_date(val: Any) -> Optional[datetime.date]:
    if not val:
        return None
    s = str(val).strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S GMT", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.datetime.strptime(s[:31], fmt).date()
        except Exception:
            pass
    try:
        return datetime.datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except Exception:
        return None


def calc_days_since(date_str: str, fallback: int) -> int:
    d = parse_flexible_date(date_str)
    if not d:
        return fallback
    diff = (now_kst().date() - d).days
    return max(1, diff)


def get_iso_week_key(date_str: str) -> str:
    d = parse_flexible_date(date_str) or now_kst().date()
    iso_year, iso_week, _ = d.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def normalize_period_key(val: Any) -> str:
    s = str(val or "").strip()
    if not s or re.match(r"^(계|합계|total)$", s, re.IGNORECASE):
        return ""
    m = re.match(r"^(\d{4})-(\d{1,2})(?:-\d{1,2})?", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*주", s)
    if m:
        return f"{m.group(1)}-W{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})(\d{2})$", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return ""


def compute_prev_year_period_key(period_key: str) -> str:
    m = re.match(r"^(\d{4})(.+)$", str(period_key or ""))
    if not m:
        return ""
    return f"{int(m.group(1)) - 1}{m.group(2)}"


def extract_region_badge(source: str, title_orig: str, title_ko: str) -> str:
    src = str(source or "")
    if "KDCA" in src or "질병관리청" in src:
        if "표본감시" in src:
            return "국내 · 표본감시"
        return "국내 · KDCA"
    if "WHO" in src:
        combined = f"{title_orig} {title_ko}"
        found = []
        for pattern, label in WHO_COUNTRY_MAP:
            if re.search(pattern, combined, re.IGNORECASE):
                if label not in found:
                    found.append(label)
        if found:
            # Prefer specific country over broad region when both match
            specific = [x for x in found if x not in ("아프리카", "글로벌", "다국가")]
            picked = specific[:2] if specific else found[:1]
            return "해외 · " + "·".join(picked)
        return "해외 · WHO"
    return src


def clean_pmid(val: Any) -> str:
    s = str(val or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s.lower() == "nan" else s


def safe_str(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def safe_float(val: Any) -> Optional[float]:
    s = str(val or "").replace(",", "").strip()
    if not s or s.lower() == "nan":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def safe_int(val: Any, default: int = 0) -> int:
    f = safe_float(val)
    return int(round(f)) if f is not None else default


# =====================================================================
# Google Sheets Connector
# =====================================================================
def connect_spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    raw_sheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    if not raw_json:
        raise RuntimeError("환경변수 GOOGLE_SERVICE_ACCOUNT_JSON 이 설정되지 않았습니다.")
    if not raw_sheet_id:
        raise RuntimeError("환경변수 SPREADSHEET_ID 가 설정되지 않았습니다.")

    try:
        info = json.loads(raw_json)
    except json.JSONDecodeError:
        # Base64 fallback
        info = json.loads(base64.b64decode(raw_json).decode("utf-8"))

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id = extract_spreadsheet_id(raw_sheet_id)
    return gc.open_by_key(sheet_id)


def get_or_create_worksheet(ss, title: str, headers: List[str]):
    import gspread

    try:
        ws = ss.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows=200, cols=max(len(headers), 10))
        ws.append_row(headers, value_input_option="RAW")
    return ws


def ensure_item_meta_columns(ws):
    header = ws.row_values(1)
    if len(header) < 19:
        if ws.col_count < 19:
            ws.resize(cols=19)
        updates = []
        if len(header) < 18 or not header[17].strip():
            updates.append({"range": "R1", "values": [["summary_meta"]]})
        if len(header) < 19 or not header[18].strip():
            updates.append({"range": "S1", "values": [["importance_reason"]]})
        if updates:
            ws.batch_update(updates, value_input_option="RAW")


def load_config_map(ss) -> Dict[str, str]:
    ws = ss.worksheet(SHEET_CONFIG)
    rows = ws.get_all_values()
    cfg = {}
    for r in rows[1:]:
        if r and r[0].strip():
            cfg[r[0].strip()] = r[1].strip() if len(r) > 1 else ""
    return cfg


def is_enabled(cfg: Dict[str, str], key: str) -> bool:
    return cfg.get(key, "").upper() == "TRUE"


def get_gemini_models(cfg: Dict[str, str]) -> Tuple[str, str, str]:
    lite = cfg.get("gemini_model", "").strip() or "gemini-3.5-flash-lite"
    deep = cfg.get("gemini_model_deep", "").strip() or "gemini-4-hydrogen"
    thinking = cfg.get("gemini_thinking_level_deep", "").strip().lower() or "medium"
    return lite, deep, thinking


def append_run_log(ss, run_id: str, ts: str, source: str, status: str, fetched: int, added: int, dup: int, note: str):
    ws = get_or_create_worksheet(
        ss,
        SHEET_LOGS,
        ["run_id", "timestamp", "source", "status", "fetched_count", "new_count", "duplicate_count", "error_message"],
    )
    ws.append_row([run_id, ts, source, status, fetched, added, dup, note or ""], value_input_option="RAW")


# =====================================================================
# Gemini API (2-Tier: Deep Paper/Guideline vs Lite Outbreak)
# =====================================================================
def probe_gemini_model(model_name: str, api_key: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{urllib.parse.quote(model_name)}:generateContent"
    gen_cfg: Dict[str, Any] = {"temperature": 0}
    if re.match(r"^gemini-[34]", model_name, re.IGNORECASE):
        gen_cfg["thinkingConfig"] = {"thinkingLevel": "low"}
    payload = {"contents": [{"role": "user", "parts": [{"text": "Reply with OK only."}]}], "generationConfig": gen_cfg}
    res = HTTP.post(url, headers={"x-goog-api-key": api_key}, json=payload, timeout=45)
    if res.status_code == 200:
        return "성공 (HTTP 200)"
    return f"실패 (HTTP {res.status_code}: {res.text[:140]})"


def call_gemini_json_batch(
    prompt: str,
    response_schema: Dict[str, Any],
    model_name: str,
    api_key: str,
    temperature: float = 0.1,
    thinking_level: str = "",
) -> Dict[str, Dict[str, Any]]:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{urllib.parse.quote(model_name)}:generateContent"
    gen_cfg: Dict[str, Any] = {
        "temperature": temperature,
        "responseMimeType": "application/json",
        "responseSchema": response_schema,
    }
    if thinking_level in ("minimal", "low", "medium", "high") and re.match(r"^gemini-[34]", model_name, re.IGNORECASE):
        gen_cfg["thinkingConfig"] = {"thinkingLevel": thinking_level}

    payload = {"contents": [{"role": "user", "parts": [{"text": prompt}]}], "generationConfig": gen_cfg}

    last_err = ""
    for attempt in range(1, 5):
        try:
            res = HTTP.post(url, headers={"x-goog-api-key": api_key}, json=payload, timeout=150)
            if res.status_code == 200:
                body = res.json()
                parts = body.get("candidates", [{}])[0].get("content", {}).get("parts", [])
                text = ""
                for p in parts:
                    if isinstance(p, dict) and not p.get("thought") and p.get("text"):
                        text = p["text"]
                        break
                if not text and parts and isinstance(parts[0], dict):
                    text = parts[0].get("text", "")
                if not text:
                    raise RuntimeError("Gemini 응답 텍스트 없음")
                cleaned = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.IGNORECASE)
                cleaned = re.sub(r"\s*```$", "", cleaned)
                parsed = json.loads(cleaned)
                out: Dict[str, Dict[str, Any]] = {}
                for item in parsed.get("items", []):
                    if isinstance(item, dict) and item.get("uid"):
                        out[str(item["uid"])] = item
                return out
            last_err = f"HTTP {res.status_code}: {res.text[:220]}"
            if res.status_code in (429, 500, 502, 503, 504) and attempt < 4:
                wait_sec = 6 * attempt
                print(f"  [재시도 {attempt}/4] {model_name} {res.status_code}, {wait_sec}초 대기...")
                time.sleep(wait_sec)
                continue
            raise RuntimeError(f"Gemini API {last_err}")
        except (requests.RequestException, json.JSONDecodeError) as e:
            last_err = str(e)
            if attempt < 4:
                time.sleep(5 * attempt)
                continue
            raise RuntimeError(f"Gemini 호출 실패: {last_err}") from e
    raise RuntimeError(f"Gemini 호출 실패: {last_err}")


def format_clinical_note(raw_note: str, importance: str, high_axis: str, is_deep: bool) -> str:
    clean = str(raw_note or "").strip()
    clean = re.sub(r"^(\[?\s*AI\s*(임상\s*)?(참고\s*)?해석\s*\]?\s*[:：\-]?\s*)", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^AI의\s*(임상\s*)?(참고\s*)?해석에\s*따르면[,]?\s*", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^본\s*원저는\s*", "", clean, flags=re.IGNORECASE)
    if is_deep and importance == "High" and high_axis and high_axis != "해당없음":
        return f"[AI 해석 · 필독 근거: {high_axis}] {clean}"
    return f"[AI 해석] {clean}"


def summarize_deep_batch(
    chunk: List[Dict[str, Any]], model_name: str, api_key: str, thinking_level: str = "medium"
) -> Dict[str, Dict[str, Any]]:
    blocks = []
    for idx, it in enumerate(chunk, 1):
        raw = str(it.get("rawText") or "").strip()
        abs_display = raw[:3500] if len(raw) >= 40 else "(초록 원문 없음 — 제목만 제공됨)"
        blocks.append(
            "\n".join(
                [
                    f"### 항목 {idx}",
                    f"uid: {it['uid']}",
                    f"유형: {it['contentType']} / 출처: {it['source']}",
                    f"원제: {it['title']}",
                    f"초록 원문: {abs_display}",
                ]
            )
        )

    prompt = "\n".join(
        [
            "당신은 대한민국 대학병원 진단검사의학과(임상미생물학·분자진단·항균제 내성) 전문의를 지원하는 수석 학술 분석 AI입니다.",
            f'아래 {len(chunk)}편의 학술 논문 또는 가이드라인을 분석하여 JSON 객체({{ "items": [...] }})로만 응답하십시오. 각 항목의 uid를 정확히 유지해야 합니다.',
            "",
            "[절대 준수 원칙 — 위반 금지]",
            '1. 원문 왜곡 및 환각 금지: 초록 원문에 명시되지 않은 표본 수(n), 민감도/특이도, 균주, 결론을 절대 추측하거나 생성하지 마십시오. 초록 원문이 없거나 극히 짧으면 fact_summary_ko 첫 문장에 "초록 정보 부족으로 원문 확인 필요."를 명시하고 제목에서 확인되는 범위만 기술하십시오.',
            "2. fact_summary_ko (원문 사실 기반 정밀 요약, 3~4문장):",
            "   - 연구 목적·설계(다기관/단일기관, 전향/후향, 대상 검체·균주 수 n), 핵심 정량 결과(민감도·특이도·일치도·TAT·OR·p값 등 초록에 기재된 수치), 저자의 핵심 결론만 객관적으로 압축하십시오.",
            "   - 가이드라인인 경우 대상 질환/균종, 핵심 권고 변경점 및 근거 수준을 중심으로 정리하십시오.",
            "3. ai_clinical_note (진단검사의학과 전문의 실무 포인트, 2~3문장):",
            '   - 아래 뻔한 상용구 사용을 절대 금지합니다: "참고가 될 것으로 사료됩니다", "중요성을 시사합니다", "유용한 도구가 될 수 있습니다", "기초 자료로 활용될 수 있습니다".',
            "   - 대신 다음 두 가지를 구체적으로 작성하십시오:",
            "     (a) 대학병원 임상미생물 검사실(검사법 세팅·분자진단 도입·MALDI-TOF·AST 판독/보고·정도관리 QC·감염관리 IC) 관점의 실질적 적용 포인트 1가지.",
            "     (b) 현장 적용·해석 시 주의해야 할 연구 한계나 검증 필요 사항(예: 소규모 단일기관, 특정 균종 편중, 위양성/위음성 위험, 대조법 한계 등) 1가지.",
            "4. [4대 축 Strict High 선별 기준 — 전체 논문의 상위 10~15%만 High 부여]:",
            '   아래 4개 축 중 하나 이상에 명확히 부합할 때만 importance="High"를 부여하고 high_axis를 지정하십시오.',
            "   - 축① 진단·검사법 혁신: 신규 분자진단(신속 PCR·mNGS·MALDI-TOF·직접 AST 등)의 대규모·다기관 성능 평가(통상 n>=200 이상 또는 게임체인저급 TAT·정확도 혁신). ※ 단일기관 소규모(n<200) 예비평가나 기존 장비 마이너 업데이트는 Medium.",
            "   - 축② 항균제 내성·AST: CLSI/EUCAST breakpoint 변경 근거, 신종 내성 기전(카바페네마제 등) 출현·확산, AST 위양성·위음성 판독 주의보. ※ 특이 내성 기전 없는 단순 지역 역학은 Medium.",
            "   - 축③ 가이드라인·검사실 QC: IDSA·ESCMID 등 주요 학회 공식 가이드라인·권고안 제·개정, 검사실 QC·채취·배양 프로토콜의 실질적 변경을 요구하는 근거. ※ 단순 Narrative Review·Commentary는 Medium/Low.",
            "   - 축④ 원내감염·유행 관리: CRE, C. auris 등 주요 원내감염균의 대규모 유행 분석 및 실질적 감염관리 프로토콜 검증. ※ 증례보고(Case Report), 소규모 단일기관 역학, 동물/in vitro 기초실험은 Low.",
            "   - 초록이 없는 일반 논문은 High를 부여하지 마십시오(공식 학회 가이드라인 제외).",
            "5. topic_category (반드시 아래 목록 중 정확히 하나만 선택):",
            '   - guideline 유형: "IDSA 지침" | "ESCMID·국제 지침" | "국내 지침"',
            '   - paper 유형: "항균제 내성·AST" | "분자진단·신기술" | "임상 감염·치료" | "감염관리·역학" | "기타"',
            "",
            "\n\n".join(blocks),
        ]
    )

    schema = {
        "type": "OBJECT",
        "properties": {
            "items": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "uid": {"type": "STRING"},
                        "title_ko": {"type": "STRING"},
                        "topic_category": {
                            "type": "STRING",
                            "enum": [
                                "IDSA 지침",
                                "ESCMID·국제 지침",
                                "국내 지침",
                                "항균제 내성·AST",
                                "분자진단·신기술",
                                "임상 감염·치료",
                                "감염관리·역학",
                                "기타",
                            ],
                        },
                        "high_axis": {
                            "type": "STRING",
                            "enum": [
                                "진단·검사법 혁신",
                                "항균제 내성·AST",
                                "가이드라인·검사실 QC",
                                "원내감염·유행 관리",
                                "해당없음",
                            ],
                        },
                        "importance_reason": {"type": "STRING"},
                        "importance": {"type": "STRING", "enum": ["High", "Medium", "Low"]},
                        "fact_summary_ko": {"type": "STRING"},
                        "ai_clinical_note": {"type": "STRING"},
                    },
                    "required": [
                        "uid",
                        "title_ko",
                        "topic_category",
                        "high_axis",
                        "importance_reason",
                        "importance",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }
    return call_gemini_json_batch(prompt, schema, model_name, api_key, 0.1, thinking_level)


def summarize_outbreak_batch(
    chunk: List[Dict[str, Any]], model_name: str, api_key: str
) -> Dict[str, Dict[str, Any]]:
    blocks = []
    for idx, it in enumerate(chunk, 1):
        blocks.append(
            "\n".join(
                [
                    f"### 항목 {idx}",
                    f"uid: {it['uid']}",
                    f"출처: {it['source']} / 게시일: {it.get('pubDate', '')}",
                    f"제목: {it['title']}",
                    f"원문 텍스트: {str(it.get('rawText') or '(없음)')[:2500]}",
                ]
            )
        )

    prompt = "\n".join(
        [
            "당신은 대한민국 대학병원 진단검사의학과 전문의를 지원하는 감염병 감시 정보 분석 AI입니다.",
            f'아래 {len(chunk)}건의 국내외 감염병 동향·경보(KDCA/WHO)를 분석해 JSON 객체({{ "items": [...] }})로 반환하십시오.',
            "",
            "[준수 원칙]",
            '1. 원문에 없는 수치·국가·방역 조치를 절대 생성하지 마십시오. 학술 논문이 아니므로 "본 원저는" 같은 표현을 절대 사용하지 마십시오.',
            "2. fact_summary_ko:",
            "   - KDCA 표본감시 주간소식지: 인플루엔자(ILI)·코로나19 입원·주요 호흡기 바이러스 검출률·장관감염증 핵심 수치를 줄바꿈(•) 4줄 이내로 압축하십시오.",
            "   - KDCA 경보 / WHO DONs: 발생 국가·지역, 확진/사망 규모(치명률 포함), 실험실 확진 검사법(PCR·항체·mNGS 등), 방역 권고를 2~3문장으로 정리하십시오.",
            "3. ai_clinical_note (1~2문장):",
            '   - 뻔한 상용구("중요성을 시사합니다" 등)를 피하고, 국내 대학병원 진단검사실·감염관리실이 취해야 할 구체적 대응 포인트(검체 채취 부위, PCR/배양 감별 진단, 여행력 문진, 원내 선제 격리 등)를 작성하십시오.',
            "4. importance 판정 기준 (무분별한 High 방지):",
            "   - High: 국내 제1급 감염병·콜레라 발생, 전국 유행주의보 발령, CRE·다제내성균 국가 관리지침 발표, WHO 신종 병원체 최초 인체감염·고위험 병원체(MERS·에볼라·신종 바이러스) 신규 발생 또는 비유행지역 첫 자국 내 전파.",
            "   - Medium: 기존 해외 유행의 단순 후속 집계 업데이트(동일 유행 반복 보고), 정례 연보 발간 안내, 급증 특이사항 없는 주간 표본감시 정례 보고.",
            "   - Low: 국내 유입·검사실 영향이 극히 제한적인 국지적 단신.",
            "5. topic_category (반드시 아래 4개 중 하나 선택):",
            '   "표본감시 주간요약(호흡기·장관)" | "의료관련감염병·내성(HAI/CRE)" | "국내 긴급경보(KDCA)" | "국제 감염병 경보(WHO)"',
            "",
            "\n\n".join(blocks),
        ]
    )

    schema = {
        "type": "OBJECT",
        "properties": {
            "items": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "uid": {"type": "STRING"},
                        "title_ko": {"type": "STRING"},
                        "topic_category": {
                            "type": "STRING",
                            "enum": [
                                "표본감시 주간요약(호흡기·장관)",
                                "의료관련감염병·내성(HAI/CRE)",
                                "국내 긴급경보(KDCA)",
                                "국제 감염병 경보(WHO)",
                            ],
                        },
                        "importance_reason": {"type": "STRING"},
                        "importance": {"type": "STRING", "enum": ["High", "Medium", "Low"]},
                        "fact_summary_ko": {"type": "STRING"},
                        "ai_clinical_note": {"type": "STRING"},
                    },
                    "required": [
                        "uid",
                        "title_ko",
                        "topic_category",
                        "importance_reason",
                        "importance",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }
    return call_gemini_json_batch(prompt, schema, model_name, api_key, 0.1, "")


# =====================================================================
# Source Fetchers (KDCA API / Sentinel / Urgent RSS / WHO / PubMed / Crossref)
# =====================================================================
def call_kdca_eid(api_key: str, op_name: str, params: Dict[str, str]) -> List[Dict[str, Any]]:
    safe_key = api_key if "%" in api_key else urllib.parse.quote(api_key, safe="")
    q_parts = [f"serviceKey={safe_key}"]
    for k, v in params.items():
        if v is not None and v != "":
            q_parts.append(f"{urllib.parse.quote(k)}={urllib.parse.quote(str(v))}")
    url = f"{KDCA_EID_BASE_URL}/{op_name}?" + "&".join(q_parts)
    res = HTTP.get(url, timeout=45)
    body = res.text.strip()
    if res.status_code != 200:
        raise RuntimeError(f"{op_name} HTTP {res.status_code}: {body[:200]}")
    if "<returnAuthMsg>" in body or "<errMsg>" in body:
        m = re.search(r"<(?:returnAuthMsg|errMsg)>([^<]+)<", body)
        raise RuntimeError(f"{op_name} 오류: {m.group(1) if m else body[:200]}")
    if body.startswith("{") or body.startswith("["):
        return _extract_kdca_json(json.loads(body))
    if body.startswith("<"):
        return _extract_kdca_xml(body)
    raise RuntimeError(f"{op_name} 알 수 없는 응답 형식")


def _extract_kdca_json(node: Any) -> List[Dict[str, Any]]:
    if not node:
        return []
    if isinstance(node, list):
        return [x for x in node if isinstance(x, dict)]
    if not isinstance(node, dict):
        return []
    for k in ("item", "items", "data", "list", "result", "rows", "body", "response"):
        if node.get(k) is not None:
            found = _extract_kdca_json(node[k])
            if found:
                return found
    for v in node.values():
        if isinstance(v, list) and v and isinstance(v[0], dict):
            return v
    return []


def _extract_kdca_xml(xml_text: str) -> List[Dict[str, Any]]:
    root = ET.fromstring(xml_text)
    out = []
    for el in root.iter():
        if el.tag.lower() in ("item", "row"):
            obj = {c.tag: (c.text or "").strip() for c in el}
            if obj:
                out.append(obj)
    return out


def normalize_kdca_rows(raw: Dict[str, Any], fallback_year: str, op_name: str, now_str: str) -> List[Dict[str, Any]]:
    disease_name = safe_str(
        raw.get("icdNm") or raw.get("diseaseNm") or raw.get("infcNm") or raw.get("dsNm") or raw.get("sickNm") or raw.get("icd_nm")
    )
    if not disease_name:
        return []
    disease_group = safe_str(raw.get("icdGroupNm") or raw.get("groupNm") or raw.get("infcGrpNm") or raw.get("icd_group_nm"))
    yr = safe_str(raw.get("year") or raw.get("stdYear") or raw.get("yr") or fallback_year)
    out = []

    pivot_keys = [k for k in raw.keys() if re.match(r"^(?:m|mm|month|mon)_?0?(1[0-2]|[1-9])$", k, re.IGNORECASE)]
    if len(pivot_keys) >= 2 and yr:
        for k in pivot_keys:
            mm = int(re.sub(r"\D", "", k))
            val = safe_float(raw.get(k))
            if val is not None:
                out.append(
                    {
                        "updated_at": now_str,
                        "period_key": f"{yr}-{mm:02d}",
                        "disease_group": disease_group,
                        "disease_name": disease_name,
                        "case_count": val,
                        "source_op": op_name,
                    }
                )
        return out

    raw_cnt = None
    for ck in ("resultVal", "cnt", "patntCnt", "occrrncCnt", "totCnt", "value"):
        if raw.get(ck) is not None:
            raw_cnt = raw[ck]
            break
    case_count = safe_float(raw_cnt if raw_cnt is not None else 0)
    if case_count is None:
        return []

    wk = safe_str(raw.get("week") or raw.get("stdWeek") or raw.get("wk"))
    mo = safe_str(raw.get("month") or raw.get("stdMonth") or raw.get("mm"))
    raw_period = safe_str(raw.get("period") or raw.get("stdDe") or raw.get("baseDe") or raw.get("date") or raw.get("ym"))

    period_key = ""
    if yr and wk:
        wk_digits = re.sub(r'\D', '', wk)
        if wk_digits:
            period_key = f"{yr}-W{int(wk_digits):02d}"
    elif yr and mo:
        mo_digits = re.sub(r'\D', '', mo)
        if mo_digits:
            period_key = f"{yr}-{int(mo_digits):02d}"
    elif raw_period:
        period_key = normalize_period_key(raw_period)

    if not period_key:
        return []
    out.append(
        {
            "updated_at": now_str,
            "period_key": period_key,
            "disease_group": disease_group,
            "disease_name": disease_name,
            "case_count": case_count,
            "source_op": op_name,
        }
    )
    return out


def sync_kdca_surveillance_api(ss, kdca_api_key: str, now_str: str) -> Tuple[int, str]:
    ws = get_or_create_worksheet(
        ss, SHEET_KDCA_SURV, ["updated_at", "period_key", "disease_group", "disease_name", "case_count", "source_op"]
    )
    now_year = now_kst().year
    years = [str(now_year - 1), str(now_year)]
    all_records = []

    for yr in years:
        for page_no in range(1, 4):
            items = call_kdca_eid(
                kdca_api_key,
                "PeriodBasic01",
                {
                    "resType": "2",
                    "searchPeriodType": "2",
                    "searchStartYear": yr,
                    "searchEndYear": yr,
                    "pageNo": str(page_no),
                    "numOfRows": "1000",
                },
            )
            if not items:
                break
            for raw in items:
                all_records.extend(normalize_kdca_rows(raw, yr, "PeriodBasic01", now_str))
            if len(items) < 1000:
                break
            time.sleep(0.15)
        time.sleep(0.15)

    if not all_records:
        return 0, "PeriodBasic01 응답 파싱 0건"

    dedup: Dict[str, Dict[str, Any]] = {}
    for r in all_records:
        dedup[f"{r['period_key']}||{r['disease_name']}"] = r
    final_list = sorted(dedup.values(), key=lambda x: (x["period_key"], x["disease_name"]))

    header = [["updated_at", "period_key", "disease_group", "disease_name", "case_count", "source_op"]]
    body = [
        [r["updated_at"], r["period_key"], r["disease_group"], r["disease_name"], r["case_count"], r["source_op"]]
        for r in final_list
    ]
    ws.clear()
    ws.update(range_name="A1", values=header + body, value_input_option="RAW")
    cnt_cur = sum(1 for r in final_list if r["period_key"].startswith(f"{now_year}-"))
    return len(body), f"{now_year - 1}~{now_year}년 전수감시 {len(body)}건 갱신 ({now_year}년 {cnt_cur}건 포함)"


def fetch_kdca_sentinel_weekly(lookback_days: int) -> List[Dict[str, Any]]:
    url = "https://dportal.kdca.go.kr/pot/bbs/BD_selectBbsList.do?q_bbsSn=1010&q_clsfNo=2"
    res = HTTP.get(url, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"표본감시 주간소식지 HTTP {res.status_code}")
    cutoff = now_kst().date() - datetime.timedelta(days=lookback_days)
    out = []
    seen = set()
    for m in re.finditer(r"q_bbsDocNo=(\d{10,})[^>]*>([^<]{5,140})</a>", res.text):
        if len(out) >= 12:
            break
        doc_no = m.group(1)
        if doc_no in seen:
            continue
        seen.add(doc_no)
        title = strip_html(m.group(2))
        if not title or "표본감시" not in title:
            continue
        date_str = f"{doc_no[:4]}-{doc_no[4:6]}-{doc_no[6:8]}"
        parsed = parse_flexible_date(date_str)
        if parsed and parsed < cutoff:
            continue
        detail_url = f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1010&q_bbsDocNo={doc_no}&q_clsfNo=2"
        body_text = title
        try:
            time.sleep(0.2)
            d_res = HTTP.get(detail_url, timeout=25)
            if d_res.status_code == 200:
                clean = strip_html(d_res.text)
                idx = clean.find(title[:12])
                body_text = clean[idx : idx + 2500] if idx > -1 else clean[:2000]
        except Exception:
            pass
        out.append(
            {
                "uid": f"KDCA_SENTINEL_{doc_no}",
                "source": "KDCA 표본감시 주간소식지",
                "contentType": "outbreak",
                "pubDate": date_str,
                "title": title,
                "rawText": body_text,
                "pmid": "",
                "doi": "",
                "sourceUrl": detail_url,
                "accessStatus": "표본감시 주간소식지",
            }
        )
    return out


def fetch_kdca_urgent_alerts(lookback_days: int) -> List[Dict[str, Any]]:
    url = "https://www.kdca.go.kr/kdcaBbs/kdca/41/rssList.do?row=100"
    res = HTTP.get(url, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"KDCA RSS HTTP {res.status_code}")
    cutoff = now_kst().date() - datetime.timedelta(days=lookback_days)
    root = ET.fromstring(res.content)
    channel = root.find("channel")
    if channel is None:
        return []
    out = []
    for item in channel.findall("item"):
        title = re.sub(r"\}+$", "", (item.findtext("title") or "")).strip()
        raw_link = (item.findtext("link") or "").strip()
        link = ("https://www.kdca.go.kr" + raw_link) if raw_link.startswith("/") else raw_link
        pub_raw = (item.findtext("pubDate") or "").strip()
        desc = strip_html(item.findtext("description") or "")
        if not title or not link:
            continue
        if not KDCA_URGENT_INCLUDE_RE.search(title) or KDCA_EXCLUDE_RE.search(title):
            continue
        parsed = parse_flexible_date(pub_raw)
        if parsed and parsed < cutoff:
            continue
        date_str = parsed.strftime("%Y-%m-%d") if parsed else today_kst_str()
        id_m = re.search(r"/(\d+)/artclView", link) or re.search(r"(\d{5,})", link)
        uid_suffix = id_m.group(1) if id_m else base64.urlsafe_b64encode(title.encode()).decode()[:16]
        out.append(
            {
                "uid": f"KDCA_{uid_suffix}",
                "source": "KDCA 긴급경보",
                "contentType": "outbreak",
                "pubDate": date_str,
                "title": title,
                "rawText": desc or title,
                "pmid": "",
                "doi": "",
                "sourceUrl": link,
                "accessStatus": "초긴급 경보",
            }
        )
    return out


def fetch_who_dons(lookback_days: int) -> List[Dict[str, Any]]:
    url = "https://www.who.int/api/news/diseaseoutbreaknews?$top=40&$orderby=PublicationDate%20desc"
    res = HTTP.get(url, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"WHO DONs API HTTP {res.status_code}")
    cutoff = now_kst().date() - datetime.timedelta(days=lookback_days)
    data = res.json()
    rows = data if isinstance(data, list) else data.get("value", [])
    out = []
    for r in rows:
        pub_raw = r.get("PublicationDate") or r.get("DateCreated") or ""
        parsed = parse_flexible_date(pub_raw)
        if parsed and parsed < cutoff:
            continue
        date_str = parsed.strftime("%Y-%m-%d") if parsed else today_kst_str()
        don_id = safe_str(r.get("DonId") or r.get("Id"))
        url_name = safe_str(r.get("UrlName"))
        title = safe_str(r.get("Title"))
        if not title:
            continue
        summary = strip_html(r.get("Overview") or r.get("Summary") or r.get("Assessment") or "")
        source_url = (
            f"https://www.who.int/emergencies/disease-outbreak-news/item/{url_name}"
            if url_name
            else "https://www.who.int/emergencies/disease-outbreak-news"
        )
        out.append(
            {
                "uid": f"WHO_{don_id or url_name or date_str}",
                "source": "WHO DONs",
                "contentType": "outbreak",
                "pubDate": date_str,
                "title": title,
                "rawText": summary[:3000] or title,
                "pmid": "",
                "doi": "",
                "sourceUrl": source_url,
                "accessStatus": "WHO 공식 경보",
            }
        )
    return out


def fetch_acm_crossref(lookback_days: int) -> List[Dict[str, Any]]:
    cutoff = now_kst().date() - datetime.timedelta(days=lookback_days)
    from_date = cutoff.strftime("%Y-%m-%d")
    url = f"https://api.crossref.org/journals/2288-6850/works?filter=from-pub-date:{from_date}&rows=50&sort=published&order=desc"
    res = HTTP.get(url, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"Crossref ACM HTTP {res.status_code}")
    items = res.json().get("message", {}).get("items", [])
    out = []
    for it in items:
        titles = it.get("title") or []
        title = safe_str(titles[0]) if titles else ""
        doi = safe_str(it.get("DOI"))
        if not title or not doi:
            continue
        if re.match(r"^(editorial|erratum|corrigendum|retraction)", title, re.IGNORECASE):
            continue
        parts = (
            (it.get("published") or {}).get("date-parts", [[]])[0]
            or (it.get("created") or {}).get("date-parts", [[]])[0]
        )
        if len(parts) >= 3:
            pub_date = f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
        elif len(parts) == 2:
            pub_date = f"{parts[0]}-{int(parts[1]):02d}-01"
        else:
            pub_date = from_date
        abs_text = strip_html(it.get("abstract") or "")
        uid_safe = re.sub(r"[^a-zA-Z0-9.-]", "_", doi)
        out.append(
            {
                "uid": f"ACM_{uid_safe}",
                "source": "ACM",
                "contentType": "guideline" if re.search(r"guideline", title, re.IGNORECASE) else "paper",
                "pubDate": pub_date,
                "title": title,
                "rawText": abs_text,
                "pmid": "",
                "doi": doi,
                "sourceUrl": f"https://doi.org/{doi}",
                "accessStatus": "Open Access (ACM)" if abs_text else "Crossref 메타데이터",
            }
        )
    return out


def fetch_crossref_abstract(doi: str) -> str:
    if not doi:
        return ""
    try:
        res = HTTP.get(f"https://api.crossref.org/works/{urllib.parse.quote(doi)}", timeout=20)
        if res.status_code == 200:
            return strip_html(res.json().get("message", {}).get("abstract") or "")
    except Exception:
        pass
    return ""


def fetch_pubmed_ids(query: str, lookback_days: int, fetch_limit: int, ncbi_api_key: str) -> List[str]:
    params = {
        "db": "pubmed",
        "retmode": "json",
        "sort": "pub date",
        "datetype": "edat",
        "tool": "ClinicalMicroBriefing",
        "reldate": str(lookback_days),
        "retmax": str(fetch_limit),
        "term": query,
    }
    if ncbi_api_key:
        params["api_key"] = ncbi_api_key
    res = HTTP.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi", params=params, timeout=30)
    if res.status_code != 200:
        raise RuntimeError(f"PubMed ESearch HTTP {res.status_code}")
    return res.json().get("esearchresult", {}).get("idlist", [])


def _format_pubmed_date(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    year = (el.findtext("Year") or "").strip()
    raw_m = (el.findtext("Month") or "01").strip()
    day = (el.findtext("Day") or "01").strip().zfill(2)
    mmap = {"Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04", "May": "05", "Jun": "06",
            "Jul": "07", "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"}
    month = mmap.get(raw_m, raw_m.zfill(2))
    return f"{year}-{month}-{day}" if year else ""


def fetch_pubmed_articles(pmids: List[str], ncbi_api_key: str) -> List[Dict[str, Any]]:
    results = []
    # Fetch in chunks of 100 PMIDs safely
    for start in range(0, len(pmids), 100):
        sub = pmids[start : start + 100]
        params = {"db": "pubmed", "retmode": "xml", "tool": "ClinicalMicroBriefing", "id": ",".join(sub)}
        if ncbi_api_key:
            params["api_key"] = ncbi_api_key
        res = HTTP.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi", params=params, timeout=45)
        if res.status_code != 200:
            raise RuntimeError(f"PubMed EFetch HTTP {res.status_code}")
        root = ET.fromstring(res.content)
        for art_node in root.findall("PubmedArticle"):
            medline = art_node.find("MedlineCitation")
            pubmed_data = art_node.find("PubmedData")
            if medline is None:
                continue
            pmid = (medline.findtext("PMID") or "").strip()
            article = medline.find("Article")
            if article is None:
                continue
            title_el = article.find("ArticleTitle")
            title = re.sub(r"\s+", " ", "".join(title_el.itertext())).strip() if title_el is not None else "(제목 없음)"

            abs_parts = []
            abs_el = article.find("Abstract")
            if abs_el is not None:
                for t in abs_el.findall("AbstractText"):
                    lbl = t.attrib.get("Label")
                    txt = re.sub(r"\s+", " ", "".join(t.itertext())).strip()
                    abs_parts.append(f"{lbl}: {txt}" if lbl else txt)
            abstract_text = " ".join(abs_parts)

            journal = article.find("Journal")
            j_title = ""
            if journal is not None:
                j_title = (journal.findtext("ISOAbbreviation") or journal.findtext("Title") or "").strip()

            art_dates = article.findall("ArticleDate")
            if art_dates:
                pub_date = _format_pubmed_date(art_dates[0])
            else:
                issue = journal.find("JournalIssue") if journal is not None else None
                pub_date = _format_pubmed_date(issue.find("PubDate") if issue is not None else None)

            doi = ""
            has_pmc = False
            if pubmed_data is not None:
                for aid in pubmed_data.findall("./ArticleIdList/ArticleId"):
                    t = aid.attrib.get("IdType")
                    if t == "doi":
                        doi = (aid.text or "").strip()
                    elif t == "pmc":
                        has_pmc = True

            results.append(
                {
                    "pmid": pmid,
                    "journal": j_title,
                    "title": title,
                    "abstractText": abstract_text,
                    "pubDate": pub_date,
                    "doi": doi,
                    "accessStatus": "Open Access (PMC)" if has_pmc else "Abstract-only",
                }
            )
        if start + 100 < len(pmids):
            time.sleep(0.15 if ncbi_api_key else 0.4)
    return results


def detect_guideline_org(title: str, journal: str) -> str:
    if re.search(r"IDSA|Infectious Diseases Society of America", title, re.IGNORECASE):
        return "IDSA"
    if re.search(r"ESCMID|European Society of Clinical Microbiology", title, re.IGNORECASE):
        return "ESCMID"
    return f"Guideline · {journal}" if journal else "Guideline"


# =====================================================================
# Dashboard JSON Builder (Works from Google Sheets OR Local CSV)
# =====================================================================
def build_surveillance_payload(surv_rows: List[Dict[str, Any]], hai_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    current_year = str(now_kst().year)
    prev_year = str(now_kst().year - 1)

    by_disease: Dict[str, Dict[str, Any]] = {}
    latest_updated_at = ""
    for r in surv_rows:
        updated = safe_str(r.get("updated_at"))
        p_key = normalize_period_key(r.get("period_key"))
        grp = safe_str(r.get("disease_group"))
        name = safe_str(r.get("disease_name"))
        cnt = safe_float(r.get("case_count"))
        if not p_key or not name or cnt is None:
            continue
        if updated > latest_updated_at:
            latest_updated_at = updated
        if name not in by_disease:
            by_disease[name] = {"name": name, "group": grp, "series": []}
        by_disease[name]["series"].append({"period": p_key, "count": cnt})

    mdr_cards = []
    surge_alerts = []
    cre_series = {"currentYear": current_year, "current": [], "prev": []}

    for name, entry in by_disease.items():
        s = sorted(entry["series"], key=lambda x: x["period"])
        if not s:
            continue
        latest = s[-1]
        prev4 = s[max(0, len(s) - 5) : len(s) - 1]
        avg4 = (sum(x["count"] for x in prev4) / len(prev4)) if prev4 else None
        vs_avg4_pct = round(((latest["count"] - avg4) / avg4) * 100, 1) if (avg4 and avg4 > 0) else None

        yoy_key = compute_prev_year_period_key(latest["period"])
        yoy_pt = next((x for x in s if x["period"] == yoy_key), None) if yoy_key else None
        vs_yoy_pct = (
            round(((latest["count"] - yoy_pt["count"]) / yoy_pt["count"]) * 100, 1)
            if (yoy_pt and yoy_pt["count"] > 0)
            else None
        )

        is_cre = bool(re.search(r"CRE|카바페넴내성", name, re.IGNORECASE))
        is_vrsa = bool(re.search(r"VRSA|반코마이신내성", name, re.IGNORECASE))

        if is_cre or is_vrsa:
            mdr_cards.append(
                {
                    "code": "CRE" if is_cre else "VRSA",
                    "label": "CRE (카바페넴내성장내세균목)" if is_cre else "VRSA (반코마이신내성황색포도알균)",
                    "period": latest["period"],
                    "count": int(latest["count"]) if float(latest["count"]).is_integer() else latest["count"],
                    "avg4": round(avg4, 1) if avg4 is not None else None,
                    "vsAvg4Pct": vs_avg4_pct,
                    "vsYoyPct": vs_yoy_pct,
                }
            )
            if is_cre:
                cre_series["current"] = [x for x in s if x["period"].startswith(f"{current_year}-")]
                cre_series["prev"] = [x for x in s if x["period"].startswith(f"{prev_year}-")]
        else:
            is_class1 = bool(
                re.search(r"제1급|1급|콜레라|에볼라|마버그|라싸|페스트|탄저|두창|MERS|중동호흡기", f"{entry['group']} {name}", re.IGNORECASE)
            )
            if is_class1 and latest["count"] > 0:
                surge_alerts.append(
                    {
                        "region": "국내 · KDCA",
                        "level": "critical",
                        "disease": name,
                        "period": latest["period"],
                        "count": int(latest["count"]),
                        "vsAvg4Pct": vs_avg4_pct,
                        "reason": f"국내 제1급·콜레라 발생 감지 ({latest['period']} · {int(latest['count'])}건)",
                    }
                )
            elif vs_avg4_pct is not None and vs_avg4_pct >= 30 and latest["count"] >= 10:
                surge_alerts.append(
                    {
                        "region": "국내 · KDCA",
                        "level": "surge",
                        "disease": name,
                        "period": latest["period"],
                        "count": int(latest["count"]),
                        "vsAvg4Pct": vs_avg4_pct,
                        "reason": f"국내 최근 4기 평균({round(avg4)}건) 대비 +{vs_avg4_pct}% 급증",
                    }
                )

    mdr_cards.sort(key=lambda x: 0 if x["code"] == "CRE" else 1)
    surge_alerts.sort(key=lambda x: (0 if x["level"] == "critical" else 1, -(x["vsAvg4Pct"] or 0)))

    hai_by_patho: Dict[str, List[Dict[str, Any]]] = {}
    for r in hai_rows:
        ym = normalize_period_key(r.get("period_ym"))
        patho = safe_str(r.get("pathogen")).upper()
        blood = safe_float(r.get("blood_rate"))
        non_blood = safe_float(r.get("non_blood_rate"))
        if not ym or not patho or (blood is None and non_blood is None):
            continue
        hai_by_patho.setdefault(patho, []).append({"period": ym, "blood": blood, "nonBlood": non_blood})

    hai_order = ["MRSA", "VRE", "MRAB", "MRPA", "C.AURIS"]
    hai_cards = []
    hai_annual_series = []
    for p in hai_order:
        arr = sorted(hai_by_patho.get(p, []), key=lambda x: x["period"])
        if not arr:
            continue
        latest = arr[-1]
        prev = arr[-2] if len(arr) >= 2 else None
        blood_delta = (
            round(((latest["blood"] - prev["blood"]) / prev["blood"]) * 100, 1)
            if (prev and prev["blood"] and latest["blood"] is not None)
            else None
        )
        non_blood_delta = (
            round(((latest["nonBlood"] - prev["nonBlood"]) / prev["nonBlood"]) * 100, 1)
            if (prev and prev["nonBlood"] and latest["nonBlood"] is not None)
            else None
        )
        hai_cards.append(
            {
                "pathogen": p,
                "period": latest["period"],
                "bloodRate": latest["blood"],
                "nonBloodRate": latest["nonBlood"],
                "bloodDeltaPct": blood_delta,
                "nonBloodDeltaPct": non_blood_delta,
            }
        )
        cur_yr = [x for x in arr if x["period"].startswith(f"{current_year}-")]
        hai_annual_series.append({"pathogen": p, "points": cur_yr if cur_yr else arr})

    return {
        "year": current_year,
        "updatedAt": latest_updated_at,
        "mdrCards": mdr_cards,
        "haiCards": hai_cards,
        "haiAnnualSeries": hai_annual_series,
        "creSeries": cre_series,
        "surgeAlerts": surge_alerts[:6],
    }


def write_dashboard_json(
    item_rows: List[Dict[str, Any]],
    surv_rows: List[Dict[str, Any]],
    hai_rows: List[Dict[str, Any]],
    log_rows: List[Dict[str, Any]],
    cfg_map: Dict[str, str],
) -> None:
    items = []
    for r in item_rows:
        uid = safe_str(r.get("uid"))
        if not uid:
            continue
        pub_date = safe_str(r.get("pub_date"))[:10]
        source = safe_str(r.get("source"))
        title_orig = safe_str(r.get("title_orig"))
        title_ko = safe_str(r.get("title_ko")) or title_orig
        ctype = safe_str(r.get("content_type")) or "paper"
        week_key = safe_str(r.get("week_key")) or get_iso_week_key(pub_date)
        items.append(
            {
                "uid": uid,
                "source": source,
                "pub_date": pub_date,
                "collected_at": safe_str(r.get("collected_at")),
                "title_orig": title_orig,
                "title_ko": title_ko,
                "topic_category": safe_str(r.get("topic_category")) or "기타",
                "importance": safe_str(r.get("importance")) or "Medium",
                "fact_summary_ko": safe_str(r.get("fact_summary_ko")),
                "ai_clinical_note": safe_str(r.get("ai_clinical_note")),
                "pmid": clean_pmid(r.get("pmid")),
                "doi": safe_str(r.get("doi")),
                "source_url": safe_str(r.get("source_url")),
                "access_status": safe_str(r.get("access_status")) or "Public",
                "bookmarked": str(r.get("bookmarked", "")).strip().upper() == "TRUE",
                "content_type": ctype,
                "week_key": week_key,
                "summary_meta": safe_str(r.get("summary_meta")),
                "importance_reason": safe_str(r.get("importance_reason")),
                "region_badge": extract_region_badge(source, title_orig, title_ko) if ctype == "outbreak" else "",
            }
        )

    items.sort(key=lambda x: (x["pub_date"], x["uid"]), reverse=True)

    logs = []
    for r in reversed(log_rows[-30:]):
        if not safe_str(r.get("run_id")):
            continue
        logs.append(
            {
                "run_id": safe_str(r.get("run_id")),
                "timestamp": safe_str(r.get("timestamp")),
                "source": safe_str(r.get("source")),
                "status": safe_str(r.get("status")),
                "fetched_count": safe_int(r.get("fetched_count")),
                "new_count": safe_int(r.get("new_count")),
                "duplicate_count": safe_int(r.get("duplicate_count")),
                "error_message": safe_str(r.get("error_message")),
            }
        )

    lite_model, deep_model, thinking = get_gemini_models(cfg_map)
    payload = {
        "meta": {
            "version": CODE_VERSION,
            "generated_at": now_kst_str(),
            "total_items": len(items),
            "models": {"deep": deep_model, "lite": lite_model, "thinking": thinking},
        },
        "items": items,
        "logs": logs,
        "surveillance": build_surveillance_payload(surv_rows, hai_rows),
    }

    OUTPUT_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"[OK] dashboard.json 생성 완료: {OUTPUT_JSON_PATH} ({len(items)}건, {OUTPUT_JSON_PATH.stat().st_size:,} bytes)")


def read_sheet_as_dicts(ss, sheet_name: str) -> List[Dict[str, Any]]:
    ws = ss.worksheet(sheet_name)
    values = ws.get_all_values()
    if len(values) < 2:
        return []
    headers = [h.strip() for h in values[0]]
    out = []
    for row in values[1:]:
        if not any(c.strip() for c in row):
            continue
        padded = row + [""] * max(0, len(headers) - len(row))
        out.append(dict(zip(headers, padded)))
    return out


def sync_sheet_to_dashboard_json(ss) -> None:
    cfg_map = load_config_map(ss)
    items = read_sheet_as_dicts(ss, SHEET_ITEMS)
    surv = read_sheet_as_dicts(ss, SHEET_KDCA_SURV)
    hai = read_sheet_as_dicts(ss, SHEET_HAI_STATS)
    logs = read_sheet_as_dicts(ss, SHEET_LOGS)
    write_dashboard_json(items, surv, hai, logs, cfg_map)


# =====================================================================
# Main Modes
# =====================================================================
def mode_test_apis() -> None:
    print(f"=== 임상미생물 브리핑 {CODE_VERSION} API 연결 진단 ===")
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    lite_model, deep_model, thinking = get_gemini_models(cfg)
    print(f"[OK] Google Sheets 연결 성공: {ss.title}")

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()

    if gemini_key:
        print(f"1. Gemini 경량({lite_model}): {probe_gemini_model(lite_model, gemini_key)}")
        print(f"   Gemini 상위({deep_model}, thinking={thinking}): {probe_gemini_model(deep_model, gemini_key)}")
    else:
        print("1. GEMINI_API_KEY: 미등록")

    if ncbi_key:
        ids = fetch_pubmed_ids('"J Clin Microbiol"[Journal]', 7, 3, ncbi_key)
        print(f"2. NCBI PubMed (고속 10 req/s): 성공 (최근 7일 JCM {len(ids)}건)")
    else:
        print("2. NCBI_API_KEY: 미등록 (기본 3 req/s 모드로 동작)")

    if kdca_key:
        yr = str(now_kst().year)
        sample = call_kdca_eid(
            kdca_key,
            "PeriodBasic01",
            {"resType": "2", "searchPeriodType": "2", "searchStartYear": yr, "searchEndYear": yr, "pageNo": "1", "numOfRows": "10"},
        )
        print(f"3. KDCA 전수신고 API ({yr}년): 성공 ({len(sample)}건 샘플 수신)")
    else:
        print("3. KDCA_API_KEY: 미등록")


def mode_collect() -> None:
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    item_ws = ss.worksheet(SHEET_ITEMS)
    ensure_item_meta_columns(item_ws)

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()
    if not gemini_key:
        raise RuntimeError("GEMINI_API_KEY 가 설정되지 않았습니다.")

    run_id = "GH_RUN_" + now_kst().strftime("%Y%m%d_%H%M%S")
    now_str = now_kst_str()
    today_str = today_kst_str()

    paper_days = calc_days_since(cfg.get("collect_start_date", DEFAULT_PAPER_START_DATE), 140)
    year_days = calc_days_since(cfg.get("kdca_start_date", DEFAULT_YEAR_START_DATE), 365)
    fetch_limit = safe_int(cfg.get("pubmed_fetch_limit"), 300)
    lite_model, deep_model, deep_thinking = get_gemini_models(cfg)
    pubmed_delay = 0.15 if ncbi_key else 0.4

    existing_uids = {v.strip() for v in item_ws.col_values(1)[1:] if v.strip()}
    candidates: List[Dict[str, Any]] = []
    fetch_errors: List[str] = []
    status_notes: List[str] = []
    total_fetched = 0
    total_dup = 0

    # 0) KDCA 전수신고 API 동기화
    if is_enabled(cfg, "source_kdca_enabled") and kdca_key:
        try:
            _, note = sync_kdca_surveillance_api(ss, kdca_key, now_str)
            status_notes.append(note)
            print(f"[KDCA API] {note}")
        except Exception as e:
            fetch_errors.append(f"전수신고API: {e}")

    # 1) KDCA 표본감시 · 긴급경보 · WHO
    if is_enabled(cfg, "source_kdca_enabled"):
        try:
            for it in fetch_kdca_sentinel_weekly(year_days):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"표본감시소식지: {e}")

        try:
            for it in fetch_kdca_urgent_alerts(year_days):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"KDCA긴급경보: {e}")

    if is_enabled(cfg, "source_who_enabled"):
        try:
            for it in fetch_who_dons(year_days):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"WHO: {e}")

    # 2) 가이드라인 (IDSA · ESCMID)
    if is_enabled(cfg, "source_guideline_enabled"):
        try:
            g_query = (
                '(("Infectious Diseases Society of America"[Title] OR "IDSA"[Title] OR "ESCMID"[Title]'
                ' OR "European Society of Clinical Microbiology"[Title]) AND ("Guideline"[Title] OR "Guidance"[Title] OR "Practice Guideline"[Publication Type]))'
                + PUBMED_CLEAN_FILTER
            )
            g_ids = fetch_pubmed_ids(g_query, year_days, fetch_limit, ncbi_key)
            total_fetched += len(g_ids)
            pending_g = [i for i in g_ids if f"GUIDE_{i}" not in existing_uids]
            total_dup += len(g_ids) - len(pending_g)
            if pending_g:
                time.sleep(pubmed_delay)
                for a in fetch_pubmed_articles(pending_g, ncbi_key):
                    candidates.append(
                        {
                            "uid": f"GUIDE_{a['pmid']}",
                            "source": detect_guideline_org(a["title"], a["journal"]),
                            "contentType": "guideline",
                            "pubDate": a["pubDate"],
                            "title": a["title"],
                            "rawText": a["abstractText"],
                            "pmid": a["pmid"],
                            "doi": a["doi"],
                            "sourceUrl": f"https://pubmed.ncbi.nlm.nih.gov/{a['pmid']}/",
                            "accessStatus": a["accessStatus"],
                        }
                    )
        except Exception as e:
            fetch_errors.append(f"GUIDELINE: {e}")

    # 3) 4대 저널 (JCM · CMI · CID · ACM)
    journals = [
        ("source_jcm_enabled", "JCM", '"J Clin Microbiol"[Journal]' + PUBMED_CLEAN_FILTER),
        ("source_cmi_enabled", "CMI", '"Clin Microbiol Infect"[Journal]' + PUBMED_CLEAN_FILTER),
        ("source_cid_enabled", "CID", '"Clin Infect Dis"[Journal]' + PUBMED_CLEAN_FILTER),
    ]
    seen_pmids = {c["pmid"] for c in candidates if c.get("pmid")}
    for cfg_key, code, query in journals:
        if not is_enabled(cfg, cfg_key):
            continue
        try:
            time.sleep(pubmed_delay)
            ids = fetch_pubmed_ids(query, paper_days, fetch_limit, ncbi_key)
            total_fetched += len(ids)
            pending = []
            for pid in ids:
                if f"{code}_{pid}" in existing_uids or f"GUIDE_{pid}" in existing_uids or pid in seen_pmids:
                    total_dup += 1
                else:
                    pending.append(pid)
            if pending:
                time.sleep(pubmed_delay)
                for a in fetch_pubmed_articles(pending, ncbi_key):
                    if a["pmid"] in seen_pmids:
                        continue
                    seen_pmids.add(a["pmid"])
                    candidates.append(
                        {
                            "uid": f"{code}_{a['pmid']}",
                            "source": code,
                            "contentType": "paper",
                            "pubDate": a["pubDate"],
                            "title": a["title"],
                            "rawText": a["abstractText"],
                            "pmid": a["pmid"],
                            "doi": a["doi"],
                            "sourceUrl": f"https://pubmed.ncbi.nlm.nih.gov/{a['pmid']}/",
                            "accessStatus": a["accessStatus"],
                        }
                    )
        except Exception as e:
            fetch_errors.append(f"{code}: {e}")

    if is_enabled(cfg, "source_acm_enabled"):
        try:
            for it in fetch_acm_crossref(paper_days):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"ACM: {e}")

    print(f"[수집 완료] 조회 {total_fetched}건 / 중복 {total_dup}건 / 신규 요약 대상 {len(candidates)}건")

    outbreak_q = [x for x in candidates if x["contentType"] == "outbreak"]
    deep_q = [x for x in candidates if x["contentType"] != "outbreak"]
    new_count = 0
    ai_errors: List[str] = []

    def process_q(queue: List[Dict[str, Any]], batch_size: int, model_name: str, is_deep: bool):
        nonlocal new_count
        for offset in range(0, len(queue), batch_size):
            chunk = queue[offset : offset + batch_size]
            try:
                ai_map = (
                    summarize_deep_batch(chunk, model_name, gemini_key, deep_thinking)
                    if is_deep
                    else summarize_outbreak_batch(chunk, model_name, gemini_key)
                )
                rows_to_append = []
                for it in chunk:
                    ai = ai_map.get(it["uid"])
                    if not ai:
                        continue
                    week_key = get_iso_week_key(it["pubDate"] or today_str)
                    note = format_clinical_note(
                        ai.get("ai_clinical_note", ""), ai.get("importance", "Medium"), ai.get("high_axis", ""), is_deep
                    )
                    rows_to_append.append(
                        [
                            it["uid"],
                            it["source"],
                            it["pubDate"],
                            now_str,
                            it["title"],
                            ai.get("title_ko") or it["title"],
                            ai.get("topic_category") or "기타",
                            ai.get("importance") or "Medium",
                            ai.get("fact_summary_ko") or "",
                            note,
                            it.get("pmid", ""),
                            it.get("doi", ""),
                            it.get("sourceUrl", ""),
                            it.get("accessStatus", "Public"),
                            False,
                            it["contentType"],
                            week_key,
                            f"{CODE_VERSION}|{model_name}",
                            ai.get("importance_reason", ""),
                        ]
                    )
                if rows_to_append:
                    item_ws.append_rows(rows_to_append, value_input_option="RAW")
                    new_count += len(rows_to_append)
                    print(f"  + {len(rows_to_append)}건 저장 ({model_name}, 누적 {new_count}/{len(candidates)})")
                time.sleep(1.5 if is_deep else 1.0)
            except Exception as e:
                ai_errors.append(f"{chunk[0]['uid']}: {e}")
                print(f"  [요약 오류] {chunk[0]['uid']}: {e}")

    process_q(outbreak_q, GEMINI_LITE_BATCH_SIZE, lite_model, False)
    process_q(deep_q, GEMINI_DEEP_BATCH_SIZE, deep_model, True)

    notes = status_notes + [f"GitHub Actions({CODE_VERSION}) · 상위:{deep_model}/{deep_thinking}, 경량:{lite_model}"]
    if not candidates and not fetch_errors:
        notes.append("신규 항목 없음")
    if fetch_errors:
        notes.append("조회 오류: " + " | ".join(fetch_errors[:3]))
    if ai_errors:
        notes.append("요약 오류: " + " | ".join(ai_errors[:3]))
    status = "Success" if (not fetch_errors and not ai_errors) else ("Partial" if new_count > 0 else "Fail")
    append_run_log(ss, run_id, now_str, "ALL", status, total_fetched, new_count, total_dup, " / ".join(notes))

    sync_sheet_to_dashboard_json(ss)


def mode_resummarize_all(force: bool = False) -> None:
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    item_ws = ss.worksheet(SHEET_ITEMS)
    ensure_item_meta_columns(item_ws)

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()
    if not gemini_key:
        raise RuntimeError("GEMINI_API_KEY 가 설정되지 않았습니다.")

    lite_model, deep_model, deep_thinking = get_gemini_models(cfg)
    run_id = "GH_RESUM_" + now_kst().strftime("%Y%m%d_%H%M%S")
    now_str = now_kst_str()

    all_values = item_ws.get_all_values()
    if len(all_values) < 2:
        print("재요약할 데이터가 없습니다.")
        return

    deep_targets = []
    outbreak_targets = []
    for idx, row in enumerate(all_values[1:], start=2):
        r = row + [""] * max(0, 19 - len(row))
        uid = r[0].strip()
        if not uid:
            continue
        ctype = r[15].strip() or "paper"
        meta = r[17].strip()
        already_done = (not force) and (
            meta.startswith(f"{CODE_VERSION}|") or (ctype != "outbreak" and deep_model in meta)
        )
        if already_done:
            continue
        rec = {
            "rowIndex": idx,
            "uid": uid,
            "source": r[1].strip(),
            "pubDate": r[2].strip()[:10],
            "title": r[4].strip(),
            "factOld": r[8].strip(),
            "pmid": clean_pmid(r[10]),
            "doi": r[11].strip(),
            "contentType": ctype,
        }
        if ctype == "outbreak":
            outbreak_targets.append(rec)
        else:
            deep_targets.append(rec)

    print(f"[재요약 대상] 논문·가이드라인 {len(deep_targets)}건 ({deep_model}/{deep_thinking}) / 감염병 경보 {len(outbreak_targets)}건 ({lite_model})")

    # 1) PubMed / Crossref 원문 초록 일괄 재조회
    pmids = [t["pmid"] for t in deep_targets if t["pmid"]]
    abs_by_pmid: Dict[str, str] = {}
    if pmids:
        print(f"  PubMed 원문 초록 {len(pmids)}건 일괄 재조회 중...")
        for art in fetch_pubmed_articles(pmids, ncbi_key):
            abs_by_pmid[str(art["pmid"])] = art["abstractText"]

    for t in deep_targets:
        if t["pmid"] and t["pmid"] in abs_by_pmid:
            t["rawText"] = abs_by_pmid[t["pmid"]]
        elif t["doi"]:
            t["rawText"] = fetch_crossref_abstract(t["doi"])
            time.sleep(0.15)
        else:
            t["rawText"] = ""

    updated_deep = 0
    errors: List[str] = []
    for offset in range(0, len(deep_targets), GEMINI_DEEP_BATCH_SIZE):
        chunk = deep_targets[offset : offset + GEMINI_DEEP_BATCH_SIZE]
        try:
            ai_map = summarize_deep_batch(chunk, deep_model, gemini_key, deep_thinking)
            batch_updates = []
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                note = format_clinical_note(
                    ai.get("ai_clinical_note", ""), ai.get("importance", "Medium"), ai.get("high_axis", ""), True
                )
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "기타",
                            ai.get("importance") or "Medium",
                            ai.get("fact_summary_ko") or "",
                            note,
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:S{row_i}",
                        "values": [[f"{CODE_VERSION}|{deep_model}", ai.get("importance_reason", "")]],
                    }
                )
                updated_deep += 1
            if batch_updates:
                item_ws.batch_update(batch_updates, value_input_option="RAW")
            print(f"  [논문·가이드라인] {updated_deep}/{len(deep_targets)}건 갱신 완료")
            time.sleep(1.5)
        except Exception as e:
            errors.append(f"{chunk[0]['uid']}: {e}")
            print(f"  [오류] {chunk[0]['uid']}: {e}")

    # 2) 감염병 경보 정비
    updated_ob = 0
    for offset in range(0, len(outbreak_targets), GEMINI_LITE_BATCH_SIZE):
        chunk = outbreak_targets[offset : offset + GEMINI_LITE_BATCH_SIZE]
        for t in chunk:
            t["rawText"] = t["factOld"] or t["title"]
        try:
            ai_map = summarize_outbreak_batch(chunk, lite_model, gemini_key)
            batch_updates = []
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                note = format_clinical_note(ai.get("ai_clinical_note", ""), ai.get("importance", "Medium"), "", False)
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "국내 긴급경보(KDCA)",
                            ai.get("importance") or "Medium",
                            ai.get("fact_summary_ko") or t["rawText"],
                            note,
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:S{row_i}",
                        "values": [[f"{CODE_VERSION}|{lite_model}", ai.get("importance_reason", "")]],
                    }
                )
                updated_ob += 1
            if batch_updates:
                item_ws.batch_update(batch_updates, value_input_option="RAW")
            print(f"  [감염병 경보] {updated_ob}/{len(outbreak_targets)}건 갱신 완료")
            time.sleep(1.0)
        except Exception as e:
            errors.append(f"{chunk[0]['uid']}: {e}")
            print(f"  [오류] {chunk[0]['uid']}: {e}")

    total_updated = updated_deep + updated_ob
    total_targets = len(deep_targets) + len(outbreak_targets)
    note = f"전체 재요약({CODE_VERSION}) 완료: 논문·가이드라인 {updated_deep}/{len(deep_targets)}, 경보 {updated_ob}/{len(outbreak_targets)}"
    if errors:
        note += " / 오류: " + " | ".join(errors[:3])
    status = "Success" if not errors else ("Partial" if total_updated > 0 else "Fail")
    append_run_log(ss, run_id, now_str, "RESUM_ALL", status, total_targets, total_updated, 0, note)
    sync_sheet_to_dashboard_json(ss)


def mode_build_from_csv(csv_dir: Path) -> None:
    def load_csv(name: str) -> List[Dict[str, Any]]:
        p = csv_dir / name
        if not p.exists():
            return []
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))

    items = load_csv("임상미생물_정보_수집_DB_MVP_JCM_items.csv")
    surv = load_csv("임상미생물_정보_수집_DB_MVP_JCM_kdca_surveillance.csv")
    hai = load_csv("임상미생물_정보_수집_DB_MVP_JCM_hai_stats.csv")
    logs = load_csv("임상미생물_정보_수집_DB_MVP_JCM_run_logs.csv")
    cfg_rows = load_csv("임상미생물_정보_수집_DB_MVP_JCM_config.csv")
    cfg_map = {safe_str(r.get("key")): safe_str(r.get("value")) for r in cfg_rows if safe_str(r.get("key"))}
    write_dashboard_json(items, surv, hai, logs, cfg_map)


def main():
    parser = argparse.ArgumentParser(description="임상미생물 주간 브리핑 파이프라인 v8.0")
    parser.add_argument(
        "--mode",
        choices=["collect", "resummarize_all", "sync_sheet_only", "test_apis", "build_from_csv"],
        default="collect",
    )
    parser.add_argument("--force", action="store_true", help="이미 재요약된 항목도 강제 재요약")
    parser.add_argument("--csv-dir", default=".", help="build_from_csv 모드에서 CSV가 있는 폴더")
    args = parser.parse_args()

    if args.mode == "build_from_csv":
        mode_build_from_csv(Path(args.csv_dir))
    elif args.mode == "test_apis":
        mode_test_apis()
    elif args.mode == "sync_sheet_only":
        ss = connect_spreadsheet()
        sync_sheet_to_dashboard_json(ss)
    elif args.mode == "resummarize_all":
        mode_resummarize_all(force=args.force)
    else:
        mode_collect()


if __name__ == "__main__":
    main()
