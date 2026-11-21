#!/usr/bin/env python3
"""
임상미생물 주간 브리핑 자동 수집 · 4단계 100점 정량 채점 · 503 자동 우회 엔진 (v8.1)
- Google Sheets(원본 DB) + GitHub Actions(무제한 수집·요약) + GitHub Pages(초고속 4탭 대시보드)
- v8.1 핵심 개선:
  1) 503 과부하·타임아웃 자동 우회(Fallback): 논문·가이드라인 2건 배치 + low 추론 + primary 실패 시 fallback(gemini-3.6-flash) 자동 전환
  2) 4단계 등급 + 100점 정량 채점: AI는 4개 축 점수(각 0~25점)만 채점, 코드가 config 컷오프로 Must-Read/High/Medium/Low 기계 판정
  3) 2026년 1월 전체 확대: 4대 저널(retmax=1000)·가이드라인·WHO·KDCA 긴급경보 + KDCA 표본감시 주간소식지 다중 페이지(1주차~) 소급 수집
  4) reclassify_only 모드: config 시트 컷오프 숫자 변경 시 AI 비용 0원으로 전체 등급 즉시 재계산
"""

import argparse
import base64
import csv
import datetime
import html
import json
import os
import re
import time
import urllib.parse
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import requests

CODE_VERSION = "v8.1"
KST = datetime.timezone(datetime.timedelta(hours=9))

SHEET_ITEMS = "items"
SHEET_LOGS = "run_logs"
SHEET_CONFIG = "config"
SHEET_KDCA_SURV = "kdca_surveillance"
SHEET_HAI_STATS = "hai_stats"

GEMINI_DEEP_BATCH_SIZE = 2
GEMINI_LITE_BATCH_SIZE = 5
KDCA_EID_BASE_URL = "https://apis.data.go.kr/1790387/EIDAPIService01"
DEFAULT_PAPER_START_DATE = "2026-01-01"
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
HTTP.headers.update({"User-Agent": "ClinicalMicroBriefing/8.1 (Academic Surveillance Bot)"})


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
    return max(1, diff + 2)


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


def clamp_score(val: Any, hi: int = 25) -> int:
    n = safe_int(val, 0)
    return max(0, min(hi, n))


def get_score_cutoffs(cfg: Dict[str, str]) -> Dict[str, int]:
    return {
        "must_read": safe_int(cfg.get("cutoff_must_read"), 85),
        "high": safe_int(cfg.get("cutoff_high"), 70),
        "medium": safe_int(cfg.get("cutoff_medium"), 45),
    }


def score_to_tier(total_score: int, cutoffs: Dict[str, int]) -> str:
    if total_score >= cutoffs["must_read"]:
        return "Must-Read"
    if total_score >= cutoffs["high"]:
        return "High"
    if total_score >= cutoffs["medium"]:
        return "Medium"
    return "Low"


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
    """items 시트 18~21열(summary_meta, importance_reason, importance_score, score_breakdown) 자동 확보"""
    header = ws.row_values(1)
    if ws.col_count < 21:
        ws.resize(cols=21)
    col_defs = [
        (18, "R1", "summary_meta"),
        (19, "S1", "importance_reason"),
        (20, "T1", "importance_score"),
        (21, "U1", "score_breakdown"),
    ]
    updates = []
    for col_idx, cell, name in col_defs:
        if len(header) < col_idx or not header[col_idx - 1].strip():
            updates.append({"range": cell, "values": [[name]]})
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


def get_gemini_models(cfg: Dict[str, str]) -> Tuple[str, str, str, str]:
    lite = cfg.get("gemini_model", "").strip() or "gemini-3.5-flash-lite"
    deep = cfg.get("gemini_model_deep", "").strip() or "gemini-4-hydrogen"
    fallback = cfg.get("gemini_model_fallback", "").strip() or "gemini-3.6-flash"
    thinking = cfg.get("gemini_thinking_level_deep", "").strip().lower() or "low"
    return lite, deep, fallback, thinking


def append_run_log(ss, run_id: str, ts: str, source: str, status: str, fetched: int, added: int, dup: int, note: str):
    ws = get_or_create_worksheet(
        ss,
        SHEET_LOGS,
        ["run_id", "timestamp", "source", "status", "fetched_count", "new_count", "duplicate_count", "error_message"],
    )
    ws.append_row([run_id, ts, source, status, fetched, added, dup, note or ""], value_input_option="RAW")


# =====================================================================
# Gemini API (503 Auto-Fallback + 4-Tier 100-Point Quantitative Rubric)
# =====================================================================
def probe_gemini_model(model_name: str, api_key: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{urllib.parse.quote(model_name)}:generateContent"
    gen_cfg: Dict[str, Any] = {"temperature": 0}
    if re.match(r"^gemini-[34]", model_name, re.IGNORECASE):
        gen_cfg["thinkingConfig"] = {"thinkingLevel": "low"}
    payload = {"contents": [{"role": "user", "parts": [{"text": "Reply with OK only."}]}], "generationConfig": gen_cfg}
    try:
        res = HTTP.post(url, headers={"x-goog-api-key": api_key}, json=payload, timeout=(15, 45))
        if res.status_code == 200:
            return "성공 (HTTP 200)"
        return f"실패 (HTTP {res.status_code}: {res.text[:140]})"
    except Exception as e:
        return f"연결 오류 ({str(e)[:120]})"


def _call_single_gemini_model(
    prompt: str,
    response_schema: Dict[str, Any],
    model_name: str,
    api_key: str,
    temperature: float,
    thinking_level: str,
    max_attempts: int = 3,
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
    for attempt in range(1, max_attempts + 1):
        try:
            res = HTTP.post(url, headers={"x-goog-api-key": api_key}, json=payload, timeout=(15, 95))
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

            last_err = f"HTTP {res.status_code}: {res.text[:200]}"
            if res.status_code in (429, 500, 502, 503, 504) and attempt < max_attempts:
                wait_sec = 5 * attempt
                print(f"  [재시도 {attempt}/{max_attempts}] {model_name} HTTP {res.status_code}, {wait_sec}초 대기...")
                time.sleep(wait_sec)
                continue
            raise RuntimeError(f"{model_name} {last_err}")
        except (requests.RequestException, json.JSONDecodeError) as e:
            last_err = str(e)
            if attempt < max_attempts:
                wait_sec = 5 * attempt
                print(f"  [연결 재시도 {attempt}/{max_attempts}] {model_name} 오류, {wait_sec}초 대기...")
                time.sleep(wait_sec)
                continue
            raise RuntimeError(f"{model_name} 호출 실패: {last_err[:180]}") from e
    raise RuntimeError(f"{model_name} 호출 실패: {last_err[:180]}")


def call_gemini_with_fallback(
    prompt: str,
    response_schema: Dict[str, Any],
    primary_model: str,
    fallback_model: str,
    api_key: str,
    temperature: float = 0.1,
    thinking_level: str = "low",
) -> Tuple[Dict[str, Dict[str, Any]], str]:
    try:
        res_map = _call_single_gemini_model(
            prompt, response_schema, primary_model, api_key, temperature, thinking_level, max_attempts=3
        )
        return res_map, primary_model
    except Exception as primary_err:
        if fallback_model and fallback_model != primary_model:
            print(f"  [503/지연 자동 우회] {primary_model} 실패({str(primary_err)[:80]}) -> {fallback_model} 전환 실행")
            fb_thinking = "low" if re.match(r"^gemini-[34]", fallback_model, re.IGNORECASE) else ""
            res_map = _call_single_gemini_model(
                prompt, response_schema, fallback_model, api_key, temperature, fb_thinking, max_attempts=3
            )
            return res_map, fallback_model
        raise


def format_clinical_note(raw_note: str, tier: str, primary_axis: str, total_score: int, is_deep: bool) -> str:
    clean = str(raw_note or "").strip()
    clean = re.sub(r"^(\[?\s*AI\s*(임상\s*)?(참고\s*)?해석\s*\]?\s*[:：\-]?\s*)", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^AI의\s*(임상\s*)?(참고\s*)?해석에\s*따르면[,]?\s*", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^본\s*원저는\s*", "", clean, flags=re.IGNORECASE)
    if is_deep and tier == "Must-Read" and primary_axis and primary_axis != "해당없음":
        return f"[AI 해석 · 필독 {total_score}점({primary_axis})] {clean}"
    return f"[AI 해석] {clean}"


def summarize_deep_batch(
    chunk: List[Dict[str, Any]],
    primary_model: str,
    fallback_model: str,
    api_key: str,
    thinking_level: str,
    cutoffs: Dict[str, int],
) -> Tuple[Dict[str, Dict[str, Any]], str]:
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
            "당신은 대한민국 대학병원 진단검사의학과(임상미생물학·분자진단·항균제 내성) 전문의를 지원하는 수석 학술 평가 AI입니다.",
            f'아래 {len(chunk)}편의 논문 또는 가이드라인을 분석하여 JSON 객체({{ "items": [...] }})로만 응답하십시오. 각 항목의 uid를 정확히 유지하십시오.',
            "",
            "[절대 준수 원칙]",
            '1. 원문 왜곡 및 환각 금지: 초록 원문에 명시되지 않은 표본 수(n), 민감도/특이도, 균주, 결론을 절대 추측하거나 지어내지 마십시오. 초록이 없거나 극히 짧으면 fact_summary_ko 첫 문장에 "초록 정보 부족으로 원문 확인 필요."를 명시하십시오.',
            "2. fact_summary_ko (원문 사실 요약, 3~4문장):",
            "   - 연구 목적·설계(다기관/단일기관, 전향/후향, 검체·균주 수 n), 초록에 명시된 핵심 수치(민감도·특이도·일치도·TAT·OR·p값), 핵심 결론을 객관적으로 압축하십시오.",
            "3. ai_clinical_note (진단검사의학과 전문의 실무 포인트, 2~3문장):",
            '   - 금지 상용구: "참고가 될 것으로 사료됩니다", "중요성을 시사합니다", "유용한 도구가 될 수 있습니다", "기초 자료로 활용될 수 있습니다".',
            "   - 반드시 (a) 대학병원 임상미생물 검사실(검사 세팅·분자진단·MALDI-TOF·AST 판독/보고·정도관리·감염관리) 실무 적용점 1가지와 (b) 연구 한계·검증 주의사항 1가지를 구체적으로 적으십시오.",
            "4. [엄격한 100점 정량 채점 기준 — 각 항목 0~25점 정수, 점수 인플레이션 엄금]:",
            "   최종 등급(필독/주목/참고/단신)은 프로그램이 총점으로 자동 판정합니다. 전체 논문의 상위 5~10%만 85점 이상이 되도록 냉정하게 채점하십시오.",
            "   - score_practice (실무·프로토콜 변경력, 0~25):",
            "     · 22~25: IDSA/ESCMID 공식 가이드라인 제·개정, CLSI/EUCAST breakpoint 변경 근거, AST 중대 오류(VME/ME) 주의보",
            "     · 15~21: 검사실 워크플로우·TAT·보고체계에 실질적 개선을 주는 검증",
            "     · 8~14: 간접적 참고 가치 / 0~7: 프로토콜 변화 없음, 단순 리뷰·사설",
            "   - score_evidence (연구 설계·근거 강도, 0~25):",
            "     · 22~25: 공식 학회 지침, 대규모 다기관 전향(n>=300), 대규모 RCT, 체계적 메타분석",
            "     · 14~21: 중규모 원저(n=150~299) 또는 탄탄한 다기관 후향 연구",
            "     · 7~13: 단일기관 소규모(n<150) 예비평가, 후향 관찰 / 0~6: 증례보고(Case Report), 기초/동물실험",
            "   - score_korea (국내 대학병원 임상미생물 적용성, 0~25):",
            "     · 22~25: CRE·카바페네마제, C. auris, 혈액배양 신속동정/AST, 결핵/NTM, 호흡기·CSF 분자진단, MALDI-TOF, 주요 항균제 내성 직결",
            "     · 13~21: 일반 임상 감염·진단 참고 가능",
            "     · 0~12: 국내 발생 극히 드문 해외 풍토병(샤가스·리슈마니아·열대기생충 등), 저소득국 한정 정책·역학, 수의·비인체 균주",
            "   - score_innovation (혁신성·임상 파급력, 0~25):",
            "     · 22~25: 게임체인저급 진단 혁신, 신종 다제내성 기전 최초 규명, 놓치면 치명적인 진단 함정",
            "     · 13~21: 의미 있는 기술 개선 또는 대규모 역학 근거",
            "     · 0~12: 기존 장비 소폭 비교, 반복적 지역 감시 보고",
            "   ※ 초록 원문이 없는 일반 논문(공식 학회 가이드라인 제외)은 4개 항목 모두 최대 14점 이하(총점 56점 이하)로 제한하십시오.",
            "5. topic_category (목록 중 정확히 하나 선택):",
            '   - guideline: "IDSA 지침" | "ESCMID·국제 지침" | "국내 지침"',
            '   - paper: "항균제 내성·AST" | "분자진단·신기술" | "임상 감염·치료" | "감염관리·역학" | "기타"',
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
                        "primary_axis": {
                            "type": "STRING",
                            "enum": [
                                "진단·검사법 혁신",
                                "항균제 내성·AST",
                                "가이드라인·검사실 QC",
                                "원내감염·유행 관리",
                                "해당없음",
                            ],
                        },
                        "score_practice": {"type": "INTEGER"},
                        "score_evidence": {"type": "INTEGER"},
                        "score_korea": {"type": "INTEGER"},
                        "score_innovation": {"type": "INTEGER"},
                        "importance_reason": {"type": "STRING"},
                        "fact_summary_ko": {"type": "STRING"},
                        "ai_clinical_note": {"type": "STRING"},
                    },
                    "required": [
                        "uid",
                        "title_ko",
                        "topic_category",
                        "primary_axis",
                        "score_practice",
                        "score_evidence",
                        "score_korea",
                        "score_innovation",
                        "importance_reason",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }

    raw_map, used_model = call_gemini_with_fallback(
        prompt, schema, primary_model, fallback_model, api_key, 0.1, thinking_level
    )

    # Mechanical score summation & tier classification
    enriched: Dict[str, Dict[str, Any]] = {}
    for uid, ai in raw_map.items():
        s1 = clamp_score(ai.get("score_practice"), 25)
        s2 = clamp_score(ai.get("score_evidence"), 25)
        s3 = clamp_score(ai.get("score_korea"), 25)
        s4 = clamp_score(ai.get("score_innovation"), 25)
        total = s1 + s2 + s3 + s4
        ai["importance_score"] = total
        ai["importance"] = score_to_tier(total, cutoffs)
        ai["score_breakdown"] = f"실무:{s1} | 근거:{s2} | 국내:{s3} | 혁신:{s4}"
        enriched[uid] = ai
    return enriched, used_model


def summarize_outbreak_batch(
    chunk: List[Dict[str, Any]],
    lite_model: str,
    fallback_model: str,
    api_key: str,
    cutoffs: Dict[str, int],
) -> Tuple[Dict[str, Dict[str, Any]], str]:
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
            "당신은 대한민국 대학병원 진단검사의학과 전문의를 지원하는 감염병 감시 평가 AI입니다.",
            f'아래 {len(chunk)}건의 국내외 감염병 동향·경보(KDCA/WHO)를 분석해 JSON 객체({{ "items": [...] }})로 반환하십시오.',
            "",
            "[준수 원칙]",
            '1. 원문에 없는 수치·국가·방역 조치를 절대 생성하지 마십시오. 학술 논문이 아니므로 "본 원저는" 같은 표현을 쓰지 마십시오.',
            "2. fact_summary_ko:",
            "   - KDCA 표본감시 주간소식지: 인플루엔자(ILI)·코로나19 입원·호흡기 바이러스 상위·장관감염증 핵심 수치를 줄바꿈(•) 4줄 이내로 압축하십시오.",
            "   - KDCA 경보 / WHO DONs: 발생 국가·지역, 확진/사망 규모(치명률), 실험실 확진 검사법(PCR·항체·mNGS), 방역 권고를 2~3문장으로 요약하십시오.",
            "3. ai_clinical_note (1~2문장):",
            '   - 상용구("중요성을 시사합니다" 등)를 피하고, 국내 대학병원 진단검사실·감염관리실의 구체적 대응 포인트(검체 채취, PCR/배양 감별, 여행력 문진, 격리 등)를 작성하십시오.',
            "4. [감염병 경보 100점 정량 채점 — 각 0~25점 정수, 무분별한 고득점 금지]:",
            "   - score_domestic (국내 유입·발생 직접성, 0~25): 국내 발생·전국 유행주의보·국내 CRE/다제내성 지침=22~25 / 인접국·주요 여행지·항공 유입 가능=14~20 / 원거리 국지 발생=0~12",
            "   - score_severity (병원체 위험도·급증 강도, 0~25): 1급 감염병·콜레라·고치명률 출혈열·중증 급증=20~25 / 중등도 유행 증가=12~19 / 정례 보고·완만 추세=0~11",
            "   - score_action (국내 검사실·감염관리 즉각 대응 필요성, 0~25): 즉시 감별검사·선제격리·프로토콜 변경 필요=20~25 / 주의 모니터링=12~19 / 단순 참고=0~11",
            "   - score_novelty (신규성·최초 발생 여부, 0~25): 최초 인체감염·10년 만의 국내 발생·비유행국 첫 자국 전파·유행주의보 신규 발령=22~25 / 새로운 국면=12~19 / 동일 해외 유행의 반복 누적 업데이트(예: 콩고 에볼라 주간 반복 보고)·정례 연보 발간·특이사항 없는 정례 표본감시=0~8",
            "5. topic_category (4개 중 하나):",
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
                        "score_domestic": {"type": "INTEGER"},
                        "score_severity": {"type": "INTEGER"},
                        "score_action": {"type": "INTEGER"},
                        "score_novelty": {"type": "INTEGER"},
                        "importance_reason": {"type": "STRING"},
                        "fact_summary_ko": {"type": "STRING"},
                        "ai_clinical_note": {"type": "STRING"},
                    },
                    "required": [
                        "uid",
                        "title_ko",
                        "topic_category",
                        "score_domestic",
                        "score_severity",
                        "score_action",
                        "score_novelty",
                        "importance_reason",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }

    raw_map, used_model = call_gemini_with_fallback(
        prompt, schema, lite_model, fallback_model, api_key, 0.1, ""
    )

    enriched: Dict[str, Dict[str, Any]] = {}
    for uid, ai in raw_map.items():
        s1 = clamp_score(ai.get("score_domestic"), 25)
        s2 = clamp_score(ai.get("score_severity"), 25)
        s3 = clamp_score(ai.get("score_action"), 25)
        s4 = clamp_score(ai.get("score_novelty"), 25)
        total = s1 + s2 + s3 + s4
        ai["importance_score"] = total
        ai["importance"] = score_to_tier(total, cutoffs)
        ai["score_breakdown"] = f"국내영향:{s1} | 위험도:{s2} | 대응시급:{s3} | 신규성:{s4}"
        enriched[uid] = ai
    return enriched, used_model


# =====================================================================
# Source Fetchers (Full 2026 Backfill + KDCA Sentinel Multi-Page)
# =====================================================================
def call_kdca_eid(api_key: str, op_name: str, params: Dict[str, str]) -> List[Dict[str, Any]]:
    safe_key = api_key if "%" in api_key else urllib.parse.quote(api_key, safe="")
    q_parts = [f"serviceKey={safe_key}"]
    for k, v in params.items():
        if v is not None and v != "":
            q_parts.append(f"{urllib.parse.quote(k)}={urllib.parse.quote(str(v))}")
    url = f"{KDCA_EID_BASE_URL}/{op_name}?" + "&".join(q_parts)
    res = HTTP.get(url, timeout=(15, 45))
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
        wk_digits = re.sub(r"\D", "", wk)
        if wk_digits:
            period_key = f"{yr}-W{int(wk_digits):02d}"
    elif yr and mo:
        mo_digits = re.sub(r"\D", "", mo)
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


def fetch_kdca_sentinel_weekly(start_date_str: str, existing_uids: Optional[Set[str]] = None) -> List[Dict[str, Any]]:
    """KDCA 표본감시 주간소식지 다중 페이지 순회 (2026년 1주차부터 전체 소급 수집)"""
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 1, 1)
    existing = existing_uids or set()
    out: List[Dict[str, Any]] = []
    seen_docs: Set[str] = set()

    for page in range(1, 11):
        url = f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbsList.do?q_bbsSn=1010&q_clsfNo=2&q_currPage={page}"
        res = HTTP.get(url, timeout=(15, 30))
        if res.status_code != 200:
            if page == 1:
                raise RuntimeError(f"표본감시 주간소식지 HTTP {res.status_code}")
            break

        page_matches = list(re.finditer(r"q_bbsDocNo=(\d{10,})[^>]*>([^<]{5,180})</a>", res.text))
        if not page_matches:
            break

        unseen_on_page = 0
        any_in_range = False
        for m in page_matches:
            doc_no = m.group(1)
            if doc_no in seen_docs:
                continue
            seen_docs.add(doc_no)
            title = strip_html(m.group(2))
            if not title or "표본감시" not in title:
                continue
            unseen_on_page += 1
            date_str = f"{doc_no[:4]}-{doc_no[4:6]}-{doc_no[6:8]}"
            parsed = parse_flexible_date(date_str)
            if parsed and parsed < cutoff:
                continue
            any_in_range = True

            uid = f"KDCA_SENTINEL_{doc_no}"
            detail_url = f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1010&q_bbsDocNo={doc_no}&q_clsfNo=2"
            body_text = title
            # Skip detail HTTP request for already-collected items to stay fast and polite
            if uid not in existing:
                try:
                    time.sleep(0.25)
                    d_res = HTTP.get(detail_url, timeout=(15, 25))
                    if d_res.status_code == 200:
                        clean = strip_html(d_res.text)
                        idx = clean.find(title[:12])
                        body_text = clean[idx : idx + 2500] if idx > -1 else clean[:2000]
                except Exception:
                    pass

            out.append(
                {
                    "uid": uid,
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
            if len(out) >= 60:
                return out

        if unseen_on_page == 0 or not any_in_range:
            break
        time.sleep(0.25)

    return out


def fetch_kdca_urgent_alerts(start_date_str: str) -> List[Dict[str, Any]]:
    url = "https://www.kdca.go.kr/kdcaBbs/kdca/41/rssList.do?row=300"
    res = HTTP.get(url, timeout=(15, 30))
    if res.status_code != 200:
        raise RuntimeError(f"KDCA RSS HTTP {res.status_code}")
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 1, 1)
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


def fetch_who_dons(start_date_str: str) -> List[Dict[str, Any]]:
    url = "https://www.who.int/api/news/diseaseoutbreaknews?$top=100&$orderby=PublicationDate%20desc"
    res = HTTP.get(url, timeout=(15, 30))
    if res.status_code != 200:
        raise RuntimeError(f"WHO DONs API HTTP {res.status_code}")
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 1, 1)
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


def fetch_acm_crossref(start_date_str: str) -> List[Dict[str, Any]]:
    from_date = (parse_flexible_date(start_date_str) or datetime.date(2026, 1, 1)).strftime("%Y-%m-%d")
    url = f"https://api.crossref.org/journals/2288-6850/works?filter=from-pub-date:{from_date}&rows=100&sort=published&order=desc"
    res = HTTP.get(url, timeout=(15, 30))
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
        res = HTTP.get(f"https://api.crossref.org/works/{urllib.parse.quote(doi)}", timeout=(10, 20))
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
    res = HTTP.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi", params=params, timeout=(15, 35))
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
    for start in range(0, len(pmids), 100):
        sub = pmids[start : start + 100]
        params = {"db": "pubmed", "retmode": "xml", "tool": "ClinicalMicroBriefing", "id": ",".join(sub)}
        if ncbi_api_key:
            params["api_key"] = ncbi_api_key
        res = HTTP.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi", params=params, timeout=(15, 50))
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
# Dashboard JSON Builder
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
        raw_score = safe_float(r.get("importance_score"))
        score_val = int(round(raw_score)) if raw_score is not None else None
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
                "importance_score": score_val,
                "score_breakdown": safe_str(r.get("score_breakdown")),
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

    items.sort(key=lambda x: (x["pub_date"], x["importance_score"] or -1, x["uid"]), reverse=True)

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

    lite_model, deep_model, fallback_model, thinking = get_gemini_models(cfg_map)
    cutoffs = get_score_cutoffs(cfg_map)
    payload = {
        "meta": {
            "version": CODE_VERSION,
            "generated_at": now_kst_str(),
            "total_items": len(items),
            "models": {"deep": deep_model, "fallback": fallback_model, "lite": lite_model, "thinking": thinking},
            "cutoffs": cutoffs,
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
    lite_model, deep_model, fallback_model, thinking = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    print(f"[OK] Google Sheets 연결 성공: {ss.title}")
    print(f"   수집 시작일: 논문={cfg.get('collect_start_date', DEFAULT_PAPER_START_DATE)}, KDCA/WHO={cfg.get('kdca_start_date', DEFAULT_YEAR_START_DATE)}, 한도={cfg.get('pubmed_fetch_limit', '1000')}")
    print(f"   4단계 컷오프: 필독>={cutoffs['must_read']}, 주목>={cutoffs['high']}, 참고>={cutoffs['medium']}")

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()

    if gemini_key:
        print(f"1. Gemini 경량({lite_model}): {probe_gemini_model(lite_model, gemini_key)}")
        print(f"   Gemini 상위({deep_model}, thinking={thinking}): {probe_gemini_model(deep_model, gemini_key)}")
        print(f"   Gemini 우회({fallback_model}): {probe_gemini_model(fallback_model, gemini_key)}")
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


def mode_reclassify_only() -> None:
    """config 시트의 4단계 점수 컷오프로 기존 점수가 있는 항목의 등급만 재계산 (AI 비용 0원)"""
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    cutoffs = get_score_cutoffs(cfg)
    item_ws = ss.worksheet(SHEET_ITEMS)
    ensure_item_meta_columns(item_ws)

    all_values = item_ws.get_all_values()
    if len(all_values) < 2:
        return

    updates = []
    changed = 0
    scored = 0
    for idx, row in enumerate(all_values[1:], start=2):
        r = row + [""] * max(0, 21 - len(row))
        sc = safe_float(r[19])
        if sc is None:
            continue
        scored += 1
        new_tier = score_to_tier(int(round(sc)), cutoffs)
        if r[7].strip() != new_tier:
            updates.append({"range": f"H{idx}", "values": [[new_tier]]})
            changed += 1

    for start in range(0, len(updates), 100):
        item_ws.batch_update(updates[start : start + 100], value_input_option="RAW")
        time.sleep(0.5)

    note = f"컷오프 재분류 완료(AI 비용 0원): 채점 {scored}건 중 {changed}건 등급 변경 (필독>={cutoffs['must_read']}, 주목>={cutoffs['high']}, 참고>={cutoffs['medium']})"
    print(f"[OK] {note}")
    append_run_log(ss, "GH_RECLS_" + now_kst().strftime("%Y%m%d_%H%M%S"), now_kst_str(), "RECLASSIFY", "Success", scored, changed, 0, note)
    sync_sheet_to_dashboard_json(ss)


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

    paper_start = cfg.get("collect_start_date", "").strip() or DEFAULT_PAPER_START_DATE
    year_start = cfg.get("kdca_start_date", "").strip() or DEFAULT_YEAR_START_DATE
    paper_days = calc_days_since(paper_start, 365)
    year_days = calc_days_since(year_start, 365)
    fetch_limit = safe_int(cfg.get("pubmed_fetch_limit"), 1000)

    lite_model, deep_model, fallback_model, deep_thinking = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
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

    # 1) KDCA 표본감시(1주차~ 다중 페이지) · 긴급경보 · WHO
    if is_enabled(cfg, "source_kdca_enabled"):
        try:
            sentinels = fetch_kdca_sentinel_weekly(year_start, existing_uids)
            print(f"[KDCA 표본감시] {year_start} 이후 {len(sentinels)}건 조회")
            for it in sentinels:
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"표본감시소식지: {e}")

        try:
            for it in fetch_kdca_urgent_alerts(year_start):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"KDCA긴급경보: {e}")

    if is_enabled(cfg, "source_who_enabled"):
        try:
            for it in fetch_who_dons(year_start):
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

    # 3) 4대 저널 (JCM · CMI · CID · ACM) - 2026-01-01부터 전체
    journals = [
        ("source_jcm_enabled", "JCM", '"J Clin Microbiol"[Journal]' + PUBMED_CLEAN_FILTER),
        ("source_cmi_enabled", "CMI", '"Clin Microbiol Infect"[Journal]' + PUBMED_CLEAN_FILTER),
        ("source_cid_enabled", "CID", '"Clin Infect Dis"[Journal]' + PUBMED_CLEAN_FILTER),
    ]
    seen_pmids = {c["pmid"] for c in candidates if c.get("pmid")}
    paper_cutoff_date = parse_flexible_date(paper_start)
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
                    # Enforce clean start date floor (allow small EDat vs PubDate margin of 10 days or exact year)
                    pd_obj = parse_flexible_date(a["pubDate"])
                    if paper_cutoff_date and pd_obj and pd_obj < (paper_cutoff_date - datetime.timedelta(days=15)):
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
            for it in fetch_acm_crossref(paper_start):
                total_fetched += 1
                if it["uid"] in existing_uids:
                    total_dup += 1
                else:
                    candidates.append(it)
        except Exception as e:
            fetch_errors.append(f"ACM: {e}")

    print(f"[수집 대상 확정] 기간 {paper_start}~ / 조회 {total_fetched}건 / 중복 {total_dup}건 / 신규 요약 대상 {len(candidates)}건")

    outbreak_q = [x for x in candidates if x["contentType"] == "outbreak"]
    deep_q = [x for x in candidates if x["contentType"] != "outbreak"]
    new_count = 0
    fallback_hits = 0
    ai_errors: List[str] = []

    def process_q(queue: List[Dict[str, Any]], batch_size: int, primary_model: str, is_deep: bool):
        nonlocal new_count, fallback_hits
        for offset in range(0, len(queue), batch_size):
            chunk = queue[offset : offset + batch_size]
            try:
                if is_deep:
                    ai_map, used_model = summarize_deep_batch(
                        chunk, primary_model, fallback_model, gemini_key, deep_thinking, cutoffs
                    )
                else:
                    ai_map, used_model = summarize_outbreak_batch(
                        chunk, primary_model, fallback_model, gemini_key, cutoffs
                    )
                if used_model != primary_model:
                    fallback_hits += len(chunk)

                rows_to_append = []
                for it in chunk:
                    ai = ai_map.get(it["uid"])
                    if not ai:
                        continue
                    week_key = get_iso_week_key(it["pubDate"] or today_str)
                    score = ai.get("importance_score", 0)
                    tier = ai.get("importance", "Medium")
                    note = format_clinical_note(
                        ai.get("ai_clinical_note", ""), tier, ai.get("primary_axis", ""), score, is_deep
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
                            tier,
                            ai.get("fact_summary_ko") or "",
                            note,
                            it.get("pmid", ""),
                            it.get("doi", ""),
                            it.get("sourceUrl", ""),
                            it.get("accessStatus", "Public"),
                            False,
                            it["contentType"],
                            week_key,
                            f"{CODE_VERSION}|{used_model}",
                            ai.get("importance_reason", ""),
                            score,
                            ai.get("score_breakdown", ""),
                        ]
                    )
                if rows_to_append:
                    item_ws.append_rows(rows_to_append, value_input_option="RAW")
                    new_count += len(rows_to_append)
                    print(f"  + {len(rows_to_append)}건 저장 ({used_model}, 누적 {new_count}/{len(candidates)})")
                time.sleep(1.2)
            except Exception as e:
                ai_errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
                print(f"  [요약 오류] {chunk[0]['uid']}: {e}")

    process_q(outbreak_q, GEMINI_LITE_BATCH_SIZE, lite_model, False)
    process_q(deep_q, GEMINI_DEEP_BATCH_SIZE, deep_model, True)

    notes = status_notes + [
        f"GitHub Actions({CODE_VERSION}) · 시작일:{paper_start} · 상위:{deep_model}/{deep_thinking}(우회:{fallback_model})"
    ]
    if fallback_hits > 0:
        notes.append(f"503 우회 처리 {fallback_hits}건")
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

    lite_model, deep_model, fallback_model, deep_thinking = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    run_id = "GH_RESUM_" + now_kst().strftime("%Y%m%d_%H%M%S")
    now_str = now_kst_str()

    all_values = item_ws.get_all_values()
    if len(all_values) < 2:
        print("재요약할 데이터가 없습니다.")
        return

    deep_targets = []
    outbreak_targets = []
    for idx, row in enumerate(all_values[1:], start=2):
        r = row + [""] * max(0, 21 - len(row))
        uid = r[0].strip()
        if not uid:
            continue
        ctype = r[15].strip() or "paper"
        meta = r[17].strip()
        if (not force) and meta.startswith(f"{CODE_VERSION}|"):
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

    print(
        f"[v8.1 재요약 대상] 논문·가이드라인 {len(deep_targets)}건 ({deep_model}/{deep_thinking}, 우회:{fallback_model})"
        f" / 감염병 경보 {len(outbreak_targets)}건 ({lite_model})"
    )

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
    fallback_hits = 0
    errors: List[str] = []

    for offset in range(0, len(deep_targets), GEMINI_DEEP_BATCH_SIZE):
        chunk = deep_targets[offset : offset + GEMINI_DEEP_BATCH_SIZE]
        try:
            ai_map, used_model = summarize_deep_batch(
                chunk, deep_model, fallback_model, gemini_key, deep_thinking, cutoffs
            )
            if used_model != deep_model:
                fallback_hits += len(chunk)
            batch_updates = []
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                score = ai.get("importance_score", 0)
                tier = ai.get("importance", "Medium")
                note = format_clinical_note(
                    ai.get("ai_clinical_note", ""), tier, ai.get("primary_axis", ""), score, True
                )
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "기타",
                            tier,
                            ai.get("fact_summary_ko") or "",
                            note,
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:U{row_i}",
                        "values": [[
                            f"{CODE_VERSION}|{used_model}",
                            ai.get("importance_reason", ""),
                            score,
                            ai.get("score_breakdown", ""),
                        ]],
                    }
                )
                updated_deep += 1
            if batch_updates:
                item_ws.batch_update(batch_updates, value_input_option="RAW")
            print(f"  [논문·가이드라인] {updated_deep}/{len(deep_targets)}건 갱신 ({used_model})")
            time.sleep(1.2)
        except Exception as e:
            errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
            print(f"  [오류] {chunk[0]['uid']}: {e}")

    updated_ob = 0
    for offset in range(0, len(outbreak_targets), GEMINI_LITE_BATCH_SIZE):
        chunk = outbreak_targets[offset : offset + GEMINI_LITE_BATCH_SIZE]
        for t in chunk:
            t["rawText"] = t["factOld"] or t["title"]
        try:
            ai_map, used_model = summarize_outbreak_batch(
                chunk, lite_model, fallback_model, gemini_key, cutoffs
            )
            batch_updates = []
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                score = ai.get("importance_score", 0)
                tier = ai.get("importance", "Medium")
                note = format_clinical_note(ai.get("ai_clinical_note", ""), tier, "", score, False)
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "국내 긴급경보(KDCA)",
                            tier,
                            ai.get("fact_summary_ko") or t["rawText"],
                            note,
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:U{row_i}",
                        "values": [[
                            f"{CODE_VERSION}|{used_model}",
                            ai.get("importance_reason", ""),
                            score,
                            ai.get("score_breakdown", ""),
                        ]],
                    }
                )
                updated_ob += 1
            if batch_updates:
                item_ws.batch_update(batch_updates, value_input_option="RAW")
            print(f"  [감염병 경보] {updated_ob}/{len(outbreak_targets)}건 갱신 ({used_model})")
            time.sleep(1.0)
        except Exception as e:
            errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
            print(f"  [오류] {chunk[0]['uid']}: {e}")

    total_updated = updated_deep + updated_ob
    total_targets = len(deep_targets) + len(outbreak_targets)
    note = (
        f"전체 재요약({CODE_VERSION}) 완료: 논문·가이드라인 {updated_deep}/{len(deep_targets)}, "
        f"경보 {updated_ob}/{len(outbreak_targets)}"
    )
    if fallback_hits > 0:
        note += f" (503 우회 {fallback_hits}건)"
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
    cfg_map.setdefault("gemini_model_fallback", "gemini-3.6-flash")
    cfg_map["gemini_thinking_level_deep"] = "low"
    cfg_map.setdefault("cutoff_must_read", "85")
    cfg_map.setdefault("cutoff_high", "70")
    cfg_map.setdefault("cutoff_medium", "45")
    write_dashboard_json(items, surv, hai, logs, cfg_map)


def main():
    parser = argparse.ArgumentParser(description="임상미생물 주간 브리핑 파이프라인 v8.1")
    parser.add_argument(
        "--mode",
        choices=["collect", "resummarize_all", "reclassify_only", "sync_sheet_only", "test_apis", "build_from_csv"],
        default="collect",
    )
    parser.add_argument("--force", action="store_true", help="이미 v8.1 완료된 항목도 강제 재요약")
    parser.add_argument("--csv-dir", default=".", help="build_from_csv 모드에서 CSV가 있는 폴더")
    args = parser.parse_args()

    if args.mode == "build_from_csv":
        mode_build_from_csv(Path(args.csv_dir))
    elif args.mode == "test_apis":
        mode_test_apis()
    elif args.mode == "reclassify_only":
        mode_reclassify_only()
    elif args.mode == "sync_sheet_only":
        ss = connect_spreadsheet()
        sync_sheet_to_dashboard_json(ss)
    elif args.mode == "resummarize_all":
        mode_resummarize_all(force=args.force)
    else:
        mode_collect()


if __name__ == "__main__":
    main()
