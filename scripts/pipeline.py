#!/usr/bin/env python3
"""
임상미생물 주간 브리핑 자동 수집 · 2-Pass 차등 정밀 요약 · 100점 고정 앵커 루브릭 엔진 (v8.2.1)
- Google Sheets(원본 DB) + GitHub Actions(무제한 수집·요약) + GitHub Pages(0.3초 고속 5탭 메인 대시보드)
- v8.2 핵심 업그레이드:
  1) 100점 완전 고정 앵커 루브릭 (35 / 30 / 20 / 15 가중치):
     실무(35/26/16/6) · 근거(30/22/13/5) · 국내적용(20/14/8/3) · 신규성(15/11/6/2), CID 대규모 치료·백신 RCT 포괄 허용(옵션 B)
  2) 초록 기반 논문 유형별 상한선(Cap): 단일기관 소규모·후향 82 / 서술형 리뷰 78 / 증례·Letter·Editorial 69 / 순수 기초연구 65 (원점수·Cap 투명 보존)
  3) Critical Exception (E1~E4): 필수 3조건 모두 충족 시 원점수·Cap 변경 없이 별도 예외 필독 판정 및 2-Pass 심층 요약 강제 수행
  4) 2-Pass 차등 요약 (안 B):
     - Pass 1(전수): gemini-4-hydrogen (thinking: low)
     - Pass 2(Must-Read 85+ 또는 Critical Exception 확정 전용): gemini-4-hydrogen (thinking: high)
     - 503/지연 자동 폴백: gemini-3.6-flash / 감염병 경보 경량 요약: gemini-3.5-flash-lite
  5) 상반기(2026-01~06) 소급 가드 기본 활성(--allow-h1 미지정 시 2026-07-01 이후 논문 집중)
  6) 본 DB 보호용 pilot_20 층화 파일럿 모드: 별도 시트 탭(pilot_v82)에만 20편 검증 출력
"""

import argparse
import base64
import csv
import datetime
import html
import json
import math
import os
import re
import time
import urllib.parse
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import requests

CODE_VERSION = "v8.3.0"
KST = datetime.timezone(datetime.timedelta(hours=9))

SHEET_ITEMS = "items"
SHEET_LOGS = "run_logs"
SHEET_CONFIG = "config"
SHEET_KDCA_SURV = "kdca_surveillance"
SHEET_HAI_STATS = "hai_stats"
SHEET_PILOT = "pilot_v82"

GEMINI_DEEP_BATCH_SIZE = 2
GEMINI_LITE_BATCH_SIZE = 5
KDCA_EID_BASE_URL = "https://apis.data.go.kr/1790387/EIDAPIService01"
H2_GUARD_START_DATE = "2026-07-01"
DEFAULT_YEAR_START_DATE = "2026-01-01"

# 100-Point Fixed Anchors (35 / 30 / 20 / 15 — Option 1 strict anchor selection)
ANCHORS_PRACTICE = (35, 26, 16, 6)
ANCHORS_EVIDENCE = (30, 22, 13, 5)
ANCHORS_KOREA = (20, 14, 8, 3)
ANCHORS_NOVELTY = (15, 11, 6, 2)

# Study-Type Caps (Abstract-based)
STUDY_TYPE_CAPS: Dict[str, Tuple[Optional[int], str]] = {
    "uncapped": (None, "상한 없음(대규모·지침·RCT·메타분석)"),
    "single_center_small_or_retro": (82, "단일기관 소규모·후향"),
    "narrative_review": (78, "서술형 종설·일반 리뷰"),
    "case_letter_editorial": (69, "증례·Letter·Comment·Editorial"),
    "basic_preclinical": (65, "순수 기초·전임상 연구"),
}

EXCEPTION_LABELS: Dict[str, str] = {
    "E1": "E1(공식 기준·지침 변경)",
    "E2": "E2(중대한 검사 오류·한계)",
    "E3": "E3(신규 고위험 병원체·내성기전)",
    "E4": "E4(긴급 감염병 대응)",
}

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


# =====================================================================
# 2026 법정감염병(91종) + 공통·검사실 KDCA 지침(8종) & CLSI(M·MM)/EUCAST 현황판
# =====================================================================
NOTIFIABLE_DISEASES_MASTER_JSON = "[{\"code\": \"제1급-1\", \"group\": \"1급\", \"name\": \"에볼라바이러스병(Ebola virus disease)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-2\", \"group\": \"1급\", \"name\": \"마버그열(Marburg Hemorrhagic fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-3\", \"group\": \"1급\", \"name\": \"라싸열(Lassa fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-4\", \"group\": \"1급\", \"name\": \"크리미안콩고출혈열(Crimean-Congo Hemorrhagic fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-5\", \"group\": \"1급\", \"name\": \"남아메리카출혈열(South American Hemorrhagic fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-6\", \"group\": \"1급\", \"name\": \"리프트밸리열(Rift Valley fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 체액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"바이러스성 출혈열 대응지침\", \"guidelineKeyword\": \"바이러스성 출혈열\", \"latestTitle\": \"바이러스성 출혈열 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-7\", \"group\": \"1급\", \"name\": \"두창(Smallpox)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 피부병변조직, 피부병변액, 가피 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"생물테러 대비·대응 지침\", \"guidelineKeyword\": \"생물테러\", \"latestTitle\": \"생물테러 대비·대응 지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-8\", \"group\": \"1급\", \"name\": \"페스트(Plague)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 가래, 기관지세척액, 림프절흡인물 등)에서 Y. pestis 분리 동정\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"생물테러 대비·대응 지침\", \"guidelineKeyword\": \"생물테러\", \"latestTitle\": \"생물테러 대비·대응 지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-9\", \"group\": \"1급\", \"name\": \"탄저(Anthrax)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 수포도말, 대변, 가래, 뇌척수액 등)에서 B. anthracis 분리 동정\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"생물테러 대비·대응 지침\", \"guidelineKeyword\": \"생물테러\", \"latestTitle\": \"생물테러 대비·대응 지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-10\", \"group\": \"1급\", \"name\": \"보툴리눔독소증(Botulism)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(대변, 구토물, 위흡인물)에서 독소 생성 C. botulinum 등 분리 동정 ○ 검체(혈액, 대변, 구토물, 위흡인물 등)에서 보툴리눔독소 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"생물테러 대비·대응 지침\", \"guidelineKeyword\": \"생물테러\", \"latestTitle\": \"생물테러 대비·대응 지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-11\", \"group\": \"1급\", \"name\": \"야토병(Tularemia)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(혈액, 가래, 기관지세척액, 림프절흡인물, 조직 등)에서 F. tularensis 분리 동정\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"생물테러 대비·대응 지침\", \"guidelineKeyword\": \"생물테러\", \"latestTitle\": \"생물테러 대비·대응 지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-12\", \"group\": \"1급\", \"name\": \"신종감염병증후군\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"국내 미발생 신종 병원체 의심 증후군(급성출혈열·호흡기·설사·황달·신경증후군) 임상·역학 판정 및 KDCA 확인검사\", \"dept\": \"위기관리총괄과, 신종감염병대응과\", \"guidelineGroup\": \"신종감염병 대응지침\", \"guidelineKeyword\": \"신종감염병\", \"latestTitle\": \"신종감염병 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-13\", \"group\": \"1급\", \"name\": \"중증급성호흡기증후군 (Severe Acute Respiratory Syndrome, SARS)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(구인두도말, 비인두도말, 비인두흡인물, 가래, 기관지폐포세척액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"메르스·중증호흡기 대응지침\", \"guidelineKeyword\": \"메르스\", \"latestTitle\": \"메르스·중증호흡기 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-14\", \"group\": \"1급\", \"name\": \"중동호흡기증후군 (Middle East Respiratory Syndrome: MERS)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(구인두도말, 비인두도말, 비인두흡인물, 비강흡인물, 가래, 기관흡인물, 기관지폐포세척액 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"메르스·중증호흡기 대응지침\", \"guidelineKeyword\": \"메르스\", \"latestTitle\": \"메르스·중증호흡기 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-15\", \"group\": \"1급\", \"name\": \"동물인플루엔자 인체감염증 (Human infection with zoonotic influenza)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(구인두도말, 비인두도말, 비인두흡인물, 가래, 기관지흡인물, 기관지폐포세척액)에서 Animal influenza virus 분리 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(구인두도말, 비인두도말, 비인두흡인물, 가래, 기관지흡인물, 기관지폐포세척액, 안(眼) 점막 도말물)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"조류·동물인플루엔자 인체감염증 대응지침\", \"guidelineKeyword\": \"조류인플루엔자\", \"latestTitle\": \"조류·동물인플루엔자 인체감염증 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-16\", \"group\": \"1급\", \"name\": \"신종인플루엔자(Novel influenza)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(구인두도말, 비인두도말, 비인두흡인물, 가래, 기관지흡인물, 기관지폐포세척액)에서 바이러스 분리 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액)에서 특이 항체 검출 ○ 검체(구인두도말, 비인두도말, 비인두흡인물, 가래, 기관지흡인물, 기관지폐포세척액)에서 특이 유전자 검출 * 새로운 바이러스의 유행...\", \"dept\": \"위기관리총괄과, 신종감염병대응과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"조류·동물인플루엔자 인체감염증 대응지침\", \"guidelineKeyword\": \"조류인플루엔자\", \"latestTitle\": \"조류·동물인플루엔자 인체감염증 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-17\", \"group\": \"1급\", \"name\": \"디프테리아(Diphtheria)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 확인 진단 - 검체(구인두도말, 비인두도말, 비강도말, 위막조직)에서 독소생성 C. diphtheriae 분리 동정 - 검체(구인두도말, 비인두도말, 비강도말, 위막조직)에서 특이(독소유전자 포함) 유전자 검출 ○ 추정 진단 - 검체(구인두도말, 비인두도말, 비강도말, 위막조직 등)에서 C. diphtheriae 분리 동정\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제1급-18\", \"group\": \"1급\", \"name\": \"니파바이러스감염증(Nipah virus infection)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"즉시 신고\", \"lab\": \"○ 검체(인후도말, 비강 및 비인두 흡인물, 뇌척수액, 소변 등)에서 특이 유전자 검출\", \"dept\": \"신종감염병대응과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"니파바이러스감염증 대응지침\", \"guidelineKeyword\": \"니파바이러스\", \"latestTitle\": \"니파바이러스감염증 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-1\", \"group\": \"2급\", \"name\": \"결핵(Tuberculosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(가래, 기관지세척액, 체액, 조직)에서 항산균도말 양성 ○ 검체(가래, 기관지세척액, 체액, 조직)에서 결핵균* 분리 동정 * 특히, M. bovis는 배양에서 동정이 되어야 확진 ○ 검체(가래, 기관지세척액, 체액, 조직)에서 특이 유전자 검출\", \"dept\": \"결핵정책과, 세균분석과(진단)\", \"guidelineGroup\": \"국가결핵관리지침\", \"guidelineKeyword\": \"결핵\", \"latestTitle\": \"국가결핵관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-2\", \"group\": \"2급\", \"name\": \"수두(Varicella)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(수포액, 가피, 구인두도말, 비인두도말, 혈액, 뇌척수액 등)에서 Human alphaherpesvirus 3 분리 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(수포액, 가피, 구인두도말, 비인두도말, 혈액, 뇌척수액 등)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-3\", \"group\": \"2급\", \"name\": \"홍역(Measles)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(비강도말, 구인두도말, 비인두도말, 혈액, 소변 등)에서 Measles morbillivirus 분리 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(비강도말, 구인두도말, 비인두도말, 혈액, 소변 등)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-4\", \"group\": \"2급\", \"name\": \"콜레라(Cholera)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(대변, 직장도말, 구토물)에서 독소형 콜레라균(V. cholerae O1 또는 O139) 분리동정 ○ 추정 진단 - 검체(대변, 직장도말, 구토물 등)에서 독소 유무를 알 수 없는 V. cholerae O1 또는 O139 분리 동정 - 검체(대변, 직장도말, 구토물 등)에서 독소 유전자(ctxA) 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-5\", \"group\": \"2급\", \"name\": \"장티푸스(Typhoid fever)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(대변, 직장도말, 소변, 담즙, 골수, 혈액)에서 S. Typhi 분리 동정 ○ 추정 진단 - 검체(혈액)에서 특이 항체 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-6\", \"group\": \"2급\", \"name\": \"파라티푸스(Paratyphoid fever)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(대변, 직장도말, 소변, 담즙, 골수, 혈액)에서 S. Paratyphi A, B, C 분리 동정 ○ 추정 진단 - 검체(혈액)에서 특이 항체 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-7\", \"group\": \"2급\", \"name\": \"세균성이질(Shigellosis, Bacillary dysentery)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(대변, 직장도말)에서 세균성이질균 분리 동정 ○ 추정 진단 - 검체(대변, 직장도말)에서 특이 유전자(ipaH) 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-8\", \"group\": \"2급\", \"name\": \"장출혈성대장균(Enterohemorrhagic Escherichia coli) 감염증\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 ­ 검체(대변, 직장도말)에서 독소 유전자(stx1, stx2)를 보유한 E. coli 분리 동정 ○ 추정 진단 ­검체(대변, 직장도말)에서 독소 유전자(stx1, stx2) 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-9\", \"group\": \"2급\", \"name\": \"A형간염(Viral hepatitis A)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(혈액, 대변, 직장도말)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"바이러스성 간염 관리지침\", \"guidelineKeyword\": \"간염\", \"latestTitle\": \"바이러스성 간염 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-10\", \"group\": \"2급\", \"name\": \"백일해(Pertussis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(비인두도말, 비인두흡인물, 가래)에서 B. pertussis 분리 동정 ○ 검체(비인두도말, 비인두흡인물, 가래)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-11\", \"group\": \"2급\", \"name\": \"유행성이하선염(Mumps)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(타액, 볼점막도말, 구인두도말, 비인두도말, 혈액, 소변, 뇌척수액)에서 Mumps orthorubulavirus 분리 ○회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○검체(혈액)에서 특이 IgM 항체 검출 ○검체(타액, 볼점막도말, 구인두도말, 비인두도말, 혈액, 소변, 뇌척수액)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-12\", \"group\": \"2급\", \"name\": \"풍진(Rubella)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"① 선천성 풍진 ○ 검체(비강도말, 구인두도말, 비인두도말, 혈액, 소변, 뇌척수액)에서 Rubivirus rubellae 분리 ○ 모체 항체가가 없어지는 시기(생후 약 9개월) 이후에도 항체 지속 검출 ○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(비강도말, 구인두도말, 비인두도말, 혈액, 소변, 뇌척수액)에서 특이 유전자 검...\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-13\", \"group\": \"2급\", \"name\": \"폴리오(Poliomyelitis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(대변, 뇌척수액, 구인두도말, 비인두도말, 혈액)에서 Poliovirus 분리\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-14\", \"group\": \"2급\", \"name\": \"수막구균 감염증(Meningococcal disease)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(뇌척수액, 혈액 등)에서 N. meningitidis 분리 동정\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-15\", \"group\": \"2급\", \"name\": \"b형헤모필루스인플루엔자(Haemophilus influenzae type b)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검조직 등)에서 H. influenzae type b 분리 동정 ○ 추정 진단 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검조직 등)에서 H. influenzae 분리 동정 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검...\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-16\", \"group\": \"2급\", \"name\": \"폐렴구균 감염증(Pneumococcal disease)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검조직 등)에서 S. pneumoniae 분리 동정 ○ 추정 진단 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검조직 등)에서 특이 항원 검출 - 검체(혈액, 뇌척수액, 관절액, 늑막액, 심낭액, 복수, 생검조직 등)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-17\", \"group\": \"2급\", \"name\": \"한센병(Hansen’s disease, Leprosy)\", \"scope\": \"환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 분자생물학적 검사법에 의한 나균 확인 ○ 조직검사상 한센병 육아종 소견 확인 ○ 병변의 도말검사(항산성 염색)에서 항산성균 확인\", \"dept\": \"결핵정책과\", \"guidelineGroup\": \"한센병 사업관리지침\", \"guidelineKeyword\": \"한센병\", \"latestTitle\": \"한센병 사업관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-18\", \"group\": \"2급\", \"name\": \"성홍열(Scarlet fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(구인두도말, 혈액)에서 S. pyogenes 분리 동정 ○ 추정 진단 - 검체(구인두도말)에서 특이 항원 검출\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"호흡기감염병 관리지침\", \"guidelineKeyword\": \"호흡기감염병\", \"latestTitle\": \"호흡기감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-19\", \"group\": \"2급\", \"name\": \"반코마이신내성황색포도알균(VRSA) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"임상검체 S. aureus 중 Vancomycin MIC 4~8(VISA) 또는 ≥16 µg/mL(VRSA) 확인 (혈액=환자, 혈액외=보유자)\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-20\", \"group\": \"2급\", \"name\": \"카바페넴내성장내세균목(CRE) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"임상검체 Enterobacterales 중 Doripenem/Imipenem/Meropenem MIC ≥4(≤19mm) 또는 Ertapenem MIC ≥2 µg/mL(≤18mm) 내성 확인 (CPE 생성 별도 신고)\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제2급-21\", \"group\": \"2급\", \"name\": \"E형간염(Viral hepatitis E)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(혈액, 대변, 직장도말)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"바이러스성 간염 관리지침\", \"guidelineKeyword\": \"간염\", \"latestTitle\": \"바이러스성 간염 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-1\", \"group\": \"3급\", \"name\": \"파상풍(Tetanus)\", \"scope\": \"환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"임상진단 기준 적용 (파상풍에 부합하는 특징적 임상증상)\", \"dept\": \"감염병관리과\", \"guidelineGroup\": \"예방접종 대상 감염병 관리지침\", \"guidelineKeyword\": \"예방접종\", \"latestTitle\": \"예방접종 대상 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-2\", \"group\": \"3급\", \"name\": \"B형간염(Viral hepatitis B)\", \"scope\": \"환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 특이 항원(HBsAg) 및 특이 항체(IgM anti-HBc) 검출 (단, 6개월 전에 B형간염을 진단받았던 자는 제외함) ○ 검체(혈액)에서 특이 항체(IgM anti-HBc) 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"바이러스성 간염 관리지침\", \"guidelineKeyword\": \"간염\", \"latestTitle\": \"바이러스성 간염 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-3\", \"group\": \"3급\", \"name\": \"일본뇌염(Japanese encephalitis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액)에서 Japanese encephalitis virus 분리 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액, 뇌척수액)에서 특이 유전자 검출 ○ 추정진단 - 검체(혈액, 뇌척수액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-4\", \"group\": \"3급\", \"name\": \"C형간염(Viral hepatitis C)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"바이러스성 간염 관리지침\", \"guidelineKeyword\": \"간염\", \"latestTitle\": \"바이러스성 간염 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-5\", \"group\": \"3급\", \"name\": \"말라리아(Malaria)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액)에서 도말검사로 말라리아 원충(충체) 확인 - 검체(혈액)에서 특이 유전자 검출 ○ 추정 진단 - 검체(혈액)에서 특이 항원 검출(신속진단검사 포함)\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"말라리아 관리지침\", \"guidelineKeyword\": \"말라리아\", \"latestTitle\": \"말라리아 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-6\", \"group\": \"3급\", \"name\": \"레지오넬라증(Legionellosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(기관지세척액, 기관지폐포액, 기관지흡인물, 가래, 폐조직, 흉수, 혈액 등)에서 레지오넬라균 분리 동정 - 검체(소변)에서 특이 항원 검출 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 추정 진단 - 검체(기관지세척액, 기관지폐포액, 기관지흡인물, 가래, 폐조직, 흉수, 혈액 등)에서 직접형광항...\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"호흡기감염병 관리지침\", \"guidelineKeyword\": \"호흡기감염병\", \"latestTitle\": \"호흡기감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-7\", \"group\": \"3급\", \"name\": \"비브리오패혈증(Vibrio vulnificus sepsis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액, 대변, 소변, 직장도말, 구토물, 수포액, 조직)에서 V. vulnificus 분리 동정\", \"dept\": \"감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-8\", \"group\": \"3급\", \"name\": \"발진티푸스(Epidemic typhus)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액, 피부조직)에서 R. prowazekii 분리 동정 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액, 피부조직)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-9\", \"group\": \"3급\", \"name\": \"발진열(Murine typhus)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액, 피부조직)에서 R. typhi 분리 동정 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액, 피부조직)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-10\", \"group\": \"3급\", \"name\": \"쯔쯔가무시증(Scrub typhus)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 조직, 가피)에서 O. tsutsugamushi 분리 동정 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액, 조직, 가피)에서 특이 유전자 검출 ○ 추정 진단 - 검체(혈액)에서 특이 항체 검출\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-11\", \"group\": \"3급\", \"name\": \"렙토스피라증(Leptospirosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 소변, 뇌척수액, 조직)에서 렙토스피라균 분리 동정 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 현미경응집법으로 혈청의 단일항체가가 1:800 이상 - 검체(혈액, 소변, 뇌척수액, 조직)에서 특이 유전자 검출 ○ 추정 진단 - 현미경응집법으로 혈청의 단일항체가가 1:200 이상 ∼ 1...\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-12\", \"group\": \"3급\", \"name\": \"브루셀라증(Brucellosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 골수, 관절액, 조직 등)에서 브루셀라균 분리 동정 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 추정 진단 - 급성기 혈청에서 미세응집법으로 항체가 1:160 이상 - 검체(혈액, 골수, 관절액, 조직 등)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"인수공통감염병 관리지침\", \"guidelineKeyword\": \"인수공통\", \"latestTitle\": \"인수공통감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-13\", \"group\": \"3급\", \"name\": \"공수병(Rabies)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(타액)에서 Lyssavirus rabies 분리 ○ 검체(목 피부조직, 뇌조직)에서 특이 항원 검출 ○ 검체(혈액, 뇌척수액)에서 특이 항체 검출 ○ 검체(타액, 목 피부조직, 뇌조직)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"인수공통감염병 관리지침\", \"guidelineKeyword\": \"인수공통\", \"latestTitle\": \"인수공통감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-14\", \"group\": \"3급\", \"name\": \"신증후군출혈열(유행성출혈열) (Hemorrhagic Fever with Renal Syndrome)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액)에서 특이 IgM 항체 검출 - 예방접종을 받지 않은 자 중에 간접면역형광항체법으로 항체가가 1:512 이상 - 검체(혈액, 뇌척수액)에서 특이 유전자 검출 ○ 추정 진단 - 예방접종을 받은 자 중에 간접면역형광항체법으로 특이 IgG 항체 검출 -...\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-15\", \"group\": \"3급\", \"name\": \"후천성면역결핍증(AIDS)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 생후 18개월 미만인 자 - 검체(혈액)에서 p24 특이 항원 검출이면서 항원중화검사 양성 - 검체(혈액)에서 특이 유전자 검출 ○ 생후 18개월 이상인 자 - 검체(혈액)에서 특이 항체 검출(웨스턴블롯법으로 양성인 경우) - 검체(혈액)에서 p24 특이 항원 검출이면서 항원중화검사에서 양성 - 검체(혈액)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"HIV/AIDS 관리지침\", \"guidelineKeyword\": \"HIV\", \"latestTitle\": \"HIV/AIDS 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-16\", \"group\": \"3급\", \"name\": \"크로이츠펠트-야콥병(Creutzfeldt-Jakob disease) 및 변종크로이츠펠트-야콥병 (vCJD)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"신경병리·면역조직화학(IHC)·뇌척수액 14-3-3/RT-QuIC·PRNP 유전자 변이·뇌파/MRI 기준\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"인수공통감염병 관리지침\", \"guidelineKeyword\": \"인수공통\", \"latestTitle\": \"인수공통감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-17\", \"group\": \"3급\", \"name\": \"황열(Yellow fever)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인진단 - 검체(혈액)에서 Yellow fever virus 분리 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액)에서 특이 유전자 검출 ○ 추정진단 - 검체(혈액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-18\", \"group\": \"3급\", \"name\": \"뎅기열(Dengue fever)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액)에서 Dengue virus 분리 - 검체(혈액)에서 특이 항원 검출 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액, 뇌척수액 등)에서 특이 유전자 검출 ○ 추정 진단 - 검체(혈액, 뇌척수액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-19\", \"group\": \"3급\", \"name\": \"큐열(Q fever)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 조직)에서 C. burnetii 분리동정 - 급성 큐열 : 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가(큐열균 phase Ⅱ 항원에 대한 특이항체) - 만성 큐열 : 간접면역형광항체법으로 측정한 phase Ⅰ 항원에 대한 특이 IgG 단일항체가가 1:800 이상이면서 phase Ⅰ항원에 대한 ...\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"인수공통감염병 관리지침\", \"guidelineKeyword\": \"인수공통\", \"latestTitle\": \"인수공통감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-20\", \"group\": \"3급\", \"name\": \"웨스트나일열(West Nile fever)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액, 조직)에서 West Nile virus 분리 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액, 뇌척수액, 조직)에서 특이 유전자 검출 ○ 추정진단 - 검체(혈액, 뇌척수액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-21\", \"group\": \"3급\", \"name\": \"라임병(Lyme Borreliosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액, 피부생검조직)에서 보렐리아속균 분리 동정 - 검체(혈액, 뇌척수액)에서 간접면역형광항체법(또는 ELISA)과 웨스턴블럿법으로 특이 항체 검출 ○ 추정 진단 - 검체(혈액, 뇌척수액, 피부생검조직)에서 간접면역형광항체법(또는 ELISA) 또는 웨스턴블럿법으로 특이 항체 검출\", \"dept\": \"인수공통감염병관리과, 세균분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-22\", \"group\": \"3급\", \"name\": \"진드기매개뇌염(Tick-borne Encephalitis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 뇌척수액)에서 Tick-borne encephalitis virus 분리 - 회복기 혈청의 항체가가 급성기에 비해 4배 이상 증가 - 검체(혈액, 뇌척수액)에서 특이 유전자 검출 ○ 추정 진단 - 검체(혈액, 뇌척수액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-23\", \"group\": \"3급\", \"name\": \"유비저(Melioidosis)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액, 소변, 가래, 농양 등)에서 B. pseudomallei 분리 동정\", \"dept\": \"인수공통감염병관리과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"인수공통감염병 관리지침\", \"guidelineKeyword\": \"인수공통\", \"latestTitle\": \"인수공통감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-24\", \"group\": \"3급\", \"name\": \"치쿤구니야열(Chikungunya fever)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 Chikungunya virus 분리 ○ 회복기 혈청의 항체가가 급성기에 비해 4배 이상 증가 ○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(혈액)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-25\", \"group\": \"3급\", \"name\": \"중증열성혈소판감소증후군 (Severe Fever with Thrombocytopenia Syndrome: SFTS)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(혈액)에서 중증열성혈소판감소증후군 바이러스 분리 ○ 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 ○ 검체(혈액)에서 특이 IgM 항체 검출 ○ 검체(혈액)에서 특이 유전자 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"진드기·설치류 매개 감염병 관리지침\", \"guidelineKeyword\": \"진드기\", \"latestTitle\": \"진드기·설치류 매개 감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-26\", \"group\": \"3급\", \"name\": \"지카바이러스 감염증(Zika virus infection)\", \"scope\": \"환자, 의사환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 확인 진단 - 검체(혈액, 소변 등)에서 Zika virus 분리 - 회복기 혈청의 항체가가 급성기에 비하여 4배 이상 증가 - 검체(혈액, 소변 등)에서 특이 유전자 검출 ○ 추정 진단 - 검체(혈액)에서 특이 IgM 항체 검출\", \"dept\": \"인수공통감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"모기매개감염병 관리지침\", \"guidelineKeyword\": \"모기매개\", \"latestTitle\": \"모기매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-27\", \"group\": \"3급\", \"name\": \"엠폭스(MPOX)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 검체(피부병변액, 피부병변조직, 가피, 구인두도말, 혈액 등)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 고위험병원체분석과(진단)\", \"guidelineGroup\": \"엠폭스 대응지침\", \"guidelineKeyword\": \"엠폭스\", \"latestTitle\": \"엠폭스 대응지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제3급-28\", \"group\": \"3급\", \"name\": \"매독(Syphilis)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"24시간 이내 신고\", \"lab\": \"○ 1기ㆍ2기ㆍ3기 매독 및 조기 잠복 매독 - 검체(경성하감 또는 편평콘딜롬과 같은 피부병변조직, 궤양부위 삼출액(농, 진물))에서 암시야현미경 검사로 매독균 검출 - 검체(혈액, 뇌척수액)에서 트레포네마검사 및 비트레포네마검사 모두 양성 - 검체(경성하감 또는 편평콘딜롬과 같은 피부병변조직, 궤양부위 삼출액(농, 진물), 혈액)에...\", \"dept\": \"에이즈관리과, 세균분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-1\", \"group\": \"4급\", \"name\": \"인플루엔자(Influenza)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"7일 이내 신고(매주 화요일까지)\", \"lab\": \"○ 검체(구인두도말, 비인두도말, 비인두흡입물, 비강흡입물, 기관지폐포세척액, 가래)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"호흡기감염병 관리지침\", \"guidelineKeyword\": \"호흡기감염병\", \"latestTitle\": \"호흡기감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-2\", \"group\": \"4급\", \"name\": \"회충증(Ascariasis)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변)에서 충란 확인 ○ 감염부위에서 충체 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-3\", \"group\": \"4급\", \"name\": \"편충증(Trichuriasis)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변)에서 충란 확인 ○ 감염부위에서 충체 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-4\", \"group\": \"4급\", \"name\": \"요충증(Enterobiasis)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(항문주위도말)에서 충란 및 충체 확인 ○ 항문주위와 여성의 질에서 충체 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-5\", \"group\": \"4급\", \"name\": \"간흡충증(Clonorchiasis)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변)에서 충란 확인 ○ 감염부위에서 충체 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-6\", \"group\": \"4급\", \"name\": \"폐흡충증(Paragonimiasis)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변, 가래)에서 충란 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-7\", \"group\": \"4급\", \"name\": \"장흡충증(Intestinal fluke infections)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변)에서 충란 및 충체 확인\", \"dept\": \"인수공통감염병관리과, 매개체분석과(진단)\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-8\", \"group\": \"4급\", \"name\": \"수족구병(Hand, foot and mouth disease)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변, 직장도말, 뇌척수액, 혈액, 구인두도말, 비인두도말, 비강세척액 등)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"엔테로바이러스·수족구병 관리지침\", \"guidelineKeyword\": \"엔테로바이러스\", \"latestTitle\": \"엔테로바이러스·수족구병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-9\", \"group\": \"4급\", \"name\": \"임질(Gonorrhea)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(남성의 요도도말)에서 세포 내 그람음성 쌍알균 현미경 검사 ○ 검체(요도·자궁경부·직장·인두도말, 결막, 혈액, 관절액)에서 N. gonorrhoeae 분리 동정 ○ 검체(요도·자궁경부·직장·인두도말, 첫 소변, 척수액, 관절액)에서 특이 항원 검출 ○ 검체(자궁경부·질도말, 첫 소변)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 세균분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-10\", \"group\": \"4급\", \"name\": \"클라미디아감염증(Chlamydial infection)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(요도·자궁경부·직장·인두도말)에서 C. trachomatis 분리 동정 ○ 검체(요도·자궁경부·직장·인두도말, 첫 소변)에서 특이 항원 검출 ○ 검체(질도말, 첫 소변 등)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 세균분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-11\", \"group\": \"4급\", \"name\": \"연성하감(Chancroid)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(병변의 분비물, 궤양부위 삼출물)에서 H. ducreyi 분리 동정 ○ 검체(병변의 분비물, 궤양부위 삼출물)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 세균분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-12\", \"group\": \"4급\", \"name\": \"성기단순포진(Genital herpes)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(수포나 궤양 병변 등에서 나오는 분비물, 궤양부위도말)에서 Simplexvirus humanalpha2 또는 Simplexvirus humanalpha1 분리 ○ 검체(혈액, 수포나 궤양 병변 등에서 나오는 분비물, 궤양부위도말)에서 특이 항체 검출 ○ 검체(혈액, 수포나 궤양 병변 등에서 나오는 분비물, 궤양부위도말)에서 ...\", \"dept\": \"에이즈관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-13\", \"group\": \"4급\", \"name\": \"첨규콘딜롬(Condyloma acuminata)\", \"scope\": \"환자, 의사환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(병변조직, 자궁경부세포)에서 Human papillomavirus 감염에 합당한 조직·병리학적 변화 확인 ○ 검체(병변조직, 자궁경부세포)에서 특이 유전자 검출\", \"dept\": \"에이즈관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-14\", \"group\": \"4급\", \"name\": \"반코마이신내성장알균(VRE) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"임상검체 Enterococcus 중 Vancomycin MIC ≥32 µg/mL(≤14mm) 내성 확인 또는 vanA/vanB 유전자 검출 (혈액=환자, 혈액외=보유자)\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-15\", \"group\": \"4급\", \"name\": \"메티실린내성황색포도알균(MRSA) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"임상검체 S. aureus 중 Oxacillin MIC ≥4 또는 Cefoxitin MIC ≥8 µg/mL(≤21mm) 내성 확인 또는 mecA 유전자 검출\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-16\", \"group\": \"4급\", \"name\": \"다제내성녹농균(MRPA) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"임상검체 P. aeruginosa 중 카바페넴계·아미노글리코사이드계·플루오로퀴놀론계 3개 계열 모두 내성 확인\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-17\", \"group\": \"4급\", \"name\": \"다제내성아시네토박터바우마니균(MRAB) 감염증\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"임상검체 A. baumannii 중 카바페넴계·아미노글리코사이드계·플루오로퀴놀론계 3개 계열 모두 내성 확인\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-18\", \"group\": \"4급\", \"name\": \"장관감염증\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"세균 11종(살모넬라·장염비브리오·ETEC/EIEC/EPEC·캄필로박터 등)·바이러스 5종(로타·노로 등)·원충 4종 분리동정 또는 특이유전자/항원 검출\", \"dept\": \"감염병관리과\", \"guidelineGroup\": \"수인성 및 식품매개감염병 관리지침\", \"guidelineKeyword\": \"수인성\", \"latestTitle\": \"수인성 및 식품매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-19\", \"group\": \"4급\", \"name\": \"급성호흡기감염증\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"바이러스 7종(아데노·보카·파라인플루엔자·RSV·리노·hMPV·사람코로나) 및 세균 2종(마이코플라스마·클라미디아 폐렴균) 분리 또는 특이유전자 검출\", \"dept\": \"감염병관리과\", \"guidelineGroup\": \"호흡기감염병 관리지침\", \"guidelineKeyword\": \"호흡기감염병\", \"latestTitle\": \"호흡기감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-20\", \"group\": \"4급\", \"name\": \"해외유입기생충감염증\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"해외유입 기생충 11종(리슈만편모충·바베스열원충·아프리카수면병·주혈흡충·샤가스병·포충·톡소포자충 등) 충체·충란 확인 또는 특이유전자/항체 검출\", \"dept\": \"인수공통감염병관리과\", \"guidelineGroup\": \"기생충감염병 관리지침\", \"guidelineKeyword\": \"기생충\", \"latestTitle\": \"기생충감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-21\", \"group\": \"4급\", \"name\": \"엔테로바이러스감염증(Enterovirus infection)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(대변, 직장도말, 뇌척수액, 혈액, 구인두도말, 비인두도말, 비강세척액)에서 특이 유전자 검출\", \"dept\": \"감염병관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"엔테로바이러스·수족구병 관리지침\", \"guidelineKeyword\": \"엔테로바이러스\", \"latestTitle\": \"엔테로바이러스·수족구병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-22\", \"group\": \"4급\", \"name\": \"사람유두종바이러스 감염증(Human papillomavirus infection)\", \"scope\": \"병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(자궁경부 또는 성기 부위의 병변조직이나 도말물)에서 HPV 특이 유전형 검출 - HPV 특이 유전형 : 16, 18, 31, 33, 35, 39, 45, 51, 52, 56, 58, 59, 68 ※ HPV 특이 유전형에 대한 기준은 IARC 분류 기준(Group 1, Group 2A)에 근거\", \"dept\": \"에이즈관리과, 바이러스분석과(진단)\", \"guidelineGroup\": \"성매개감염병 관리지침\", \"guidelineKeyword\": \"성매개감염병\", \"latestTitle\": \"성매개감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-23\", \"group\": \"4급\", \"name\": \"코로나바이러스감염증-19(COVID-19)\", \"scope\": \"환자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"○ 검체(비인두도말, 구인두도말, 가래 등)에서 특이유전자 검출\", \"dept\": \"감염병관리과, 신종병원체분석과(진단)\", \"guidelineGroup\": \"호흡기감염병 관리지침\", \"guidelineKeyword\": \"호흡기감염병\", \"latestTitle\": \"호흡기감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"제4급-24\", \"group\": \"4급\", \"name\": \"칸디다 오리스 감염증(Candida auris infection)\", \"scope\": \"환자, 병원체보유자\", \"deadline\": \"7일 이내 신고\", \"lab\": \"검체(귀 검체 제외, 혈액·소변·감시배양 포함)에서 Candida auris 분리 동정 (2026.3.29. 제4급 신규 지정)\", \"dept\": \"의료감염관리과, 세균분석과(진단)\", \"guidelineGroup\": \"의료관련감염병 관리지침\", \"guidelineKeyword\": \"의료관련감염병\", \"latestTitle\": \"의료관련감염병 관리지침 (KDCA 지침 게시판 확인)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-1\", \"group\": \"공통·검사실\", \"name\": \"법정감염병 진단·신고기준 (2026.5 고시 제2026-4호)\", \"scope\": \"전 의료기관 · 병원체 확인기관\", \"deadline\": \"상시 기준 (매년/고시 개정 시)\", \"lab\": \"1~4급 법정감염병 91종(2026.3.29. 제4급 칸디다 오리스 감염증 신설 반영) 신고범위·진단검사기준·표본감시 기준 총괄\", \"dept\": \"감염병관리과\", \"guidelineGroup\": \"법정감염병 진단·신고기준\", \"guidelineKeyword\": \"법정감염병 진단\", \"latestTitle\": \"2026 법정감염병 진단·신고기준 (2026.05)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-2\", \"group\": \"공통·검사실\", \"name\": \"법정감염병 진단검사 통합지침\", \"scope\": \"진단검사의학과 · 보건환경연구원\", \"deadline\": \"상시 기준 (개정 시 갱신)\", \"lab\": \"법정감염병별 검체 채취·보관·운송, 확인진단 표준시험법(배양·PCR·혈청검사) 및 확인기관 의뢰 절차\", \"dept\": \"진단분석국(세균·바이러스·고위험병원체분석과)\", \"guidelineGroup\": \"법정감염병 진단검사 통합지침\", \"guidelineKeyword\": \"진단검사 통합지침\", \"latestTitle\": \"법정감염병 진단검사 통합지침 (KDCA 지침 게시판)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-3\", \"group\": \"공통·검사실\", \"name\": \"의료관련감염 표준예방지침 및 집단발생 대응 안내\", \"scope\": \"전 의료기관 · 감염관리실\", \"deadline\": \"상시 기준 (개정 시 갱신)\", \"lab\": \"손위생·표준주의·전파경로별 주의·환경소독·삽입기구 감염예방 및 다제내성균(CRE·C. auris 등) 집단발생 대응 매뉴얼\", \"dept\": \"의료감염관리과\", \"guidelineGroup\": \"의료관련감염 표준예방지침\", \"guidelineKeyword\": \"표준예방지침\", \"latestTitle\": \"의료관련감염 표준예방지침 · 집단발생 대응 안내\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-4\", \"group\": \"공통·검사실\", \"name\": \"칸디다 오리스(Candida auris) 의료기관 감염관리 권고안\", \"scope\": \"종합병원 · 감염관리실 · 미생물검사실\", \"deadline\": \"수시 개정 (2026년 4급 신설)\", \"lab\": \"선제 감시배양·접촉자 검사·격리해제 기준·환경소독 및 MALDI-TOF/분자동정 확인 권고사항\", \"dept\": \"의료감염관리과\", \"guidelineGroup\": \"칸디다 오리스 의료기관 감염관리 권고안\", \"guidelineKeyword\": \"칸디다 오리스\", \"latestTitle\": \"칸디다 오리스(C. auris) 의료기관 감염관리 권고안 (2026)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-5\", \"group\": \"공통·검사실\", \"name\": \"감염성물질 안전수송 지침\", \"scope\": \"검체 채취·수송·의뢰 기관\", \"deadline\": \"상시 기준\", \"lab\": \"카테고리 A(UN2814)·카테고리 B(UN3373) 3중 포장 기준, 4℃ 냉장·-70℃ 드라이아이스 수송 및 고위험병원체 이동 신고 절차\", \"dept\": \"생물안전평가과\", \"guidelineGroup\": \"감염성물질 안전수송 지침\", \"guidelineKeyword\": \"안전수송\", \"latestTitle\": \"감염성물질 안전수송 지침 (KDCA 지침 게시판)\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-6\", \"group\": \"공통·검사실\", \"name\": \"실험실 생물안전지침 및 고위험병원체 취급·보존 관리지침\", \"scope\": \"BL2/BL3 검사실 · 고위험병원체 취급기관\", \"deadline\": \"상시 기준\", \"lab\": \"고위험병원체(세균·진균 14종, 바이러스·프리온 24종) 분리신고·이동신고·보존관리 및 생물안전작업대(BSC) 운영 기준\", \"dept\": \"생물안전평가과\", \"guidelineGroup\": \"실험실 생물안전지침\", \"guidelineKeyword\": \"생물안전\", \"latestTitle\": \"실험실 생물안전지침 · 고위험병원체 관리 안내\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-7\", \"group\": \"공통·검사실\", \"name\": \"의료기관 항생제 적정사용 관리(ASP) 지침\", \"scope\": \"종합병원 · ASP팀 · 진단검사의학과\", \"deadline\": \"상시 기준\", \"lab\": \"원내 항생제 사용량(DOT/DDD) 모니터링, 누적 항생제 감수성(Antibiogram) 작성·환류 및 선택적 보고(Cascade reporting) 연계\", \"dept\": \"항생제내성관리과\", \"guidelineGroup\": \"항생제 적정사용 관리 지침\", \"guidelineKeyword\": \"항생제 적정사용\", \"latestTitle\": \"의료기관 항생제 적정사용 관리(ASP) 지침\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}, {\"code\": \"공통-8\", \"group\": \"공통·검사실\", \"name\": \"감염병 감시체계(전수·표본감시) 운영 지침\", \"scope\": \"전 의료기관 · 표본감시기관\", \"deadline\": \"매년 개정\", \"lab\": \"전수감시(1~3급) 및 표본감시(4급 24종·호흡기/장관/HAI) 신고 서식, 방역통합정보시스템(eid.kdca.go.kr) 보고 절차\", \"dept\": \"감염병관리과 · 감염병정보과\", \"guidelineGroup\": \"감염병 감시체계 운영 지침\", \"guidelineKeyword\": \"표본감시 운영\", \"latestTitle\": \"의료관련감염병 및 법정감염병 감시 운영 지침 (2026)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/kdca/2861/subview.do\"}]"
REFERENCE_STANDARDS_MASTER_JSON = "{\"updatedNote\": \"2026-11-23 공식 홈페이지 기준 검수 (Phase 2에서 주 1회 자동 개정 감지 연동 예정)\", \"clsi_m\": [{\"code\": \"CLSI M100\", \"series\": \"M\", \"title\": \"Performance Standards for Antimicrobial Susceptibility Testing\", \"titleKo\": \"항균제 감수성 검사(AST) 수행 표준 브레이크포인트·QC 표\", \"edition\": \"36th Ed (M100-Ed36)\", \"publishedDate\": \"2026-01-26\", \"status\": \"2026 최신판\", \"note\": \"매년 1~2월 개정. M02·M07·M11과 연계 적용. Salmonella/Shigella 표 재배치 등 반영.\", \"url\": \"https://clsi.org/shop/standards/m100/\"}, {\"code\": \"CLSI M27M44S\", \"series\": \"M\", \"title\": \"Performance Standards for Antifungal Susceptibility Testing of Yeasts\", \"titleKo\": \"효모균(Candida 등) 항진균제 감수성 검사 브레이크포인트·QC 표\", \"edition\": \"4th Ed (M27M44S-Ed4)\", \"publishedDate\": \"2026-03-04\", \"status\": \"2026 최신판\", \"note\": \"M27(액체배지희석법)·M44(디스크확산법) 연계 효모균 최신 판정표.\", \"url\": \"https://clsi.org/shop/standards/m27m44s/\"}, {\"code\": \"CLSI M38M51S\", \"series\": \"M\", \"title\": \"Performance Standards for Antifungal Susceptibility Testing of Filamentous Fungi\", \"titleKo\": \"사상균(Aspergillus 등) 항진균제 감수성 검사 브레이크포인트·QC 표\", \"edition\": \"4th Ed (M38M51S-Ed4)\", \"publishedDate\": \"2026-03-04\", \"status\": \"2026 최신판\", \"note\": \"M38·M51 연계 사상균 MIC/ECV·QC 최신 판정표.\", \"url\": \"https://clsi.org/shop/standards/m38m51s/\"}, {\"code\": \"CLSI M02\", \"series\": \"M\", \"title\": \"Performance Standards for Antimicrobial Disk Susceptibility Tests\", \"titleKo\": \"항균제 디스크 확산법 표준 시험법\", \"edition\": \"14th Ed (M02-Ed14)\", \"publishedDate\": \"2024-03-19\", \"status\": \"현행 유효\", \"note\": \"호기성 세균 디스크 확산법 표준 절차 (M100과 병행 적용).\", \"url\": \"https://clsi.org/shop/standards/m02/\"}, {\"code\": \"CLSI M07\", \"series\": \"M\", \"title\": \"Methods for Dilution Antimicrobial Susceptibility Tests for Bacteria That Grow Aerobically\", \"titleKo\": \"호기성 세균 액체·한천배지 희석법(MIC) 표준 시험법\", \"edition\": \"12th Ed (M07-Ed12)\", \"publishedDate\": \"2024-03-19\", \"status\": \"현행 유효 (2025.05 정오표)\", \"note\": \"액체배지미량희석법(BMD) 표준 기준. 2025년 5월 정오표(Correction Notice) 확인 필요.\", \"url\": \"https://clsi.org/shop/standards/m07/\"}, {\"code\": \"CLSI M11\", \"series\": \"M\", \"title\": \"Methods for Antimicrobial Susceptibility Testing of Anaerobic Bacteria\", \"titleKo\": \"혐기성 세균 항균제 감수성 검사 표준 시험법\", \"edition\": \"9th Ed (M11-Ed9)\", \"publishedDate\": \"2018-10-16\", \"status\": \"현행 유효 (2025.04 재확인)\", \"note\": \"2025년 4월 Reaffirmed. Bacteroides/Parabacteroides BMD 대안법 포함.\", \"url\": \"https://clsi.org/shop/standards/m11/\"}, {\"code\": \"CLSI M24S\", \"series\": \"M\", \"title\": \"Performance Standards for Susceptibility Testing of Mycobacteria, Nocardia spp., and Other Aerobic Actinomycetes\", \"titleKo\": \"항산균(결핵·NTM)·Nocardia 감수성 검사 판정표\", \"edition\": \"2nd Ed (M24S-Ed2)\", \"publishedDate\": \"2023-02-27\", \"status\": \"현행 유효\", \"note\": \"M24 연계 결핵균·비결핵항산균(NTM)·Nocardia 브레이크포인트 및 QC 표.\", \"url\": \"https://clsi.org/shop/standards/m24s/\"}, {\"code\": \"CLSI M45\", \"series\": \"M\", \"title\": \"Methods for Antimicrobial Dilution and Disk Susceptibility Testing of Infrequently Isolated or Fastidious Bacteria\", \"titleKo\": \"드문 분리균·까다로운 세균 감수성 검사 지침\", \"edition\": \"3rd Ed (M45-Ed3)\", \"publishedDate\": \"2016-08-15\", \"status\": \"현행 유효\", \"note\": \"Aeromonas, Bacillus, Campylobacter, Corynebacterium, Pasteurella, Vibrio 등 특수 균종 AST 기준.\", \"url\": \"https://clsi.org/shop/standards/m45/\"}, {\"code\": \"CLSI M58\", \"series\": \"M\", \"title\": \"Methods for the Identification of Cultured Microorganisms Using MALDI-TOF MS\", \"titleKo\": \"MALDI-TOF MS를 이용한 배양 미생물 동정 지침\", \"edition\": \"1st Ed (M58-Ed1)\", \"publishedDate\": \"2017-04-26\", \"status\": \"현행 유효\", \"note\": \"MALDI-TOF MS 동정 검증, 스코어 해석 및 운영·QC 기준.\", \"url\": \"https://clsi.org/shop/standards/m58/\"}, {\"code\": \"CLSI M64\", \"series\": \"M\", \"title\": \"Implementation of Taxonomy Nomenclature Changes\", \"titleKo\": \"미생물 분류학 명명법 변경의 검사실 적용 지침\", \"edition\": \"1st Ed (M64-Ed1)\", \"publishedDate\": \"2024-12-05\", \"status\": \"현행 유효\", \"note\": \"Candidozyma(Candida) auris 등 개정 균종명 LIS 반영·임상 공지 가이드.\", \"url\": \"https://clsi.org/shop/standards/m64/\"}], \"clsi_mm\": [{\"code\": \"CLSI MM06\", \"series\": \"MM\", \"title\": \"Quantitative Molecular Methods for Infectious Diseases\", \"titleKo\": \"감염병 정량 분자진단 검사 지침 (HIV·HBV·HCV·CMV 등)\", \"edition\": \"3rd Ed (MM06-Ed3)\", \"publishedDate\": \"2026-07-22\", \"status\": \"2026 신규 개정\", \"note\": \"2010년 Ed2 이후 16년 만의 전면 개정(2026.07.22 발간). 바이러스 정량 PCR 검증·정도관리 반영.\", \"url\": \"https://clsi.org/shop/standards/mm06/\"}, {\"code\": \"CLSI MM03\", \"series\": \"MM\", \"title\": \"Molecular Diagnostic Methods for Infectious Diseases\", \"titleKo\": \"감염병 분자진단 검사 표준 지침\", \"edition\": \"현행 최신판\", \"publishedDate\": \"공식 페이지 확인\", \"status\": \"현행 유효\", \"note\": \"감염병 핵산증폭검사(NAAT) 전반의 설계·검증·오염방지·정도관리 핵심 기준.\", \"url\": \"https://clsi.org/shop/standards/mm03/\"}, {\"code\": \"CLSI MM17\", \"series\": \"MM\", \"title\": \"Validation and Verification of Multiplex Nucleic Acid Assays\", \"titleKo\": \"다중(Multiplex) 핵산증폭 패널 검사 성능 검증 지침\", \"edition\": \"2nd Ed (MM17-Ed2)\", \"publishedDate\": \"2018-05-31\", \"status\": \"현행 유효\", \"note\": \"호흡기·위장관·뇌수막염 다중 PCR 패널(FilmArray, QIAstat-Dx 등) 분석적 검증·QC 기준.\", \"url\": \"https://clsi.org/shop/standards/mm17/\"}, {\"code\": \"CLSI MM18\", \"series\": \"MM\", \"title\": \"Interpretive Criteria for Identification of Bacteria and Fungi by Targeted DNA Sequencing\", \"titleKo\": \"표적 DNA 염기서열분석(16S rRNA·ITS 등) 세균·진균 동정 판독 기준\", \"edition\": \"2nd Ed (MM18-Ed2)\", \"publishedDate\": \"2018-07-16\", \"status\": \"현행 유효\", \"note\": \"16S rRNA·ITS·D1/D2 시퀀싱 동정 일치율 판정 및 보고 기준.\", \"url\": \"https://clsi.org/shop/standards/mm18/\"}, {\"code\": \"CLSI MM24\", \"series\": \"MM\", \"title\": \"Molecular Methods for Genotyping and Strain Typing of Infectious Organisms\", \"titleKo\": \"감염성 병원체 유전형 분석 및 균주 형별(WGS·MLST) 지침\", \"edition\": \"1st Ed (MM24-Ed1)\", \"publishedDate\": \"2021-09-30\", \"status\": \"현행 유효\", \"note\": \"원내 유행 역학 추적·cgMLST·유전형 분석 분자검사 표준.\", \"url\": \"https://clsi.org/shop/standards/mm24/\"}], \"eucast\": [{\"code\": \"EUCAST Bacteria Breakpoints\", \"series\": \"EUCAST\", \"title\": \"Clinical Breakpoint Tables for Bacteria (MICs & Zone Diameters)\", \"titleKo\": \"EUCAST 세균 임상 브레이크포인트 표 (무료 원문 PDF/XLS)\", \"edition\": \"v16.1 (Dosages v16.0)\", \"publishedDate\": \"2026-06-24 ~ 2026-12-31\", \"status\": \"2026 최신판\", \"note\": \"2026년 6월 v16.1 개정 반영. S / I(Increased exposure) / R 정의 및 ATU 주의.\", \"url\": \"https://www.eucast.org/bacteria/clinical-breakpoints-and-interpretation/clinical-breakpoint-tables/\"}, {\"code\": \"EUCAST Fungi (AFST)\", \"series\": \"EUCAST\", \"title\": \"Clinical Breakpoints & ECOFFs for Fungi (Yeasts, Moulds, Dermatophytes)\", \"titleKo\": \"EUCAST 진균 항진균제 임상 브레이크포인트 및 ECOFF 표\", \"edition\": \"Breakpoints v12.1 / ECOFFs v6.1\", \"publishedDate\": \"2026-04-10 / 2026-04-18\", \"status\": \"2026 최신판\", \"note\": \"Candida auris 신규 브레이크포인트 해석 지침 및 2026년 Voriconazole/Azole rationale 포함.\", \"url\": \"https://www.eucast.org/fungi-afst/clinical-breakpoints-and-interpretation/clinical-breakpoint-table/\"}, {\"code\": \"EUCAST QC Tables\", \"series\": \"EUCAST\", \"title\": \"Quality Control Criteria for Disk Diffusion & MIC (Aerobic/Anaerobic)\", \"titleKo\": \"EUCAST 내부정도관리(QC) 기준표\", \"edition\": \"v16.0\", \"publishedDate\": \"2026-01-01\", \"status\": \"2026 최신판\", \"note\": \"호기성·혐기성 세균 ATCC/NCTC QC 균주 허용 범위.\", \"url\": \"https://www.eucast.org/bacteria/methodology-and-instructions/disk-diffusion-and-quality-control/\"}, {\"code\": \"EUCAST RAST\", \"series\": \"EUCAST\", \"title\": \"Rapid Antimicrobial Susceptibility Testing (RAST) directly from Positive Blood Cultures\", \"titleKo\": \"양성 혈액배양 직접 신속 디스크확산법(RAST) 지침\", \"edition\": \"Method v7.0 / QC v8.0\", \"publishedDate\": \"2026-01-16\", \"status\": \"2026 최신판\", \"note\": \"양성 혈액배양병에서 4·6·8·16-20시간 판독용 신속 AST 기준.\", \"url\": \"https://www.eucast.org/bacteria/rapid-ast-in-blood-culture/methods-and-quality-control/\"}, {\"code\": \"EUCAST Anaerobic Disk\", \"series\": \"EUCAST\", \"title\": \"EUCAST Disk Diffusion Manual for Rapidly Growing Anaerobic Bacteria\", \"titleKo\": \"신속 발육 혐기성 세균 디스크 확산법 매뉴얼\", \"edition\": \"v3.0 (Aerobic Manual v13.0)\", \"publishedDate\": \"2026-06\", \"status\": \"2026 최신판\", \"note\": \"2026년 6월 혐기성 세균 디스크 확산법 매뉴얼 v3.0 개정 반영.\", \"url\": \"https://www.eucast.org/bacteria/methodology-and-instructions/disk-diffusion-and-quality-control/\"}]}"

ROOT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_JSON_PATH = ROOT_DIR / "docs" / "data" / "dashboard.json"

HTTP = requests.Session()
HTTP.headers.update({"User-Agent": "ClinicalMicroBriefing/8.3.0 (Academic Surveillance Bot)"})


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


def snap_anchor(val: Any, allowed: Tuple[int, ...]) -> int:
    n = safe_int(val, allowed[-1])
    return min(allowed, key=lambda a: (abs(a - n), -a))


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


def evaluate_paper_score_and_exception(
    ai: Dict[str, Any],
    content_type: str,
    has_abstract: bool,
    cutoffs: Dict[str, int],
) -> Dict[str, Any]:
    """
    완전 고정 앵커(35/30/20/15) + 유형별 Cap + Critical Exception(E1~E4, 필수 3조건) 기계적 판정.
    원점수·Cap은 예외가 발생하더라도 임의로 변경하지 않고 투명하게 보존함.
    """
    s1 = snap_anchor(ai.get("score_practice"), ANCHORS_PRACTICE)
    s2 = snap_anchor(ai.get("score_evidence"), ANCHORS_EVIDENCE)
    s3 = snap_anchor(ai.get("score_korea"), ANCHORS_KOREA)
    s4 = snap_anchor(ai.get("score_innovation"), ANCHORS_NOVELTY)
    raw_score = s1 + s2 + s3 + s4

    study_type = safe_str(ai.get("study_type")) or "uncapped"
    if study_type not in STUDY_TYPE_CAPS:
        study_type = "uncapped"

    cap_val, cap_label = STUDY_TYPE_CAPS[study_type]
    # 공식 학회 가이드라인은 유형 Cap 예외(uncapped), 초록 없는 일반 논문은 65점 상한
    if content_type == "guideline":
        cap_val = None
        study_type = "uncapped"
        cap_label = STUDY_TYPE_CAPS["uncapped"][1]
    elif not has_abstract:
        if cap_val is None or cap_val > 65:
            cap_val = 65
            cap_label = "초록 원문 부재(확인 필요)"

    final_score = min(raw_score, cap_val) if cap_val is not None else raw_score
    cap_applied = cap_val is not None and raw_score > cap_val
    tier = score_to_tier(final_score, cutoffs)

    breakdown = f"실무:{s1}/35 | 근거:{s2}/30 | 국내:{s3}/20 | 신규:{s4}/15"
    if cap_applied:
        breakdown += f" | Cap:{raw_score}→{final_score}({cap_label})"
    elif cap_val is not None:
        breakdown += f" | 유형:{cap_label}(Cap {cap_val})"

    # Critical Exception (필수 3조건 모두 True + E1~E4 + 근거 사유 존재)
    exc_type = safe_str(ai.get("exception_type")).upper()
    cond1 = bool(ai.get("cond_lab_direct"))
    cond2 = bool(ai.get("cond_high_impact"))
    cond3 = bool(ai.get("cond_reliable_evidence"))
    exc_reason = safe_str(ai.get("exception_reason"))

    is_ce = exc_type in EXCEPTION_LABELS and cond1 and cond2 and cond3 and len(exc_reason) >= 5
    critical_exception = exc_type if is_ce else ""
    if not is_ce:
        exc_reason = ""

    ai["study_type"] = study_type
    ai["cap_value"] = cap_val
    ai["cap_label"] = cap_label if cap_applied else ""
    ai["raw_score"] = raw_score
    ai["importance_score"] = final_score
    ai["importance"] = tier
    ai["score_breakdown"] = breakdown
    ai["critical_exception"] = critical_exception
    ai["exception_reason"] = exc_reason
    ai["needs_pass2"] = (tier == "Must-Read") or bool(critical_exception)
    return ai


def evaluate_outbreak_score_and_exception(
    ai: Dict[str, Any],
    cutoffs: Dict[str, int],
) -> Dict[str, Any]:
    s1 = snap_anchor(ai.get("score_domestic"), ANCHORS_PRACTICE)
    s2 = snap_anchor(ai.get("score_severity"), ANCHORS_EVIDENCE)
    s3 = snap_anchor(ai.get("score_action"), ANCHORS_KOREA)
    s4 = snap_anchor(ai.get("score_novelty"), ANCHORS_NOVELTY)
    total = s1 + s2 + s3 + s4
    tier = score_to_tier(total, cutoffs)

    exc_type = safe_str(ai.get("exception_type")).upper()
    cond1 = bool(ai.get("cond_lab_direct"))
    cond2 = bool(ai.get("cond_high_impact"))
    cond3 = bool(ai.get("cond_reliable_evidence"))
    exc_reason = safe_str(ai.get("exception_reason"))
    is_ce = exc_type == "E4" and cond1 and cond2 and cond3 and len(exc_reason) >= 5

    ai["importance_score"] = total
    ai["importance"] = tier
    ai["score_breakdown"] = f"국내영향:{s1}/35 | 위험도:{s2}/30 | 대응시급:{s3}/20 | 신규성:{s4}/15"
    ai["critical_exception"] = "E4" if is_ce else ""
    ai["exception_reason"] = exc_reason if is_ce else ""
    return ai


def resolve_effective_paper_start(cfg: Dict[str, str], allow_h1: bool) -> Tuple[str, bool]:
    """상반기 소급 가드: allow_h1=False(기본값)이면 논문 처리 범위를 2026-07-01 이후로 강제 보호"""
    cfg_start = cfg.get("collect_start_date", "").strip() or H2_GUARD_START_DATE
    cfg_allow_h1 = cfg.get("allow_h1_backfill", "").strip().upper() == "TRUE"
    h1_unlocked = allow_h1 or cfg_allow_h1
    if not h1_unlocked and cfg_start < H2_GUARD_START_DATE:
        return H2_GUARD_START_DATE, True
    return cfg_start, False


# =====================================================================
# 3-Stage Cross-Journal Co-Publication Duplicate Prevention (v8.2.1)
# =====================================================================
TITLE_DEDUP_STOPWORDS = {
    "a", "an", "the", "of", "in", "on", "for", "and", "or", "to", "with", "from", "by", "at", "as",
    "study", "clinical", "patients", "adults", "children", "using", "among", "between", "against",
    "during", "after", "associated", "infection", "infections", "disease", "diseases", "update",
}


def normalize_doi(doi: Any) -> str:
    s = safe_str(doi).lower()
    s = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", s)
    return s.strip().rstrip(".")


def _title_tokens(title: str) -> Set[str]:
    words = re.findall(r"[a-z0-9]{3,}", str(title or "").lower())
    return {w for w in words if w not in TITLE_DEDUP_STOPWORDS}


def _abstract_words(text: str) -> List[str]:
    clean = re.sub(r"\b(?:background|methods?|results?|conclusions?|objectives?|aims?|introduction|summary)\s*:", " ", str(text or "").lower())
    return re.findall(r"[a-z0-9]{2,}", clean)


def compute_abstract_similarity(abs_a: str, abs_b: str) -> Tuple[float, float]:
    """두 영문 초록 간 (단어 Jaccard, 3-gram 슁글 Jaccard) 반환"""
    wa = _abstract_words(abs_a)
    wb = _abstract_words(abs_b)
    if len(wa) < 25 or len(wb) < 25:
        return 0.0, 0.0
    set_a, set_b = set(wa), set(wb)
    word_jacc = len(set_a & set_b) / len(set_a | set_b) if (set_a | set_b) else 0.0

    sh_a = set(zip(wa, wa[1:], wa[2:]))
    sh_b = set(zip(wb, wb[1:], wb[2:]))
    shingle_jacc = len(sh_a & sh_b) / len(sh_a | sh_b) if (sh_a | sh_b) else 0.0
    return word_jacc, shingle_jacc


def is_copublication_abstract_match(abs_a: str, abs_b: str) -> bool:
    """서로 다른 저널 간 동시 게재 판정 (단어 유사도 >= 0.70 또는 3-gram 유사도 >= 0.55)"""
    word_jacc, shingle_jacc = compute_abstract_similarity(abs_a, abs_b)
    return word_jacc >= 0.70 or shingle_jacc >= 0.55


def _title_cross_suspect(title_a: str, title_b: str) -> bool:
    """서로 다른 저널 간 초록 비교 대상(동시 게재 의심 쌍)을 포착하는 1차 필터"""
    ta = _title_tokens(title_a)
    tb = _title_tokens(title_b)
    if len(ta) < 3 or len(tb) < 3:
        return False
    inter = len(ta & tb)
    if inter < 3:
        return False
    jacc = inter / len(ta | tb)
    overlap = inter / min(len(ta), len(tb))
    return jacc >= 0.30 or (overlap >= 0.50 and inter >= 4)


def detect_cross_journal_duplicates(
    candidates: List[Dict[str, Any]],
    existing_records: List[Dict[str, Any]],
    ncbi_api_key: str,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    신규 수집 후보군(candidates)에서 이종 저널 동시 게재 중복을 탐지하여
    (정상 요약 대상 리스트, 0원 duplicate 마킹 대상 리스트)로 분리 반환.
    - 가이드라인(guideline)이 일반 저널(paper)보다 우선 보존되도록 정렬 후 판정.
    """
    active_existing = [
        r for r in existing_records
        if r.get("contentType") in ("paper", "guideline") and r.get("uid")
    ]
    existing_dois = {normalize_doi(r.get("doi")): r["uid"] for r in existing_records if normalize_doi(r.get("doi"))}
    existing_pmids = {clean_pmid(r.get("pmid")): r["uid"] for r in existing_records if clean_pmid(r.get("pmid"))}

    suspect_existing_by_cand: Dict[str, List[Dict[str, Any]]] = {}
    suspect_pmids: Set[str] = set()
    suspect_dois: Set[str] = set()

    for c in candidates:
        if c.get("contentType") not in ("paper", "guideline"):
            continue
        if len(safe_str(c.get("rawText"))) < 120:
            continue
        c_src = safe_str(c.get("source")).upper()
        c_date = parse_flexible_date(c.get("pubDate"))
        matches = []
        for ex in active_existing:
            if safe_str(ex.get("source")).upper() == c_src:
                continue
            ex_date = parse_flexible_date(ex.get("pubDate"))
            if c_date and ex_date and abs((c_date - ex_date).days) > 365:
                continue
            if _title_cross_suspect(c.get("title", ""), ex.get("title", "")):
                matches.append(ex)
                if ex.get("pmid"):
                    suspect_pmids.add(ex["pmid"])
                elif ex.get("doi"):
                    suspect_dois.add(ex["doi"])
        if matches:
            suspect_existing_by_cand[c["uid"]] = matches

    ex_abs_by_pmid: Dict[str, str] = {}
    if suspect_pmids:
        try:
            print(f"  [중복 검증] 이종 저널 의심 기존 항목 {len(suspect_pmids)}건 원문 초록 대조 조회...")
            for art in fetch_pubmed_articles(sorted(suspect_pmids), ncbi_api_key):
                ex_abs_by_pmid[str(art["pmid"])] = art.get("abstractText", "")
        except Exception as e:
            print(f"  [중복 검증 경고] PubMed 대조 초록 조회 실패(무시 후 진행): {e}")

    ex_abs_by_doi: Dict[str, str] = {}
    for d in sorted(suspect_dois):
        ex_abs_by_doi[d] = fetch_crossref_abstract(d)
        time.sleep(0.15)

    ordered = sorted(
        candidates,
        key=lambda x: (0 if x.get("contentType") == "guideline" else (1 if x.get("contentType") == "outbreak" else 2)),
    )

    clean_list: List[Dict[str, Any]] = []
    dup_list: List[Dict[str, Any]] = []
    accepted_deep: List[Dict[str, Any]] = []
    accepted_dois: Dict[str, str] = dict(existing_dois)
    accepted_pmids: Dict[str, str] = dict(existing_pmids)

    for c in ordered:
        if c.get("contentType") == "outbreak":
            clean_list.append(c)
            continue

        c_pmid = clean_pmid(c.get("pmid"))
        c_doi = normalize_doi(c.get("doi"))
        dup_of = ""

        if c_pmid and c_pmid in accepted_pmids:
            dup_of = accepted_pmids[c_pmid]
        elif c_doi and c_doi in accepted_dois:
            dup_of = accepted_dois[c_doi]

        c_abs = safe_str(c.get("rawText"))
        if not dup_of and c_abs and c["uid"] in suspect_existing_by_cand:
            for ex in suspect_existing_by_cand[c["uid"]]:
                ex_abs = ex_abs_by_pmid.get(ex.get("pmid", ""), "") or ex_abs_by_doi.get(ex.get("doi", ""), "")
                if ex_abs and is_copublication_abstract_match(c_abs, ex_abs):
                    dup_of = ex["uid"]
                    break

        if not dup_of and len(c_abs) >= 120:
            c_src = safe_str(c.get("source")).upper()
            for prev in accepted_deep:
                if safe_str(prev.get("source")).upper() == c_src:
                    continue
                prev_abs = safe_str(prev.get("rawText"))
                if len(prev_abs) >= 120 and is_copublication_abstract_match(c_abs, prev_abs):
                    dup_of = prev["uid"]
                    break

        if dup_of:
            c_copy = dict(c)
            c_copy["duplicateOf"] = dup_of
            dup_list.append(c_copy)
            print(f"  [동시 게재 자동 차단 (0원)] {c['uid']} ({c['source']}) == {dup_of}")
        else:
            clean_list.append(c)
            accepted_deep.append(c)
            if c_pmid:
                accepted_pmids[c_pmid] = c["uid"]
            if c_doi:
                accepted_dois[c_doi] = c["uid"]

    return clean_list, dup_list


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
    """items 시트 18~23열(R~W: summary_meta ~ exception_reason) 자동 확보"""
    header = ws.row_values(1)
    if ws.col_count < 23:
        ws.resize(cols=23)
    col_defs = [
        (18, "R1", "summary_meta"),
        (19, "S1", "importance_reason"),
        (20, "T1", "importance_score"),
        (21, "U1", "score_breakdown"),
        (22, "V1", "critical_exception"),
        (23, "W1", "exception_reason"),
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


def get_gemini_models(cfg: Dict[str, str]) -> Tuple[str, str, str, str, str]:
    lite = cfg.get("gemini_model", "").strip() or "gemini-3.5-flash-lite"
    deep = cfg.get("gemini_model_deep", "").strip() or "gemini-4-hydrogen"
    fallback = cfg.get("gemini_model_fallback", "").strip() or "gemini-3.6-flash"
    thinking_p1 = cfg.get("gemini_thinking_level_deep", "").strip().lower() or "low"
    thinking_p2 = cfg.get("gemini_thinking_level_pass2", "").strip().lower() or "high"
    return lite, deep, fallback, thinking_p1, thinking_p2


def append_run_log(ss, run_id: str, ts: str, source: str, status: str, fetched: int, added: int, dup: int, note: str):
    ws = get_or_create_worksheet(
        ss,
        SHEET_LOGS,
        ["run_id", "timestamp", "source", "status", "fetched_count", "new_count", "duplicate_count", "error_message"],
    )
    ws.append_row([run_id, ts, source, status, fetched, added, dup, note or ""], value_input_option="RAW")


# =====================================================================
# Gemini API (2-Pass Differential Summary + 503 Auto-Fallback)
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


def _call_single_gemini_json(
    prompt: str,
    response_schema: Dict[str, Any],
    model_name: str,
    api_key: str,
    temperature: float,
    thinking_level: str,
    timeout_read: int = 105,
    max_attempts: int = 3,
) -> Dict[str, Any]:
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
            res = HTTP.post(url, headers={"x-goog-api-key": api_key}, json=payload, timeout=(15, timeout_read))
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
                return json.loads(cleaned)

            last_err = f"HTTP {res.status_code}: {res.text[:220]}"
            if res.status_code in (429, 500, 502, 503, 504) and attempt < max_attempts:
                wait_sec = 6 * attempt
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
    timeout_read: int = 105,
) -> Tuple[Dict[str, Any], str]:
    try:
        parsed = _call_single_gemini_json(
            prompt, response_schema, primary_model, api_key, temperature, thinking_level, timeout_read, max_attempts=3
        )
        return parsed, primary_model
    except Exception as primary_err:
        if fallback_model and fallback_model != primary_model:
            print(f"  [503/지연 자동 우회] {primary_model} 실패({str(primary_err)[:80]}) -> {fallback_model} 전환 실행")
            fb_thinking = thinking_level if re.match(r"^gemini-[34]", fallback_model, re.IGNORECASE) else ""
            if fb_thinking == "high":
                fb_thinking = "medium"
            parsed = _call_single_gemini_json(
                prompt, response_schema, fallback_model, api_key, temperature, fb_thinking, timeout_read, max_attempts=3
            )
            return parsed, fallback_model
        raise


def format_clinical_note(
    raw_note: str,
    tier: str,
    total_score: int,
    critical_exception: str,
    pass2_applied: bool,
) -> str:
    clean = str(raw_note or "").strip()
    clean = re.sub(r"^(\[?\s*AI\s*(임상\s*)?(참고\s*)?해석[^\]]*\]?\s*[:：\-]?\s*)", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^AI의\s*(임상\s*)?(참고\s*)?해석에\s*따르면[,]?\s*", "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^본\s*원저는\s*", "", clean, flags=re.IGNORECASE)
    prefix_parts = ["AI 해석"]
    if critical_exception:
        prefix_parts.append(f"예외필독 {critical_exception}")
    elif tier == "Must-Read":
        prefix_parts.append(f"필독 {total_score}점")
    if pass2_applied:
        prefix_parts.append("2-Pass 심층")
    return f"[{' · '.join(prefix_parts)}] {clean}"


def summarize_deep_pass1_batch(
    chunk: List[Dict[str, Any]],
    primary_model: str,
    fallback_model: str,
    api_key: str,
    thinking_level_p1: str,
    cutoffs: Dict[str, int],
) -> Tuple[Dict[str, Dict[str, Any]], str]:
    """Pass 1: 전수 채점(35/30/20/15 완전 고정 앵커) + 유형별 Cap 분류 + Critical Exception 판정 + 간략 요약"""
    blocks = []
    has_abs_map: Dict[str, bool] = {}
    ctype_map: Dict[str, str] = {}
    for idx, it in enumerate(chunk, 1):
        raw = str(it.get("rawText") or "").strip()
        has_abs = len(raw) >= 40
        has_abs_map[it["uid"]] = has_abs
        ctype_map[it["uid"]] = it["contentType"]
        abs_display = raw[:3500] if has_abs else "(초록 원문 없음 — 제목만 제공됨)"
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
            f'아래 {len(chunk)}편의 논문 또는 가이드라인에 대해 [Pass 1: 전수 정량 채점 + 예외 판정 + 간략 요약]을 수행하고 JSON({{ "items": [...] }})으로만 응답하십시오.',
            "",
            "[절대 준수 원칙]",
            '1. 환각·과장 금지: 초록 원문에 없는 표본 수(n), 민감도/특이도, 균주, 결론을 절대 지어내지 마십시오. 초록 원문이 없으면 fact_summary_ko 첫 문장에 "초록 정보 부족으로 원문 확인 필요."를 명시하십시오.',
            "2. [100점 완전 고정 앵커 채점 — 각 축에서 제시된 4개 숫자 문자열 중 정확히 1개만 선택, 미세조정 금지]:",
            "   ※ 옵션 B 적용: CID 등 주요 내성균(CRE·MRSA 등) 및 중증 감염병의 표준 치료·예방을 바꾸는 대규모 RCT는 포괄적으로 높게 평가하여 Must-Read 진입을 허용합니다.",
            '   - score_practice (실무·검사실 변경력 35점, 선택지: "35" | "26" | "16" | "6"):',
            "     · 35: 공식 가이드라인 제·개정, CLSI/EUCAST breakpoint 변경 근거, AST 중대 오류(VME/ME) 주의보, 주요 내성균·중증감염 표준 치료를 바꾸는 대규모 RCT",
            "     · 26: 검사실 워크플로우·TAT·분자진단·MALDI-TOF·감염관리 프로토콜에 실질적 개선을 제시",
            "     · 16: 간접적 임상·검사실 참고 가치 / 6: 실무 변화 없음, 단순 리뷰·정책·기초",
            '   - score_evidence (근거 강도·연구 설계 30점, 선택지: "30" | "22" | "13" | "5"):',
            "     · 30: 공식 학회 가이드라인, 대규모 다기관 전향 연구(n>=300), 대규모 RCT, 체계적 문헌고찰·메타분석",
            "     · 22: 중규모 원저(n=150~299) 또는 탄탄한 다기관 후향 코호트",
            "     · 13: 단일기관 소규모(n<150) 평가, 후향적 기술연구, 서술형 리뷰 / 5: 증례보고, Letter/Editorial, 순수 기초·동물실험",
            '   - score_korea (국내 대학병원 적용성 20점, 선택지: "20" | "14" | "8" | "3"):',
            "     · 20: CRE·카바페네마제, MRSA/VRE, C. auris, 혈액배양 신속진단/AST, 결핵/NTM, 호흡기·CSF 분자진단, MALDI-TOF 등 국내 대학병원 직결",
            "     · 14: 국내 임상 감염·검사실에서 일반 참고 가능",
            "     · 8: 국내 발생이 드물거나 해외 특정 지역 중심 / 3: 국내 극히 희귀한 열대풍토병·저소득국 한정 정책·수의/비인체",
            '   - score_innovation (학술적 신규성·파급력 15점, 선택지: "15" | "11" | "6" | "2"):',
            "     · 15: 게임체인저급 진단 혁신, 신종 내성 기전 최초 규명, 놓치면 치명적인 진단·치료 함정",
            "     · 11: 의미 있는 기술·임상 근거 확장 / 6: 기존 장비·방법의 점진적 비교 / 2: 반복적 지역 감시·단순 보고",
            "3. [논문 유형 분류 (study_type) — 초록 기준 정확히 1개 선택, 코드가 유형별 Cap 자동 적용]:",
            '   - "uncapped": 공식 학회 지침(IDSA/ESCMID 등), 다기관 평가, 대규모 RCT, 체계적 문헌고찰·메타분석, 국가·국제 규모 감시/표준화 연구',
            '   - "single_center_small_or_retro": 단일기관 소규모 평가 또는 단일기관 후향적 기술연구 (Cap 82)',
            '   - "narrative_review": 서술형 종설·일반 리뷰·미니리뷰·입장문(비체계적) (Cap 78)',
            '   - "case_letter_editorial": 증례보고(Case Report)·Letter·Comment·Editorial (Cap 69)',
            '   - "basic_preclinical": 순수 기초연구·in vitro·동물실험 (Cap 65)',
            "4. [Critical Exception (중요 논문 예외 선정 — 엄격 판정, 전체의 극소수만 해당)]:",
            "   점수·Cap과 별도로 아래 4유형 중 하나에 해당하면서 [필수 3조건]을 모두 충족할 때만 예외로 지정하십시오.",
            '   - exception_type: "NONE" | "E1" | "E2" | "E3" | "E4"',
            "     · E1: 공식 기준·지침 변경 (CLSI/EUCAST breakpoint 개정, 주요 검사 지침·보고 기준 변경)",
            "     · E2: 중대한 검사 오류·한계 (기존 상용 검사법의 중대한 위음성·위양성 위험, 특정 변이·내성 기전 검출 실패)",
            "     · E3: 새로운 고위험 병원체·내성 기전 (검사실 검출·동정·감수성검사 전략에 영향을 주는 신규 발견)",
            "     · E4: 긴급 감염병 대응 (WHO·CDC·KDCA 등의 긴급 경보 및 검사 전략 변경 권고)",
            "   - 필수 3조건 (각각 true/false 엄격 판정):",
            "     · cond_lab_direct: 검사실 업무와 직접적인 관련성이 있는가?",
            "     · cond_high_impact: 검사 결과 해석·환자 진료·검사실 대응에 중대한 영향을 미치는가?",
            "     · cond_reliable_evidence: 공식 기관 권고 또는 신뢰할 만한 연구 근거가 있는가?",
            "   ※ 단순 증례보고, 일반 소규모 예비평가, 임상 적용 가능성이 불분명한 초기 연구는 반드시 NONE으로 판정하십시오.",
            "5. Pass 1 간략 요약 규칙:",
            "   - fact_summary_ko: 원문 핵심 사실(대상·설계·주요 수치·결론)을 객관적으로 2~3문장 압축.",
            '   - ai_clinical_note: 금지 상용구("참고가 될 것으로 사료됩니다", "중요성을 시사합니다" 등) 없이 대학병원 검사실 관점의 실무 포인트와 한계 각 1문장.',
            "6. topic_category:",
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
                        "study_type": {
                            "type": "STRING",
                            "enum": [
                                "uncapped",
                                "single_center_small_or_retro",
                                "narrative_review",
                                "case_letter_editorial",
                                "basic_preclinical",
                            ],
                        },
                        "score_practice": {"type": "STRING", "enum": ["35", "26", "16", "6"]},
                        "score_evidence": {"type": "STRING", "enum": ["30", "22", "13", "5"]},
                        "score_korea": {"type": "STRING", "enum": ["20", "14", "8", "3"]},
                        "score_innovation": {"type": "STRING", "enum": ["15", "11", "6", "2"]},
                        "exception_type": {"type": "STRING", "enum": ["NONE", "E1", "E2", "E3", "E4"]},
                        "cond_lab_direct": {"type": "BOOLEAN"},
                        "cond_high_impact": {"type": "BOOLEAN"},
                        "cond_reliable_evidence": {"type": "BOOLEAN"},
                        "exception_reason": {"type": "STRING"},
                        "importance_reason": {"type": "STRING"},
                        "fact_summary_ko": {"type": "STRING"},
                        "ai_clinical_note": {"type": "STRING"},
                    },
                    "required": [
                        "uid",
                        "title_ko",
                        "topic_category",
                        "study_type",
                        "score_practice",
                        "score_evidence",
                        "score_korea",
                        "score_innovation",
                        "exception_type",
                        "cond_lab_direct",
                        "cond_high_impact",
                        "cond_reliable_evidence",
                        "exception_reason",
                        "importance_reason",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }

    parsed, used_model = call_gemini_with_fallback(
        prompt, schema, primary_model, fallback_model, api_key, 0.1, thinking_level_p1, timeout_read=105
    )

    out: Dict[str, Dict[str, Any]] = {}
    for item in parsed.get("items", []):
        if isinstance(item, dict) and item.get("uid"):
            uid = str(item["uid"])
            out[uid] = evaluate_paper_score_and_exception(
                item, ctype_map.get(uid, "paper"), has_abs_map.get(uid, False), cutoffs
            )
    return out, used_model


def run_pass2_deep_summary(
    item: Dict[str, Any],
    p1_ai: Dict[str, Any],
    primary_model: str,
    fallback_model: str,
    api_key: str,
    thinking_level_p2: str = "high",
) -> Tuple[Dict[str, Any], str]:
    """Pass 2: Must-Read(85+) 또는 Critical Exception(E1~E4) 확정 논문 전용 심층 요약 (thinking: high)"""
    raw = str(item.get("rawText") or "").strip()
    abs_display = raw[:4500] if len(raw) >= 40 else "(초록 원문 없음 — 제목만 제공됨)"
    ce = p1_ai.get("critical_exception", "")
    ce_desc = f"{EXCEPTION_LABELS.get(ce, ce)} — {p1_ai.get('exception_reason', '')}" if ce else "해당없음"

    prompt = "\n".join(
        [
            "당신은 대한민국 대학병원 진단검사의학과(임상미생물학·분자진단·항균제 내성) 전문의를 위한 수석 학술 분석 AI입니다.",
            "아래 논문은 1차 정량 평가에서 [필독(Must-Read, 85점 이상)] 또는 [Critical Exception]으로 최종 확정된 핵심 논문입니다.",
            "초록 원문에 철저히 근거하여 [Pass 2: 심층 임상·검사실 브리핑]을 JSON으로 작성하십시오.",
            "",
            f"출처: {item['source']} / 유형: {item['contentType']} / 발행일: {item.get('pubDate', '')}",
            f"원제: {item['title']}",
            f"1차 채점 결과: {p1_ai.get('importance')} ({p1_ai.get('importance_score')}점 · {p1_ai.get('score_breakdown')})",
            f"Critical Exception 여부: {ce_desc}",
            f"초록 원문:\n{abs_display}",
            "",
            "[작성 원칙 — 엄격 준수]",
            "1. 환각·추측 절대 금지: 초록에 명시된 표본 수(n), 대조법, 민감도/특이도/일치도/MIC/OR/p값 등만 인용하십시오. 초록에 없는 수치는 절대 생성하지 마십시오.",
            "2. fact_summary_ko (구조화된 원문 사실 심층 요약, 줄바꿈(•) 3~4개 항목):",
            "   • [배경·설계] 연구 목적, 기관 규모·설계, 검체·균주 수(n) 및 대조법",
            "   • [핵심 결과] 초록에 제시된 주요 정량 성능·임상 결과 수치",
            "   • [결론·의의] 저자의 핵심 결론 및 기존 기준/방법 대비 차별점",
            "3. ai_clinical_note (대학병원 진단검사의학과 전문의 심층 실무 가이드, 3~4문장):",
            '   - 뻔한 상용구("참고가 될 것으로 사료됩니다", "중요성을 시사합니다" 등) 절대 금지.',
            "   - (a) 검사실 실무 적용 포인트(분자진단·배양·MALDI-TOF·AST 판독/보고·정도관리 QC·감염관리 IC 중 해당 영역의 구체적 액션)",
            "   - (b) 현장 도입·판독 시 검증해야 할 연구 한계 및 주의 함정(위양성/위음성, 균종 편중, 대조법 한계, 국내 역학 차이 등)",
            "   - (c) Critical Exception 지정 논문인 경우, 검사실이 당장 점검해야 할 핵심 사항 1문장 포함.",
        ]
    )

    schema = {
        "type": "OBJECT",
        "properties": {
            "title_ko": {"type": "STRING"},
            "fact_summary_ko": {"type": "STRING"},
            "ai_clinical_note": {"type": "STRING"},
        },
        "required": ["title_ko", "fact_summary_ko", "ai_clinical_note"],
    }

    parsed, used_model = call_gemini_with_fallback(
        prompt, schema, primary_model, fallback_model, api_key, 0.1, thinking_level_p2, timeout_read=135
    )
    return parsed, used_model


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
            f'아래 {len(chunk)}건의 국내외 감염병 동향·경보(KDCA/WHO)를 분석해 JSON({{ "items": [...] }})으로 반환하십시오.',
            "",
            "[준수 원칙]",
            '1. 원문에 없는 수치·국가·방역 조치를 절대 생성하지 마십시오. "본 원저는" 같은 논문 표현을 쓰지 마십시오.',
            "2. fact_summary_ko:",
            "   - KDCA 표본감시 주간소식지: 인플루엔자(ILI)·코로나19 입원·호흡기 바이러스 상위·장관감염증 핵심 수치를 줄바꿈(•) 4줄 이내로 압축.",
            "   - KDCA 경보 / WHO DONs: 발생 국가·지역, 확진/사망 규모(치명률), 실험실 확진 검사법(PCR·항체·mNGS), 방역 권고를 2~3문장으로 요약.",
            "3. ai_clinical_note (1~2문장): 뻔한 상용구를 피하고 국내 대학병원 진단검사실·감염관리실의 구체적 대응 포인트 작성.",
            "4. [100점 고정 앵커 채점 (35/30/20/15) — 각 선택지 중 정확히 1개만 선택]:",
            '   - score_domestic (국내 유입·발생 직접성 35점, "35"|"26"|"16"|"6"): 국내 발생·전국 유행주의보·국내 CRE 지침=35 / 인접국·주요 여행지 유입 우려=26 / 원거리 주요국=16 / 국지적 원거리=6',
            '   - score_severity (위험도·급증 강도 30점, "30"|"22"|"13"|"5"): 1급 감염병·콜레라·고치명률 출혈열·중증 급증=30 / 중등도 유행 증가=22 / 완만 추세·정례 보고=13 / 경미=5',
            '   - score_action (검사실·감염관리 즉각 대응 필요성 20점, "20"|"14"|"8"|"3"): 즉시 감별검사·선제격리·지침 변경=20 / 주의 모니터링=14 / 참고=8 / 영향 미미=3',
            '   - score_novelty (신규성·최초 발생 여부 15점, "15"|"11"|"6"|"2"): 최초 인체감염·10년 만의 국내 발생·신규 유행주의보=15 / 신규 국면=11 / 동일 해외 유행 반복 업데이트(예: 콩고 에볼라 주간 집계)·정례 연보·평시 표본감시=2',
            '5. Critical Exception E4 판정: "NONE" | "E4" — 국내 검사 전략 변경·긴급 대응이 즉각 필요한 공식 경보이고 필수 3조건(검사실 직결, 중대 영향, 공식 근거)을 모두 충족할 때만 "E4".',
            '6. topic_category: "표본감시 주간요약(호흡기·장관)" | "의료관련감염병·내성(HAI/CRE)" | "국내 긴급경보(KDCA)" | "국제 감염병 경보(WHO)"',
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
                        "score_domestic": {"type": "STRING", "enum": ["35", "26", "16", "6"]},
                        "score_severity": {"type": "STRING", "enum": ["30", "22", "13", "5"]},
                        "score_action": {"type": "STRING", "enum": ["20", "14", "8", "3"]},
                        "score_novelty": {"type": "STRING", "enum": ["15", "11", "6", "2"]},
                        "exception_type": {"type": "STRING", "enum": ["NONE", "E4"]},
                        "cond_lab_direct": {"type": "BOOLEAN"},
                        "cond_high_impact": {"type": "BOOLEAN"},
                        "cond_reliable_evidence": {"type": "BOOLEAN"},
                        "exception_reason": {"type": "STRING"},
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
                        "exception_type",
                        "cond_lab_direct",
                        "cond_high_impact",
                        "cond_reliable_evidence",
                        "exception_reason",
                        "importance_reason",
                        "fact_summary_ko",
                        "ai_clinical_note",
                    ],
                },
            }
        },
        "required": ["items"],
    }

    parsed, used_model = call_gemini_with_fallback(
        prompt, schema, lite_model, fallback_model, api_key, 0.1, "", timeout_read=90
    )

    out: Dict[str, Dict[str, Any]] = {}
    for item in parsed.get("items", []):
        if isinstance(item, dict) and item.get("uid"):
            out[str(item["uid"])] = evaluate_outbreak_score_and_exception(item, cutoffs)
    return out, used_model


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
    years = [str(now_year - 3), str(now_year - 2), str(now_year - 1), str(now_year)]
    all_records = []

    for yr in years:
        for page_no in range(1, 5):
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
    return len(body), f"{now_year - 3}~{now_year}년 전수감시 {len(body)}건 갱신 (3년 동월 기저선 확보, {now_year}년 {cnt_cur}건 포함)"


def fetch_kdca_sentinel_weekly(start_date_str: str, existing_uids: Optional[Set[str]] = None) -> List[Dict[str, Any]]:
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 7, 1)
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
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 7, 1)
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
    cutoff = parse_flexible_date(start_date_str) or datetime.date(2026, 7, 1)
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
    from_date = (parse_flexible_date(start_date_str) or datetime.date(2026, 7, 1)).strftime("%Y-%m-%d")
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
# Stratified 20-Item Pilot Selector (2026-07-01+ Papers & Guidelines)
# =====================================================================
def select_stratified_pilot_20(records: List[Dict[str, Any]], target_n: int = 20) -> List[Dict[str, Any]]:
    """
    2026-07-01 이후 논문·가이드라인에서 저널·유형(가이드라인/대규모평가/리뷰/소규모/증례 등)별
    균형 잡힌 20편을 결정론적으로 층화 추출함.
    """
    h2_pool = [
        r for r in records
        if r.get("contentType") in ("paper", "guideline") and safe_str(r.get("pubDate")) >= H2_GUARD_START_DATE
    ]
    if len(h2_pool) <= target_n:
        return h2_pool

    def classify_stratum_tag(r: Dict[str, Any]) -> str:
        t = (r.get("title") or "").lower()
        if r.get("contentType") == "guideline" or "guideline" in t or "consensus" in t:
            return "guideline"
        if re.search(r"\b(review|narrative|overview|update on|primer|myths|how do i|how to)\b", t):
            return "review"
        if re.search(r"\b(case|mastoiditis|infant|first evidence|outbreak in a|letter|comment)\b", t):
            return "edge_case"
        if re.search(r"\b(randomized|trial|multicenter|meta-analysis|nationwide|breakpoint|carbapenemase|clsi|eucast)\b", t):
            return "high_impact"
        return "standard"

    def bucket_of(r: Dict[str, Any]) -> str:
        if r.get("contentType") == "guideline":
            return "GUIDELINE"
        src = safe_str(r.get("source")).upper()
        if src in ("JCM", "CMI", "CID", "ACM"):
            return src
        return "OTHER"

    quotas = [("GUIDELINE", 2), ("JCM", 6), ("CMI", 5), ("CID", 5), ("ACM", 2)]
    by_bucket: Dict[str, List[Dict[str, Any]]] = {}
    for r in sorted(h2_pool, key=lambda x: (x.get("pubDate", ""), x.get("uid", "")), reverse=True):
        by_bucket.setdefault(bucket_of(r), []).append(r)

    selected: List[Dict[str, Any]] = []
    selected_uids: Set[str] = set()

    def pick_diverse(pool: List[Dict[str, Any]], k: int) -> List[Dict[str, Any]]:
        if not pool or k <= 0:
            return []
        order = ["high_impact", "review", "edge_case", "standard", "guideline"]
        tag_groups: Dict[str, List[Dict[str, Any]]] = {tg: [] for tg in order}
        for item in pool:
            if item["uid"] in selected_uids:
                continue
            tg = classify_stratum_tag(item)
            tag_groups.setdefault(tg, []).append(item)

        picked: List[Dict[str, Any]] = []
        # Round-robin across strata tags for balanced boundary testing
        while len(picked) < k:
            progressed = False
            for tg in order:
                if tag_groups.get(tg) and len(picked) < k:
                    cand = tag_groups[tg].pop(0)
                    picked.append(cand)
                    selected_uids.add(cand["uid"])
                    progressed = True
            if not progressed:
                break
        return picked

    for bname, q in quotas:
        selected.extend(pick_diverse(by_bucket.get(bname, []), q))

    # Top up from remaining pool if any bucket fell short
    if len(selected) < target_n:
        rem = [r for r in h2_pool if r["uid"] not in selected_uids]
        selected.extend(pick_diverse(rem, target_n - len(selected)))

    return selected[:target_n]


# =====================================================================
# Dashboard JSON Builder
# =====================================================================
def _calc_sample_mean_sd(values: List[float]) -> Tuple[Optional[float], Optional[float]]:
    n = len(values)
    if n == 0:
        return None, None
    mean_v = sum(values) / n
    if n < 2:
        return mean_v, None
    var = sum((x - mean_v) ** 2 for x in values) / (n - 1)
    return mean_v, math.sqrt(var)


def _eval_12m_spc_status(
    points: List[Tuple[str, float]],
    exclude_partial_month: Optional[str] = None,
) -> Dict[str, Any]:
    """
    최근 12개월 이동 기저선(평균 ± 2SD) + 3개월 연속 상승 기반 증감 상태 판정.
    - exclude_partial_month: 진행 중인 당월(일수 미완료) 누적 건수 왜곡 방지를 위해
      CRE 월간 건수 평가 시 진행 중 당월을 평가 대상에서 제외하고 직전 완료월 기준으로 판정.
    """
    valid = [(p, v) for p, v in sorted(points, key=lambda x: x[0]) if v is not None]
    if not valid:
        return {
            "status": "insufficient",
            "label": "• 데이터없음",
            "evalPeriod": "",
            "evalValue": None,
            "isPartialLatest": False,
            "baselineN": 0,
            "mean": None,
            "sd": None,
            "lower2sd": None,
            "upper2sd": None,
            "zScore": None,
            "streak3Up": False,
            "vsMeanPct": None,
            "note": "데이터 없음",
        }

    is_partial_latest = bool(exclude_partial_month and valid[-1][0] == exclude_partial_month and len(valid) >= 2)
    eval_series = valid[:-1] if is_partial_latest else valid
    eval_period, eval_val = eval_series[-1]
    base_pts = eval_series[:-1][-12:]
    base_vals = [v for _, v in base_pts]
    n_base = len(base_vals)
    mean_v, sd_v = _calc_sample_mean_sd(base_vals)

    streak3_up = False
    if len(eval_series) >= 4:
        v0, v1, v2, v3 = [x[1] for x in eval_series[-4:]]
        streak3_up = (v0 < v1 < v2 < v3)

    vs_mean_pct = round(((eval_val - mean_v) / mean_v) * 100, 1) if (mean_v and mean_v > 0) else None
    z_score = None
    lower2sd = None
    upper2sd = None
    if mean_v is not None and sd_v is not None:
        lower2sd = max(0.0, mean_v - 2.0 * sd_v)
        upper2sd = mean_v + 2.0 * sd_v
        if sd_v > 0:
            z_score = round((eval_val - mean_v) / sd_v, 2)

    if n_base < 6:
        status, label = "insufficient", "• 축적중"
        note = f"기저선 축적 중 ({n_base}개월 · 6개월 이상 시 ±2SD 판정)"
    else:
        if sd_v is not None and sd_v > 0 and z_score is not None:
            if z_score > 2.0:
                status, label = "surge", "▲ 급증"
            elif z_score > 1.0 or (streak3_up and eval_val > mean_v):
                status, label = "up", "▲ 증가"
            elif z_score < -2.0:
                status, label = "down", "▼ 감소"
            else:
                status, label = "stable", "─ 유지"
        else:
            # sd == 0 (기저 기간 전체 동일값)
            if mean_v is not None and eval_val > mean_v:
                status, label = "up", "▲ 증가"
            elif mean_v is not None and eval_val < mean_v:
                status, label = "down", "▼ 감소"
            else:
                status, label = "stable", "─ 유지"
        note = f"직전 {n_base}개월 평균±2SD 기준 ({eval_period} 평가)"
        if is_partial_latest:
            note += f" · {valid[-1][0]} 집계중 제외"

    return {
        "status": status,
        "label": label,
        "evalPeriod": eval_period,
        "evalValue": round(eval_val, 3) if not float(eval_val).is_integer() else int(eval_val),
        "isPartialLatest": is_partial_latest,
        "baselineN": n_base,
        "mean": round(mean_v, 4) if mean_v is not None else None,
        "sd": round(sd_v, 4) if sd_v is not None else None,
        "lower2sd": round(lower2sd, 4) if lower2sd is not None else None,
        "upper2sd": round(upper2sd, 4) if upper2sd is not None else None,
        "zScore": z_score,
        "streak3Up": streak3_up,
        "vsMeanPct": vs_mean_pct,
        "note": note,
    }


def build_surveillance_payload(surv_rows: List[Dict[str, Any]], hai_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    now_dt = now_kst()
    current_year = str(now_dt.year)
    prev_year = str(now_dt.year - 1)
    # 당월 28일 이전이면 전수신고 월간 누적건수는 진행 중(집계중)으로 표시
    partial_ym = now_dt.strftime("%Y-%m") if now_dt.day < 28 else None

    by_disease: Dict[str, Dict[str, Any]] = {}
    all_periods: Set[str] = set()
    all_years: Set[str] = set()
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
        all_periods.add(p_key)
        all_years.add(p_key[:4])
        if name not in by_disease:
            by_disease[name] = {"name": name, "group": grp, "byPeriod": {}, "series": []}
        by_disease[name]["byPeriod"][p_key] = cnt

    for entry in by_disease.values():
        entry["series"] = [{"period": p, "count": c} for p, c in sorted(entry["byPeriod"].items())]

    class1_or_critical_re = re.compile(
        r"제1급|1급|콜레라|에볼라|마버그|라싸|페스트|탄저|두창|MERS|중동호흡기|SARS|중증급성호흡기|"
        r"니파|야토병|보툴리눔|디프테리아|신종인플루엔자|동물인플루엔자|VRSA|반코마이신내성",
        re.IGNORECASE,
    )

    # 1) 월별 전수감시 급증 스냅샷 (주간 브리핑 선택 주차 기준월 연동용)
    target_months = sorted([p for p in all_periods if re.match(r"^\d{4}-\d{2}$", p) and p >= f"{prev_year}-12"])
    if not target_months:
        target_months = sorted([p for p in all_periods if re.match(r"^\d{4}-\d{2}$", p)])[-12:]

    monthly_snapshots: Dict[str, List[Dict[str, Any]]] = {}
    for ym in target_months:
        yr_int = int(ym[:4])
        mm_str = ym[5:7]
        m_alerts: List[Dict[str, Any]] = []
        for name, entry in by_disease.items():
            cnt = entry["byPeriod"].get(ym)
            if cnt is None:
                continue
            is_cre = bool(re.search(r"CRE|카바페넴내성", name, re.IGNORECASE))
            is_critical_grp = bool(class1_or_critical_re.search(f"{entry['group']} {name}"))

            if is_critical_grp:
                if cnt >= 1:
                    m_alerts.append(
                        {
                            "region": "국내 · KDCA",
                            "level": "critical",
                            "method": "1급·초저발생 즉시감지",
                            "disease": name,
                            "period": ym,
                            "count": int(cnt),
                            "zScore": None,
                            "vsBaselinePct": None,
                            "reason": (f"국내 제1급 감염병 발생 감지 ({ym} · {int(cnt):,}건)" if "1급" in entry["group"] else f"국내 고위험·초저발생({entry['group'] or '전수감시'}) 발생 감지 ({ym} · {int(cnt):,}건)"),
                        }
                    )
                continue

            if is_cre:
                # 비계절성 우상향 내성균: 직전 12개월 이동평균 + 2SD 초과 시 경보 (진행 중 당월 제외)
                if partial_ym and ym == partial_ym:
                    continue
                prior_pts = [(p, c) for p, c in sorted(entry["byPeriod"].items()) if p < ym][-12:]
                if len(prior_pts) >= 6:
                    mean_v, sd_v = _calc_sample_mean_sd([c for _, c in prior_pts])
                    if mean_v and sd_v and sd_v > 0:
                        z = (cnt - mean_v) / sd_v
                        pct = round(((cnt - mean_v) / mean_v) * 100, 1)
                        if z > 2.0 and (cnt - mean_v) >= 50:
                            m_alerts.append(
                                {
                                    "region": "국내 · KDCA",
                                    "level": "surge",
                                    "method": f"최근 {len(prior_pts)}개월 평균+2SD",
                                    "disease": name,
                                    "period": ym,
                                    "count": int(cnt),
                                    "zScore": round(z, 2),
                                    "vsBaselinePct": pct,
                                    "reason": f"최근 {len(prior_pts)}개월 평균({mean_v:,.0f}건) 대비 +{pct}% (+{z:.1f} SD · 금월 {int(cnt):,}건)",
                                }
                            )
                continue

            # 계절성·일반 전수신고: 최근 3년(YYYY-3 ~ YYYY-1) 동월 기저선 평균 + 2SD 하이브리드
            hist_same_mm = []
            for delta_yr in (1, 2, 3):
                hk = f"{yr_int - delta_yr}-{mm_str}"
                if hk in entry["byPeriod"]:
                    hist_same_mm.append(entry["byPeriod"][hk])

            if len(hist_same_mm) >= 2:
                mean_v, sd_v = _calc_sample_mean_sd(hist_same_mm)
                if mean_v is not None and mean_v > 0:
                    pct = round(((cnt - mean_v) / mean_v) * 100, 1)
                    abs_diff = cnt - mean_v
                    z = ((cnt - mean_v) / sd_v) if (sd_v and sd_v > 0) else (9.9 if cnt > mean_v else 0.0)
                    is_surge = (
                        cnt >= 10
                        and abs_diff >= 10
                        and pct >= 30.0
                        and ((sd_v and sd_v > 0 and z > 2.0) or ((not sd_v or sd_v == 0) and cnt >= mean_v * 1.5))
                    )
                    if is_surge:
                        z_txt = f", +{z:.1f} SD" if (sd_v and sd_v > 0) else ""
                        m_alerts.append(
                            {
                                "region": "국내 · KDCA",
                                "level": "surge",
                                "method": f"최근 {len(hist_same_mm)}년 동월 평균+2SD",
                                "disease": name,
                                "period": ym,
                                "count": int(cnt),
                                "zScore": round(z, 2) if (sd_v and sd_v > 0) else None,
                                "vsBaselinePct": pct,
                                "reason": f"최근 {len(hist_same_mm)}년 동월 평균({mean_v:,.1f}건) 대비 +{pct}%{z_txt} (금월 {int(cnt):,}건)",
                            }
                        )
            elif len(hist_same_mm) == 1:
                # 2023~2024 소급 전 또는 신규 전수감시 전환 질환 안전 폴백: 전년 동월(YoY) 비교 (절대 직전 4개월 비교 금지)
                yoy_val = hist_same_mm[0]
                if yoy_val > 0:
                    pct = round(((cnt - yoy_val) / yoy_val) * 100, 1)
                    if cnt >= 10 and (cnt - yoy_val) >= 10 and pct >= 50.0:
                        m_alerts.append(
                            {
                                "region": "국내 · KDCA",
                                "level": "surge",
                                "method": "전년 동월 대비(YoY)",
                                "disease": name,
                                "period": ym,
                                "count": int(cnt),
                                "zScore": None,
                                "vsBaselinePct": pct,
                                "reason": f"전년 동월({int(yoy_val):,}건) 대비 +{pct}% (금월 {int(cnt):,}건 · 3년 기저선 확보 전)",
                            }
                        )

        m_alerts.sort(
            key=lambda x: (
                0 if x["level"] == "critical" else 1,
                -(x["zScore"] if x["zScore"] is not None else 0),
                -(x["vsBaselinePct"] if x["vsBaselinePct"] is not None else 0),
            )
        )
        monthly_snapshots[ym] = m_alerts[:8]

    latest_month = target_months[-1] if target_months else ""
    surge_alerts = monthly_snapshots.get(latest_month, [])

    # 2) CRE / VRSA 카드 및 CRE 12개월 ± 2SD 통계
    mdr_cards = []
    cre_series = {"currentYear": current_year, "current": [], "prev": []}
    cre_stats = _eval_12m_spc_status([])
    for name, entry in by_disease.items():
        s = entry["series"]
        if not s:
            continue
        is_cre = bool(re.search(r"CRE|카바페넴내성", name, re.IGNORECASE))
        is_vrsa = bool(re.search(r"VRSA|반코마이신내성", name, re.IGNORECASE))
        if not (is_cre or is_vrsa):
            continue
        latest = s[-1]
        yoy_key = compute_prev_year_period_key(latest["period"])
        yoy_pt = next((x for x in s if x["period"] == yoy_key), None) if yoy_key else None
        vs_yoy_pct = (
            round(((latest["count"] - yoy_pt["count"]) / yoy_pt["count"]) * 100, 1)
            if (yoy_pt and yoy_pt["count"] > 0)
            else None
        )
        spc = _eval_12m_spc_status([(x["period"], x["count"]) for x in s], exclude_partial_month=partial_ym if is_cre else None)
        mdr_cards.append(
            {
                "code": "CRE" if is_cre else "VRSA",
                "label": "CRE (카바페넴내성장내세균목)" if is_cre else "VRSA (반코마이신내성황색포도알균)",
                "period": latest["period"],
                "count": int(latest["count"]) if float(latest["count"]).is_integer() else latest["count"],
                "vsYoyPct": vs_yoy_pct,
                "spc": spc,
            }
        )
        if is_cre:
            cre_series["current"] = [x for x in s if x["period"].startswith(f"{current_year}-")]
            cre_series["prev"] = [x for x in s if x["period"].startswith(f"{prev_year}-")]
            cre_stats = spc

    mdr_cards.sort(key=lambda x: 0 if x["code"] == "CRE" else 1)

    # 3) 4급 HAI 5종 12개월 기저선 ± 2SD 증감 판정
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
        # 월 중복 방어
        dedup_p = {x["period"]: x for x in hai_by_patho.get(p, [])}
        arr = [dedup_p[k] for k in sorted(dedup_p.keys())]
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
        blood_spc = _eval_12m_spc_status([(x["period"], x["blood"]) for x in arr if x["blood"] is not None])
        non_blood_spc = _eval_12m_spc_status([(x["period"], x["nonBlood"]) for x in arr if x["nonBlood"] is not None])
        hai_cards.append(
            {
                "pathogen": p,
                "period": latest["period"],
                "bloodRate": latest["blood"],
                "nonBloodRate": latest["nonBlood"],
                "bloodDeltaPct": blood_delta,
                "nonBloodDeltaPct": non_blood_delta,
                "bloodSpc": blood_spc,
                "nonBloodSpc": non_blood_spc,
            }
        )
        cur_yr = [x for x in arr if x["period"].startswith(f"{current_year}-")]
        hai_annual_series.append(
            {
                "pathogen": p,
                "points": cur_yr if cur_yr else arr,
                "bloodSpc": blood_spc,
                "nonBloodSpc": non_blood_spc,
            }
        )

    return {
        "year": current_year,
        "updatedAt": latest_updated_at,
        "latestMonth": latest_month,
        "partialMonth": partial_ym,
        "baselineYears": sorted(all_years),
        "mdrCards": mdr_cards,
        "haiCards": hai_cards,
        "haiAnnualSeries": hai_annual_series,
        "creSeries": cre_series,
        "creStats": cre_stats,
        "surgeAlerts": surge_alerts,
        "monthlySnapshots": monthly_snapshots,
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
        if ctype == "duplicate":
            continue
        week_key = safe_str(r.get("week_key")) or get_iso_week_key(pub_date)
        raw_score = safe_float(r.get("importance_score"))
        score_val = int(round(raw_score)) if raw_score is not None else None
        ce_code = safe_str(r.get("critical_exception")).upper()
        if ce_code not in EXCEPTION_LABELS:
            ce_code = ""
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
                "critical_exception": ce_code,
                "exception_reason": safe_str(r.get("exception_reason")) if ce_code else "",
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

    items.sort(key=lambda x: (x["pub_date"], 1 if x["critical_exception"] else 0, x["importance_score"] or -1, x["uid"]), reverse=True)

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

    lite_model, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg_map)
    cutoffs = get_score_cutoffs(cfg_map)
    payload = {
        "meta": {
            "version": CODE_VERSION,
            "generated_at": now_kst_str(),
            "total_items": len(items),
            "serving_mode": "GitHub Pages (메인 단일화)",
            "h2_guard_date": H2_GUARD_START_DATE,
            "models": {
                "deep": deep_model,
                "fallback": fallback_model,
                "lite": lite_model,
                "thinking_p1": thinking_p1,
                "thinking_p2": thinking_p2,
            },
            "cutoffs": cutoffs,
        },
        "items": items,
        "logs": logs,
        "surveillance": build_surveillance_payload(surv_rows, hai_rows),
        "notifiableDiseases": json.loads(NOTIFIABLE_DISEASES_MASTER_JSON),
        "referenceStandards": json.loads(REFERENCE_STANDARDS_MASTER_JSON),
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
# Shared 2-Pass Execution Helper for Papers & Guidelines
# =====================================================================
def execute_two_pass_chunk(
    chunk: List[Dict[str, Any]],
    deep_model: str,
    fallback_model: str,
    gemini_key: str,
    thinking_p1: str,
    thinking_p2: str,
    cutoffs: Dict[str, int],
) -> Tuple[Dict[str, Dict[str, Any]], int, int]:
    """Pass 1(low) 전수 채점 -> Must-Read 또는 Critical Exception 대상만 Pass 2(high) 심층 요약"""
    p1_map, used_p1 = summarize_deep_pass1_batch(
        chunk, deep_model, fallback_model, gemini_key, thinking_p1, cutoffs
    )
    fallback_hits = len(chunk) if used_p1 != deep_model else 0
    pass2_count = 0

    for it in chunk:
        ai = p1_map.get(it["uid"])
        if not ai:
            continue
        pass2_done = False
        used_p2 = ""
        if ai.get("needs_pass2"):
            try:
                p2_res, used_p2 = run_pass2_deep_summary(
                    it, ai, deep_model, fallback_model, gemini_key, thinking_p2
                )
                if used_p2 != deep_model:
                    fallback_hits += 1
                ai["title_ko"] = p2_res.get("title_ko") or ai.get("title_ko") or it["title"]
                ai["fact_summary_ko"] = p2_res.get("fact_summary_ko") or ai.get("fact_summary_ko", "")
                ai["ai_clinical_note"] = p2_res.get("ai_clinical_note") or ai.get("ai_clinical_note", "")
                pass2_done = True
                pass2_count += 1
                print(
                    f"    ★ Pass 2 심층 요약 완료: {it['uid']} "
                    f"[{ai['importance']} {ai['importance_score']}점"
                    f"{' / ' + ai['critical_exception'] if ai['critical_exception'] else ''}] ({used_p2})"
                )
                time.sleep(1.0)
            except Exception as p2_err:
                print(f"    [Pass 2 경고 — Pass 1 결과 보존] {it['uid']}: {p2_err}")

        ai["pass2_applied"] = pass2_done
        if pass2_done:
            ai["summary_meta"] = f"{CODE_VERSION}|P1:{used_p1}({thinking_p1})+P2:{used_p2}({thinking_p2})"
        else:
            ai["summary_meta"] = f"{CODE_VERSION}|P1:{used_p1}({thinking_p1})"
        ai["formatted_note"] = format_clinical_note(
            ai.get("ai_clinical_note", ""),
            ai.get("importance", "Medium"),
            ai.get("importance_score", 0),
            ai.get("critical_exception", ""),
            pass2_done,
        )
    return p1_map, fallback_hits, pass2_count


# =====================================================================
# Main Modes
# =====================================================================
def mode_test_apis(allow_h1: bool = False) -> None:
    print(f"=== 임상미생물 브리핑 {CODE_VERSION} API 연결 진단 ===")
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    lite_model, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    eff_start, guard_active = resolve_effective_paper_start(cfg, allow_h1)
    print(f"[OK] Google Sheets 연결 성공: {ss.title}")
    print(f"   논문 수집 시작일: {eff_start} (상반기 소급 가드: {'ON - 7월 이후 보호' if guard_active else 'OFF'})")
    print(f"   2-Pass 모델: P1={deep_model}({thinking_p1}) / P2={deep_model}({thinking_p2}) / 우회={fallback_model} / 경보={lite_model}")
    print(f"   4단계 컷오프: Must-Read>={cutoffs['must_read']}, High>={cutoffs['high']}, Medium>={cutoffs['medium']}")

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()

    if gemini_key:
        print(f"1. Gemini 경량({lite_model}): {probe_gemini_model(lite_model, gemini_key)}")
        print(f"   Gemini 주력({deep_model}): {probe_gemini_model(deep_model, gemini_key)}")
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


def mode_pilot_20() -> None:
    """
    [1순위 검증 모드] 본 DB(items)를 전혀 수정하지 않고,
    2026-07-01 이후 논문·가이드라인 중 층화 표본 20편을 추출하여
    v8.2 루브릭(35/30/20/15)·유형별 Cap·Critical Exception·2-Pass 차등 요약을 실행하고
    별도 시트 탭(pilot_v82)에만 출력함.
    """
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    item_ws = ss.worksheet(SHEET_ITEMS)

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()
    if not gemini_key:
        raise RuntimeError("GEMINI_API_KEY 가 설정되지 않았습니다.")

    _, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    run_id = "GH_PILOT20_" + now_kst().strftime("%Y%m%d_%H%M%S")
    now_str = now_kst_str()

    all_values = item_ws.get_all_values()
    if len(all_values) < 2:
        print("파일럿 대상 데이터가 없습니다.")
        return

    records = []
    for row in all_values[1:]:
        r = row + [""] * max(0, 23 - len(row))
        uid = r[0].strip()
        if not uid:
            continue
        records.append(
            {
                "uid": uid,
                "source": r[1].strip(),
                "pubDate": r[2].strip()[:10],
                "title": r[4].strip(),
                "oldImportance": r[7].strip(),
                "pmid": clean_pmid(r[10]),
                "doi": r[11].strip(),
                "sourceUrl": r[12].strip(),
                "contentType": r[15].strip() or "paper",
            }
        )

    sample_20 = select_stratified_pilot_20(records, target_n=20)
    print(f"=== [pilot_20] 2026-07 이후 층화 표본 {len(sample_20)}편 선정 (본 DB items 보호) ===")
    for s in sample_20:
        print(f"  - {s['uid']} | {s['source']} | {s['pubDate']} | 기존:{s['oldImportance']} | {s['title'][:60]}")

    pmids = [t["pmid"] for t in sample_20 if t["pmid"]]
    abs_by_pmid: Dict[str, str] = {}
    if pmids:
        print(f"PubMed 원문 초록 {len(pmids)}건 조회 중...")
        for art in fetch_pubmed_articles(pmids, ncbi_key):
            abs_by_pmid[str(art["pmid"])] = art["abstractText"]

    for t in sample_20:
        if t["pmid"] and t["pmid"] in abs_by_pmid:
            t["rawText"] = abs_by_pmid[t["pmid"]]
        elif t["doi"]:
            t["rawText"] = fetch_crossref_abstract(t["doi"])
            time.sleep(0.15)
        else:
            t["rawText"] = ""

    pilot_headers = [
        "uid",
        "source",
        "pub_date",
        "content_type",
        "old_importance",
        "new_importance",
        "importance_score",
        "raw_score",
        "study_type",
        "cap_applied",
        "critical_exception",
        "exception_reason",
        "pass2_applied",
        "score_breakdown",
        "title_ko",
        "fact_summary_ko",
        "ai_clinical_note",
        "importance_reason",
        "summary_meta",
        "source_url",
    ]
    out_rows: List[List[Any]] = []
    tier_counts = {"Must-Read": 0, "High": 0, "Medium": 0, "Low": 0}
    cap_hits = 0
    ce_hits = 0
    p2_hits = 0
    fb_hits = 0
    errors: List[str] = []

    for offset in range(0, len(sample_20), GEMINI_DEEP_BATCH_SIZE):
        chunk = sample_20[offset : offset + GEMINI_DEEP_BATCH_SIZE]
        try:
            ai_map, fbh, p2h = execute_two_pass_chunk(
                chunk, deep_model, fallback_model, gemini_key, thinking_p1, thinking_p2, cutoffs
            )
            fb_hits += fbh
            p2_hits += p2h
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                tier = ai["importance"]
                tier_counts[tier] = tier_counts.get(tier, 0) + 1
                if ai.get("cap_label"):
                    cap_hits += 1
                if ai.get("critical_exception"):
                    ce_hits += 1
                out_rows.append(
                    [
                        t["uid"],
                        t["source"],
                        t["pubDate"],
                        t["contentType"],
                        t["oldImportance"],
                        tier,
                        ai["importance_score"],
                        ai["raw_score"],
                        ai["study_type"],
                        ai["cap_label"] or "-",
                        ai["critical_exception"] or "-",
                        ai["exception_reason"],
                        "TRUE" if ai["pass2_applied"] else "FALSE",
                        ai["score_breakdown"],
                        ai.get("title_ko") or t["title"],
                        ai.get("fact_summary_ko") or "",
                        ai.get("formatted_note") or "",
                        ai.get("importance_reason") or "",
                        ai.get("summary_meta") or "",
                        t["sourceUrl"],
                    ]
                )
            time.sleep(1.2)
        except Exception as e:
            errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
            print(f"  [파일럿 오류] {chunk[0]['uid']}: {e}")

    # Sort output rows: Critical Exception first, then importance_score descending
    out_rows.sort(key=lambda r: (1 if r[10] != "-" else 0, safe_int(r[6])), reverse=True)

    pilot_ws = get_or_create_worksheet(ss, SHEET_PILOT, pilot_headers)
    pilot_ws.clear()
    pilot_ws.update(range_name="A1", values=[pilot_headers] + out_rows, value_input_option="RAW")

    summary_note = (
        f"파일럿 20편({CODE_VERSION}) 완료 -> {SHEET_PILOT} 탭 저장 (본 DB 미변경) | "
        f"Must-Read:{tier_counts['Must-Read']}, High:{tier_counts['High']}, Medium:{tier_counts['Medium']}, Low:{tier_counts['Low']} | "
        f"Cap적용:{cap_hits}건, CriticalException:{ce_hits}건, Pass2심층:{p2_hits}건"
    )
    if fb_hits > 0:
        summary_note += f" (503우회:{fb_hits}건)"
    if errors:
        summary_note += " / 오류: " + " | ".join(errors[:3])
    print(f"[OK] {summary_note}")
    status = "Success" if not errors else ("Partial" if out_rows else "Fail")
    append_run_log(ss, run_id, now_str, "PILOT_20", status, len(sample_20), len(out_rows), 0, summary_note)


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
        r = row + [""] * max(0, 23 - len(row))
        if r[15].strip() == "duplicate":
            continue
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


def mode_collect(allow_h1: bool = False) -> None:
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

    paper_start, guard_active = resolve_effective_paper_start(cfg, allow_h1)
    if guard_active:
        print(f"[상반기 소급 가드 활성] 논문 수집 범위를 {paper_start} 이후로 제한합니다 (상반기 소급 시 --allow-h1 사용).")
    year_start = H2_GUARD_START_DATE if guard_active else (cfg.get("kdca_start_date", "").strip() or DEFAULT_YEAR_START_DATE)
    guideline_start = cfg.get("kdca_start_date", "").strip() or DEFAULT_YEAR_START_DATE

    paper_days = calc_days_since(paper_start, 160)
    year_days = calc_days_since(year_start, 160)
    guideline_days = calc_days_since(guideline_start, 365)
    fetch_limit = safe_int(cfg.get("pubmed_fetch_limit"), 1000)

    lite_model, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    pubmed_delay = 0.15 if ncbi_key else 0.4

    all_existing_vals = item_ws.get_all_values()
    existing_uids: Set[str] = set()
    existing_records: List[Dict[str, Any]] = []
    for row in all_existing_vals[1:]:
        r = row + [""] * max(0, 23 - len(row))
        u = r[0].strip()
        if not u:
            continue
        existing_uids.add(u)
        existing_records.append(
            {
                "uid": u,
                "source": r[1].strip(),
                "pubDate": r[2].strip()[:10],
                "title": r[4].strip(),
                "pmid": clean_pmid(r[10]),
                "doi": r[11].strip(),
                "contentType": r[15].strip() or "paper",
            }
        )
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

    # 2) 가이드라인 (IDSA · ESCMID — 2026년 전체, 편수 적음)
    if is_enabled(cfg, "source_guideline_enabled"):
        try:
            g_query = (
                '(("Infectious Diseases Society of America"[Title] OR "IDSA"[Title] OR "ESCMID"[Title]'
                ' OR "European Society of Clinical Microbiology"[Title]) AND ("Guideline"[Title] OR "Guidance"[Title] OR "Practice Guideline"[Publication Type]))'
                + PUBMED_CLEAN_FILTER
            )
            g_ids = fetch_pubmed_ids(g_query, guideline_days, fetch_limit, ncbi_key)
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

    # 3) 4대 저널 (JCM · CMI · CID · ACM) — 상반기 소급 가드 적용
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
                    pd_obj = parse_flexible_date(a["pubDate"])
                    if paper_cutoff_date and pd_obj and pd_obj < paper_cutoff_date:
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

    candidates, copub_dups = detect_cross_journal_duplicates(candidates, existing_records, ncbi_key)
    copub_dup_count = len(copub_dups)
    if copub_dups:
        dup_rows = []
        for it in copub_dups:
            week_key = get_iso_week_key(it["pubDate"] or today_str)
            dup_of = it.get("duplicateOf", "")
            dup_rows.append(
                [
                    it["uid"],
                    it["source"],
                    it["pubDate"],
                    now_str,
                    it["title"],
                    it["title"],
                    "기타",
                    "Low",
                    f"이종 저널 동시 게재 중복 항목 (원본: {dup_of})",
                    "",
                    it.get("pmid", ""),
                    it.get("doi", ""),
                    it.get("sourceUrl", ""),
                    it.get("accessStatus", "Public"),
                    False,
                    "duplicate",
                    week_key,
                    f"{CODE_VERSION}|duplicate_of:{dup_of}",
                    f"이종 저널 초록 지문 일치로 자동 비노출 마킹 (원본: {dup_of})",
                    "",
                    "",
                    "",
                    "",
                ]
            )
        item_ws.append_rows(dup_rows, value_input_option="RAW")
        total_dup += copub_dup_count
        status_notes.append(f"동시게재 중복 자동차단 {copub_dup_count}건(0원)")

    print(f"[수집 대상 확정] 논문 시작일:{paper_start} / 조회 {total_fetched}건 / 중복 {total_dup}건(동시게재 {copub_dup_count}건 포함) / 신규 요약 대상 {len(candidates)}건")

    outbreak_q = [x for x in candidates if x["contentType"] == "outbreak"]
    deep_q = [x for x in candidates if x["contentType"] != "outbreak"]
    new_count = 0
    fallback_hits = 0
    pass2_hits = 0
    ai_errors: List[str] = []

    # Outbreak queue (lite)
    for offset in range(0, len(outbreak_q), GEMINI_LITE_BATCH_SIZE):
        chunk = outbreak_q[offset : offset + GEMINI_LITE_BATCH_SIZE]
        try:
            ai_map, used_model = summarize_outbreak_batch(
                chunk, lite_model, fallback_model, gemini_key, cutoffs
            )
            if used_model != lite_model:
                fallback_hits += len(chunk)
            rows_to_append = []
            for it in chunk:
                ai = ai_map.get(it["uid"])
                if not ai:
                    continue
                week_key = get_iso_week_key(it["pubDate"] or today_str)
                note = format_clinical_note(
                    ai.get("ai_clinical_note", ""),
                    ai["importance"],
                    ai["importance_score"],
                    ai["critical_exception"],
                    False,
                )
                rows_to_append.append(
                    [
                        it["uid"],
                        it["source"],
                        it["pubDate"],
                        now_str,
                        it["title"],
                        ai.get("title_ko") or it["title"],
                        ai.get("topic_category") or "국내 긴급경보(KDCA)",
                        ai["importance"],
                        ai.get("fact_summary_ko") or "",
                        note,
                        "",
                        "",
                        it.get("sourceUrl", ""),
                        it.get("accessStatus", "Public"),
                        False,
                        "outbreak",
                        week_key,
                        f"{CODE_VERSION}|{used_model}",
                        ai.get("importance_reason", ""),
                        ai["importance_score"],
                        ai["score_breakdown"],
                        ai["critical_exception"],
                        ai["exception_reason"],
                    ]
                )
            if rows_to_append:
                item_ws.append_rows(rows_to_append, value_input_option="RAW")
                new_count += len(rows_to_append)
                print(f"  + 경보 {len(rows_to_append)}건 저장 ({used_model}, 누적 {new_count}/{len(candidates)})")
            time.sleep(1.0)
        except Exception as e:
            ai_errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
            print(f"  [경보 요약 오류] {chunk[0]['uid']}: {e}")

    # Deep queue (2-Pass)
    for offset in range(0, len(deep_q), GEMINI_DEEP_BATCH_SIZE):
        chunk = deep_q[offset : offset + GEMINI_DEEP_BATCH_SIZE]
        try:
            ai_map, fbh, p2h = execute_two_pass_chunk(
                chunk, deep_model, fallback_model, gemini_key, thinking_p1, thinking_p2, cutoffs
            )
            fallback_hits += fbh
            pass2_hits += p2h
            rows_to_append = []
            for it in chunk:
                ai = ai_map.get(it["uid"])
                if not ai:
                    continue
                week_key = get_iso_week_key(it["pubDate"] or today_str)
                rows_to_append.append(
                    [
                        it["uid"],
                        it["source"],
                        it["pubDate"],
                        now_str,
                        it["title"],
                        ai.get("title_ko") or it["title"],
                        ai.get("topic_category") or "기타",
                        ai["importance"],
                        ai.get("fact_summary_ko") or "",
                        ai["formatted_note"],
                        it.get("pmid", ""),
                        it.get("doi", ""),
                        it.get("sourceUrl", ""),
                        it.get("accessStatus", "Public"),
                        False,
                        it["contentType"],
                        week_key,
                        ai["summary_meta"],
                        ai.get("importance_reason", ""),
                        ai["importance_score"],
                        ai["score_breakdown"],
                        ai["critical_exception"],
                        ai["exception_reason"],
                    ]
                )
            if rows_to_append:
                item_ws.append_rows(rows_to_append, value_input_option="RAW")
                new_count += len(rows_to_append)
                print(f"  + 논문·지침 {len(rows_to_append)}건 저장 (누적 {new_count}/{len(candidates)})")
            time.sleep(1.2)
        except Exception as e:
            ai_errors.append(f"{chunk[0]['uid']}: {str(e)[:120]}")
            print(f"  [요약 오류] {chunk[0]['uid']}: {e}")

    notes = status_notes + [
        f"GitHub Actions({CODE_VERSION}) · 논문기준:{paper_start}({'하반기가드ON' if guard_active else '전체'}) · Pass2심층:{pass2_hits}건"
    ]
    if fallback_hits > 0:
        notes.append(f"503 우회 {fallback_hits}건")
    if not candidates and not fetch_errors:
        notes.append("신규 항목 없음")
    if fetch_errors:
        notes.append("조회 오류: " + " | ".join(fetch_errors[:3]))
    if ai_errors:
        notes.append("요약 오류: " + " | ".join(ai_errors[:3]))
    status = "Success" if (not fetch_errors and not ai_errors) else ("Partial" if new_count > 0 else "Fail")
    append_run_log(ss, run_id, now_str, "ALL", status, total_fetched, new_count, total_dup, " / ".join(notes))

    sync_sheet_to_dashboard_json(ss)


def mode_resummarize_all(force: bool = False, allow_h1: bool = False) -> None:
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    item_ws = ss.worksheet(SHEET_ITEMS)
    ensure_item_meta_columns(item_ws)

    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()
    if not gemini_key:
        raise RuntimeError("GEMINI_API_KEY 가 설정되지 않았습니다.")

    lite_model, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    eff_start, guard_active = resolve_effective_paper_start(cfg, allow_h1)
    run_id = "GH_RESUM_" + now_kst().strftime("%Y%m%d_%H%M%S")
    now_str = now_kst_str()

    all_values = item_ws.get_all_values()
    if len(all_values) < 2:
        print("재요약할 데이터가 없습니다.")
        return

    deep_targets = []
    outbreak_targets = []
    h1_skipped = 0
    for idx, row in enumerate(all_values[1:], start=2):
        r = row + [""] * max(0, 23 - len(row))
        uid = r[0].strip()
        if not uid:
            continue
        pub_date = r[2].strip()[:10]
        ctype = r[15].strip() or "paper"
        meta = r[17].strip()
        if ctype == "duplicate":
            continue
        if (not force) and any(meta.startswith(p) for p in (f"{CODE_VERSION}|", "v8.3|", "v8.2.1|", "v8.2|")):
            continue
        # 상반기 소급 가드: allow_h1=False일 때 2026-07-01 이전 논문은 익월 크레딧 단계로 보류
        if guard_active and ctype == "paper" and pub_date and pub_date < H2_GUARD_START_DATE:
            h1_skipped += 1
            continue
        rec = {
            "rowIndex": idx,
            "uid": uid,
            "source": r[1].strip(),
            "pubDate": pub_date,
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
        f"[v8.2 재요약 대상] 논문·가이드라인 {len(deep_targets)}건 (상반기가드 제외 {h1_skipped}건) / "
        f"감염병 경보 {len(outbreak_targets)}건"
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
    pass2_hits = 0
    errors: List[str] = []

    for offset in range(0, len(deep_targets), GEMINI_DEEP_BATCH_SIZE):
        chunk = deep_targets[offset : offset + GEMINI_DEEP_BATCH_SIZE]
        try:
            ai_map, fbh, p2h = execute_two_pass_chunk(
                chunk, deep_model, fallback_model, gemini_key, thinking_p1, thinking_p2, cutoffs
            )
            fallback_hits += fbh
            pass2_hits += p2h
            batch_updates = []
            for t in chunk:
                ai = ai_map.get(t["uid"])
                if not ai:
                    continue
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "기타",
                            ai["importance"],
                            ai.get("fact_summary_ko") or "",
                            ai["formatted_note"],
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:W{row_i}",
                        "values": [[
                            ai["summary_meta"],
                            ai.get("importance_reason", ""),
                            ai["importance_score"],
                            ai["score_breakdown"],
                            ai["critical_exception"],
                            ai["exception_reason"],
                        ]],
                    }
                )
                updated_deep += 1
            if batch_updates:
                item_ws.batch_update(batch_updates, value_input_option="RAW")
            print(f"  [논문·가이드라인] {updated_deep}/{len(deep_targets)}건 갱신 (Pass2 누계:{pass2_hits})")
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
                note = format_clinical_note(
                    ai.get("ai_clinical_note", ""),
                    ai["importance"],
                    ai["importance_score"],
                    ai["critical_exception"],
                    False,
                )
                row_i = t["rowIndex"]
                batch_updates.append(
                    {
                        "range": f"F{row_i}:J{row_i}",
                        "values": [[
                            ai.get("title_ko") or t["title"],
                            ai.get("topic_category") or "국내 긴급경보(KDCA)",
                            ai["importance"],
                            ai.get("fact_summary_ko") or t["rawText"],
                            note,
                        ]],
                    }
                )
                batch_updates.append(
                    {
                        "range": f"R{row_i}:W{row_i}",
                        "values": [[
                            f"{CODE_VERSION}|{used_model}",
                            ai.get("importance_reason", ""),
                            ai["importance_score"],
                            ai["score_breakdown"],
                            ai["critical_exception"],
                            ai["exception_reason"],
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
        f"재요약({CODE_VERSION}) 완료: 논문·지침 {updated_deep}/{len(deep_targets)}(Pass2:{pass2_hits}), "
        f"경보 {updated_ob}/{len(outbreak_targets)} | 상반기가드 제외:{h1_skipped}건"
    )
    if fallback_hits > 0:
        note += f" (503우회:{fallback_hits}건)"
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
    cfg_map["gemini_thinking_level_pass2"] = "high"
    cfg_map.setdefault("cutoff_must_read", "85")
    cfg_map.setdefault("cutoff_high", "70")
    cfg_map.setdefault("cutoff_medium", "45")
    write_dashboard_json(items, surv, hai, logs, cfg_map)


def mode_sync_sheet_only() -> None:
    """[AI API 비용 0원] KDCA 전수신고 4년치(최근 3년 동월 기저선) 자동 보강 + 시트 수정분(hai_stats 등) 대시보드 즉시 반영"""
    ss = connect_spreadsheet()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    now_str = now_kst_str()
    run_id = "GH_SYNC_" + now_kst().strftime("%Y%m%d_%H%M%S")
    kdca_note = ""
    kdca_cnt = 0
    if kdca_key:
        try:
            kdca_cnt, kdca_note = sync_kdca_surveillance_api(ss, kdca_key, now_str)
            print(f"[KDCA 4년치 동기화 완료 (AI 비용 0원)] {kdca_note}")
        except Exception as e:
            kdca_note = f"KDCA API 동기화 건너뜀({str(e)[:120]}) — 기존 시트 데이터 유지"
            print(f"[KDCA 경고] {kdca_note}")
    else:
        kdca_note = "KDCA_API_KEY 미설정 — 기존 시트 데이터로 갱신"
    append_run_log(
        ss,
        run_id,
        now_str,
        "SYNC_SHEET",
        "Success",
        kdca_cnt,
        kdca_cnt,
        0,
        f"{CODE_VERSION} AI비용0원 시트·전수감시 동기화 · {kdca_note}",
    )
    sync_sheet_to_dashboard_json(ss)


def mode_sync_kdca_and_dashboard() -> None:
    """[AI API 비용 0원] KDCA 전수신고 4년치(최근 3년 동월 기저선 포함) 소급 갱신 + dashboard.json 재생성"""
    ss = connect_spreadsheet()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    now_str = now_kst_str()
    run_id = "GH_KDCA_" + now_kst().strftime("%Y%m%d_%H%M%S")
    if not kdca_key:
        raise RuntimeError("KDCA_API_KEY 가 설정되지 않았습니다.")
    cnt, note = sync_kdca_surveillance_api(ss, kdca_key, now_str)
    print(f"[KDCA 4년치 소급 완료 (AI 비용 0원)] {note}")
    append_run_log(ss, run_id, now_str, "KDCA_BACKFILL", "Success", cnt, cnt, 0, f"{CODE_VERSION} AI비용0원 · {note}")
    sync_sheet_to_dashboard_json(ss)


def main():
    parser = argparse.ArgumentParser(description="임상미생물 주간 브리핑 파이프라인 v8.3.0")
    parser.add_argument(
        "--mode",
        choices=["sync_kdca_and_dashboard", "sync_sheet_only", "collect", "resummarize_all", "reclassify_only", "pilot_20", "test_apis", "build_from_csv"],
        default="collect",
    )
    parser.add_argument("--allow-h1", action="store_true", help="2026년 상반기(1~6월) 소급 가드 해제")
    parser.add_argument("--force", action="store_true", help="이미 v8.2 완료된 항목도 강제 재요약")
    parser.add_argument("--csv-dir", default=".", help="build_from_csv 모드에서 CSV가 있는 폴더")
    args = parser.parse_args()

    if args.mode == "build_from_csv":
        mode_build_from_csv(Path(args.csv_dir))
    elif args.mode == "test_apis":
        mode_test_apis(allow_h1=args.allow_h1)
    elif args.mode == "pilot_20":
        mode_pilot_20()
    elif args.mode == "reclassify_only":
        mode_reclassify_only()
    elif args.mode == "sync_kdca_and_dashboard":
        mode_sync_kdca_and_dashboard()
    elif args.mode == "sync_sheet_only":
        mode_sync_sheet_only()
    elif args.mode == "resummarize_all":
        mode_resummarize_all(force=args.force, allow_h1=args.allow_h1)
    else:
        mode_collect(allow_h1=args.allow_h1)


if __name__ == "__main__":
    main()
