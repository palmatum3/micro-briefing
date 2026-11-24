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

CODE_VERSION = "v8.4.3"
KST = datetime.timezone(datetime.timedelta(hours=9))

SHEET_ITEMS = "items"
SHEET_LOGS = "run_logs"
SHEET_CONFIG = "config"
SHEET_KDCA_SURV = "kdca_surveillance"
SHEET_HAI_STATS = "hai_stats"
SHEET_PILOT = "pilot_v82"
SHEET_DRYRUN = "dryrun_v83"
TREND_CONFIG_KEY = "paper_macro_trend_v843"

# v8.4.2 동적 분기별 MDS 키워드 추출 풀 (임상미생물·감염병 핵심 44개 개념 + 미등록 신종 용어 자동 발굴)
DYNAMIC_CONCEPT_POOL_V842: List[Tuple[str, str, str, str]] = [
    # Pathogens (균종·감염병)
    ("tb_ntm", "결핵균·NTM", "pathogen", r"\b(?:tuberculosis|Mycobacterium|abscessus|NTM|MTB)\b"),
    ("influenza", "인플루엔자", "pathogen", r"\b(?:influenza|H3N2|H1N1|subclade K)\b"),
    ("rsv", "RSV", "pathogen", r"\b(?:RSV|respiratory syncytial)\b"),
    ("covid19", "COVID-19·SARS-CoV-2", "pathogen", r"\b(?:SARS-CoV-2|COVID-19|coronavirus)\b"),
    ("cre_entero", "CRE·장내세균목", "pathogen", r"\b(?:CRE|CPE|Enterobacterales|Klebsiella pneumoniae|Escherichia coli|carbapenemase|ertapenem-resistant)\b"),
    ("aspergillus", "아스페르길루스·사상균", "pathogen", r"\b(?:Aspergillus|aspergillosis|mucormycosis|mold|mould)\b"),
    ("candida", "Candida·C. auris", "pathogen", r"\b(?:Candida|Candidozyma|candidemia|candidiasis)\b"),
    ("staph_mrsa", "황색포도알균·MRSA", "pathogen", r"\b(?:Staphylococcus aureus|MRSA|MSSA|VRSA)\b"),
    ("pa_ab", "녹농균·아시네토박터", "pathogen", r"\b(?:Pseudomonas aeruginosa|Acinetobacter baumannii|CRAB|DTR)\b"),
    ("hiv", "HIV·AIDS", "pathogen", r"\b(?:HIV|HIV-1|antiretroviral|PrEP)\b"),
    ("enterococcus_vre", "장알균·VRE", "pathogen", r"\b(?:Enterococcus|VRE|enterococcal)\b"),
    ("cdiff", "C. difficile", "pathogen", r"\b(?:Clostridioides difficile|C\. difficile|CDI)\b"),
    ("measles", "홍역(Measles)", "pathogen", r"\b(?:measles|morbillivirus)\b"),
    ("malaria", "말라리아", "pathogen", r"\b(?:malaria|Plasmodium)\b"),
    ("hepatitis", "바이러스 간염(HBV·HCV·HEV)", "pathogen", r"\b(?:hepatitis [ABCE]|HBV|HCV|HEV|HBsAg)\b"),
    ("sti_syph_gono", "매독·임균·성매개감염", "pathogen", r"\b(?:syphilis|Treponema|gonorrhoeae|Chlamydia|sexually transmitted)\b"),
    ("strep_pneumo_gas", "폐렴구균·A군연쇄구균", "pathogen", r"\b(?:Streptococcus pneumoniae|pneumococcal|Streptococcus pyogenes|Group A Streptococcus)\b"),
    ("arbovirus", "뎅기·치쿤구니야·모기매개", "pathogen", r"\b(?:dengue|chikungunya|Zika|Oropouche|arboviral|West Nile)\b"),
    ("vhf_ebola_lassa", "출혈열(에볼라·라싸·CCHF)", "pathogen", r"\b(?:Ebola|Bundibugyo|Lassa|Crimean-Congo|CCHF|Marburg|Nipah)\b"),
    ("endocarditis", "감염성 심내막염(IE)", "pathogen", r"\b(?:infective endocarditis|endocarditis|Duke criteria)\b"),
    ("sepsis_bsi", "패혈증·균혈증(BSI)", "pathogen", r"\b(?:sepsis|septic shock|bacteremia|bloodstream infection)\b"),
    ("cryptococcus", "크립토콕쿠스·풍토진균", "pathogen", r"\b(?:Cryptococcus|cryptococcosis|Histoplasma|histoplasmosis|blastomycosis)\b"),
    ("enteric_bact", "장관세균(살모넬라·캄필로)", "pathogen", r"\b(?:Salmonella|Campylobacter|Shigella|typhoid|cholera)\b"),
    ("mpox", "엠폭스(Mpox)", "pathogen", r"\b(?:mpox|monkeypox)\b"),
    ("adenovirus_cmv", "아데노·CMV·이식바이러스", "pathogen", r"\b(?:adenovirus|cytomegalovirus|CMV|BK virus|JC virus|HTLV)\b"),
    # Antimicrobials / Vaccines (항균제·백신)
    ("carbapenem", "카바페넴계", "drug", r"\b(?:carbapenem|meropenem|imipenem|ertapenem)\b"),
    ("new_bl_bli", "신약 BL/BLI", "drug", r"\b(?:aztreonam-avibactam|relebactam|funobactam|zidebactam|ceftazidime-avibactam|vaborbactam|BL/BLI)\b"),
    ("cefiderocol", "세피데로콜", "drug", r"\b(?:cefiderocol)\b"),
    ("glyco_oxa", "반코마이신·리네졸리드", "drug", r"\b(?:vancomycin|linezolid|daptomycin|teicoplanin)\b"),
    ("antifungal", "항진균제(아졸·암포·에키노)", "drug", r"\b(?:voriconazole|posaconazole|amphotericin|echinocandin|caspofungin|isavuconazole|antifungal)\b"),
    ("anti_tb_drug", "리팜피신·신규 항결핵제", "drug", r"\b(?:rifampicin|rifampin|isoniazid|bedaquiline|ganfeborole|BTZ-043)\b"),
    ("antiviral", "항바이러스제·LA-ART", "drug", r"\b(?:baloxavir|remdesivir|nirmatrelvir|tenofovir|cabotegravir|dolutegravir|doravirine|rilpivirine)\b"),
    ("vaccine", "백신·면역원성", "drug", r"\b(?:vaccine|vaccination|immunogenicity|monoclonal antibody|nirsevimab)\b"),
    ("phage", "박테리오파지 치료", "drug", r"\b(?:bacteriophage|phage therapy|phage susceptibility)\b"),
    ("asp_short_course", "항균제 스튜어드십·단기치료", "drug", r"\b(?:stewardship|short-course|treatment duration|step-down|antibiogram|WISCA)\b"),
    # Platforms / Diagnostics (장비·진단법)
    ("ngs_wgs", "NGS·mNGS·WGS", "platform", r"\b(?:mNGS|metagenomic|whole genome sequencing|WGS|tNGS|cgMLST|next-generation sequencing)\b"),
    ("rapid_ast", "신속 AST·감수성법", "platform", r"\b(?:rapid AST|rapid susceptibility|disk diffusion|broth microdilution|BMD|Etest|mCIM|AST|EUCAST|CLSI)\b"),
    ("ai_ml", "AI·머신러닝·LLM", "platform", r"\b(?:machine learning|artificial intelligence|AI/ML|deep learning|automated.*microscopy|large language model|LLM)\b"),
    ("ddpcr_poct", "ddPCR·CRISPR·POCT", "platform", r"\b(?:ddPCR|droplet digital|CRISPR|RPA|point-of-care|POCT|lateral flow)\b"),
    ("maldi_ftir", "MALDI-TOF·IR Biotyper", "platform", r"\b(?:MALDI-TOF|IR Biotyper|FT-IR|mass spectrometry)\b"),
    ("multiplex_pcr", "다중 PCR 패널", "platform", r"\b(?:BioFire|FilmArray|QIAstat|multiplex PCR|syndromic panel|Xpert)\b"),
    ("blood_culture", "혈액배양·추적배양", "platform", r"\b(?:blood culture|BCID|follow-up blood culture)\b"),
    ("tdm_pkpd", "약물농도감시(TDM)·PK/PD", "platform", r"\b(?:therapeutic drug monitoring|TDM|pharmacokinetic|PK/PD)\b"),
    ("serology_ag", "항원·항체 신속/면역검사", "platform", r"\b(?:rapid antigen|serology|IgM|IgG|Elecsys|VIDAS|VirClia|galactomannan)\b"),
]

# 비지도 신종 약제/병원체 자동 발굴용 제외 일반어
NOVEL_TERM_STOPWORDS: Set[str] = {
    "antibiotic", "antibiotics", "antimicrobial", "antimicrobials", "bacterial", "clinical",
    "diagnostic", "infection", "infections", "infectious", "microbiology", "molecular",
    "multicenter", "patients", "prospective", "randomized", "resistance", "resistant",
    "retrospective", "surveillance", "susceptibility", "systematic", "treatment", "therapy",
    "hospital", "united", "states", "national", "global", "review", "cohort", "trial",
    "study", "evaluation", "performance", "detection", "diagnosis", "management", "outcomes",
    "children", "adults", "pediatric", "disease", "syndrome", "severe", "acute", "chronic",
}

# 옵션 2: 최상위 리뷰 저널 6종은 서술형 종설·리뷰(78점) 상한(Cap)만 해제 (증례·기초연구 Cap은 유지)
TOP_REVIEW_JOURNALS_NO_REVIEW_CAP: Set[str] = {
    "CMR",
    "NEJM",
    "JAMA",
    "LANCET",
    "LANCET_ID",
}

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

PUBMED_GENERAL_JOURNAL_CLEAN = (
    PUBMED_CLEAN_FILTER
    + ' NOT "Case Reports"[Publication Type] NOT "Biography"[Publication Type]'
    + ' NOT "Historical Article"[Publication Type] NOT "Clinical Conference"[Publication Type]'
    + ' NOT "Interview"[Publication Type] NOT "Personal Narrative"[Publication Type]'
    + ' NOT "Images in Clinical Medicine"[Title] NOT "Clinical Picture"[Title]'
    + ' NOT "Clinical Challenge"[Title] NOT "JAMA Patient Page"[Title]'
    + ' NOT "Case Records of the Massachusetts General Hospital"[Title]'
    + ' NOT "Clinical Problem-Solving"[Title] NOT "Viewpoint"[Title]'
    + ' NOT "Perspective"[Title] NOT "World Report"[Title] NOT "Medical News"[Title]'
    + ' NOT "Vaccines for Cancer"[Title] NOT "Cancer Vaccine*"[Title]'
    + ' NOT "Healthy Dialogue"[Title] NOT "Conversations With"[Title]'
    + ' NOT "What to Know About"[Title] NOT "Why Do Some People"[Title]'
    + ' NOT "Restoring Confidence"[Title] NOT "Vaccine-Hesitant"[Title]'
)

# 종합 의학 저널(NEJM·JAMA·Lancet)용 정밀 조임 0원 사전 필터 (v8.3.4)
# - MeSH는 주주제어([MeSH Major Topic])만 허용해 타과 논문 부차 태그 유입 차단
# - pneumonia·sepsis·vaccine·HIV·간염 등 타과 초록에 흔한 넓은 단어는 [Title]로 한정
# - 내성균·임상미생물 고유 핵심어는 [Title/Abstract] 유지해 최신 미색인 핵심 논문 누락 방지
PUBMED_ID_MICRO_TOPIC_FILTER = (
    ' AND ('
    '"Communicable Diseases"[MeSH Major Topic] OR "Bacterial Infections and Mycoses"[MeSH Major Topic]'
    ' OR "Virus Diseases"[MeSH Major Topic] OR "Parasitic Diseases"[MeSH Major Topic]'
    ' OR "Anti-Infective Agents"[MeSH Major Topic] OR "Drug Resistance, Microbial"[MeSH Major Topic]'
    ' OR "Microbiological Techniques"[MeSH Major Topic] OR "Sepsis"[MeSH Major Topic]'
    ' OR "sepsis"[Title] OR "septic shock"[Title] OR "pneumonia"[Title]'
    ' OR "endocarditis"[Title] OR "meningitis"[Title] OR "tuberculosis"[Title]'
    ' OR "antimicrobial"[Title] OR "antibiotic*"[Title] OR "vaccine*"[Title] OR "vaccination"[Title]'
    ' OR "influenza"[Title] OR "SARS-CoV-2"[Title] OR "COVID-19"[Title]'
    ' OR "respiratory syncytial virus"[Title] OR "RSV"[Title]'
    ' OR "HIV"[Title] OR "hepatitis B"[Title] OR "hepatitis C"[Title]'
    ' OR "malaria"[Title] OR "mpox"[Title] OR "dengue"[Title]'
    ' OR "syphilis"[Title] OR "gonorrhoeae"[Title]'
    ' OR "infectious disease*"[Title] OR "pathogen*"[Title] OR "fungal"[Title]'
    ' OR "Candida"[Title] OR "Enterococcus"[Title] OR "Staphylococcus"[Title]'
    ' OR "Pseudomonas"[Title] OR "Acinetobacter"[Title]'
    ' OR "antimicrobial resistance"[Title/Abstract] OR "antibiotic resistance"[Title/Abstract]'
    ' OR "carbapenem*"[Title/Abstract] OR "beta-lactamase"[Title/Abstract]'
    ' OR "MRSA"[Title/Abstract] OR "VRE"[Title/Abstract] OR "CRE"[Title/Abstract]'
    ' OR "bacteremia"[Title/Abstract] OR "bloodstream infection*"[Title/Abstract]'
    ' OR "antifungal"[Title/Abstract] OR "antiviral"[Title/Abstract]'
    ' OR "Candida auris"[Title/Abstract] OR "Aspergillus"[Title/Abstract]'
    ' OR "Staphylococcus aureus"[Title/Abstract] OR "Enterobacterales"[Title/Abstract]'
    ' OR "Pseudomonas aeruginosa"[Title/Abstract] OR "Acinetobacter baumannii"[Title/Abstract]'
    ' OR "Clostridioides difficile"[Title/Abstract] OR "mycobacteri*"[Title/Abstract]'
    ' OR "blood culture"[Title/Abstract] OR "MALDI-TOF"[Title/Abstract] OR "bacteriophage"[Title/Abstract]'
    ')'
)

# ALM (Annals of Laboratory Medicine)용 임상미생물·감염분자진단·감염면역 정밀 필터 (v8.3.4)
PUBMED_ALM_MICRO_FILTER = (
    ' AND ('
    '"Microbiology"[MeSH Major Topic] OR "Microbiological Techniques"[MeSH Major Topic]'
    ' OR "Bacterial Infections and Mycoses"[MeSH Major Topic] OR "Virus Diseases"[MeSH Major Topic]'
    ' OR "Drug Resistance, Microbial"[MeSH Major Topic]'
    ' OR "microbiolog*"[Title/Abstract] OR "antimicrobial"[Title/Abstract] OR "antibiotic"[Title/Abstract]'
    ' OR "carbapenem*"[Title/Abstract] OR "beta-lactamase"[Title/Abstract] OR "ESBL"[Title/Abstract]'
    ' OR "MRSA"[Title/Abstract] OR "VRE"[Title/Abstract] OR "CRE"[Title/Abstract]'
    ' OR "MALDI-TOF"[Title/Abstract] OR "blood culture"[Title/Abstract] OR "bacteremia"[Title/Abstract]'
    ' OR "bacteriophage"[Title/Abstract] OR "tuberculosis"[Title/Abstract] OR "mycobacteri*"[Title/Abstract]'
    ' OR "fungal"[Title/Abstract] OR "Candida"[Title/Abstract] OR "Aspergillus"[Title/Abstract]'
    ' OR "Clostridioides"[Title/Abstract] OR "Enterobacterales"[Title/Abstract]'
    ' OR "Staphylococcus"[Title/Abstract] OR "Pseudomonas"[Title/Abstract] OR "Acinetobacter"[Title/Abstract]'
    ' OR "susceptibility testing"[Title/Abstract] OR "breakpoint*"[Title/Abstract]'
    ' OR "sepsis"[Title] OR "infection*"[Title] OR "infectious"[Title] OR "pathogen*"[Title]'
    ' OR "virus"[Title] OR "viral"[Title] OR "influenza"[Title] OR "SARS-CoV-2"[Title]'
    ' OR "RSV"[Title] OR "hepatitis"[Title] OR "HIV"[Title] OR "syphilis"[Title] OR "parasit*"[Title]'
    ')'
)

# Nature Reviews Microbiology용 임상·감염·항균제내성·진단 필터 (순수 환경·토양·식물·해양 미생물 제외)
PUBMED_NRM_CLINICAL_FILTER = (
    ' AND ('
    '"Drug Resistance, Microbial"[MeSH Terms] OR "Bacterial Infections and Mycoses"[MeSH Terms]'
    ' OR "Virus Diseases"[MeSH Terms] OR "Anti-Infective Agents"[MeSH Terms]'
    ' OR "antimicrobial"[Title/Abstract] OR "antibiotic"[Title/Abstract] OR "antifungal"[Title/Abstract]'
    ' OR "clinical"[Title/Abstract] OR "human"[Title/Abstract] OR "patient*"[Title/Abstract]'
    ' OR "hospital"[Title/Abstract] OR "nosocomial"[Title/Abstract] OR "diagnostic*"[Title/Abstract]'
    ' OR "sepsis"[Title/Abstract] OR "tuberculosis"[Title/Abstract] OR "vaccine*"[Title/Abstract]'
    ' OR "phage therapy"[Title/Abstract] OR "outbreak*"[Title/Abstract] OR "surveillance"[Title/Abstract]'
    ' OR "carbapenem*"[Title/Abstract] OR "Staphylococcus"[Title/Abstract] OR "Candida"[Title/Abstract]'
    ' OR "Aspergillus"[Title/Abstract] OR "Enterobacterales"[Title/Abstract] OR "Pseudomonas"[Title/Abstract]'
    ' OR "Clostridioides"[Title/Abstract] OR "influenza"[Title/Abstract] OR "coronavirus"[Title/Abstract]'
    ' OR "HIV"[Title/Abstract] OR "malaria"[Title/Abstract] OR "pathogen*"[Title/Abstract]'
    ') NOT ('
    '"soil"[Title] OR "plant"[Title] OR "plants"[Title] OR "ocean"[Title] OR "marine"[Title]'
    ' OR "rhizosphere"[Title] OR "coral"[Title] OR "photosynthesis"[Title] OR "peatland"[Title]'
    ' OR "agriculture"[Title] OR "crop"[Title]'
    ')'
)

# v8.3 신규 7개 저널 정의 (config의 source_v83_new_journals_enabled=TRUE 전까지는 collect 미포함 · test_apis Dry-run 전용)
NEW_JOURNALS_V83: List[Tuple[str, str, str, str]] = [
    ("CMR", "Clinical Microbiology Reviews", '"Clin Microbiol Rev"[Journal]' + PUBMED_CLEAN_FILTER, "전수 수집 · 리뷰Cap해제(옵션2)"),
    ("ALM", "Annals of Laboratory Medicine", '"Ann Lab Med"[Journal]' + PUBMED_ALM_MICRO_FILTER + PUBMED_CLEAN_FILTER, "미생물·감염면역 필터 · 리뷰Cap유지"),
    ("NEJM", "New England Journal of Medicine", '"N Engl J Med"[Journal]' + PUBMED_ID_MICRO_TOPIC_FILTER + PUBMED_GENERAL_JOURNAL_CLEAN, "정밀필터(MeSH주주제·제목/고유균종) · 리뷰Cap해제(옵션2)"),
    ("JAMA", "JAMA", '"JAMA"[Journal]' + PUBMED_ID_MICRO_TOPIC_FILTER + PUBMED_GENERAL_JOURNAL_CLEAN, "정밀필터(MeSH주주제·제목/고유균종) · 리뷰Cap해제(옵션2)"),
    ("LANCET", "The Lancet", '"Lancet"[Journal]' + PUBMED_ID_MICRO_TOPIC_FILTER + PUBMED_GENERAL_JOURNAL_CLEAN, "정밀필터(MeSH주주제·제목/고유균종) · 리뷰Cap해제(옵션2)"),
    ("LANCET_ID", "The Lancet Infectious Diseases", '"Lancet Infect Dis"[Journal]' + PUBMED_GENERAL_JOURNAL_CLEAN, "원저/리뷰/지침 선별 · 리뷰Cap해제(옵션2)"),
]

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
NOTIFIABLE_DISEASES_MASTER_JSON = "[{\"code\": \"\uc81c1\uae09-1\", \"group\": \"1\uae09\", \"name\": \"\uc5d0\ubcfc\ub77c\ubc14\uc774\ub7ec\uc2a4\ubcd1(Ebola virus disease)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-2\", \"group\": \"1\uae09\", \"name\": \"\ub9c8\ubc84\uadf8\uc5f4(Marburg Hemorrhagic fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-3\", \"group\": \"1\uae09\", \"name\": \"\ub77c\uc2f8\uc5f4(Lassa fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-4\", \"group\": \"1\uae09\", \"name\": \"\ud06c\ub9ac\ubbf8\uc548\ucf69\uace0\ucd9c\ud608\uc5f4(Crimean-Congo Hemorrhagic fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-5\", \"group\": \"1\uae09\", \"name\": \"\ub0a8\uc544\uba54\ub9ac\uce74\ucd9c\ud608\uc5f4(South American Hemorrhagic fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-6\", \"group\": \"1\uae09\", \"name\": \"\ub9ac\ud504\ud2b8\ubc38\ub9ac\uc5f4(Rift Valley fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uccb4\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ucd9c\ud608\uc5f4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ubc14\uc774\ub7ec\uc2a4\uc131\ucd9c\ud608\uc5f4) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309745/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-7\", \"group\": \"1\uae09\", \"name\": \"\ub450\ucc3d(Smallpox)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \ud53c\ubd80\ubcd1\ubcc0\uc870\uc9c1, \ud53c\ubd80\ubcd1\ubcc0\uc561, \uac00\ud53c \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ub450\ucc3d,\uc0dd\ubb3c\ud14c\ub7ec\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309744/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-8\", \"group\": \"1\uae09\", \"name\": \"\ud398\uc2a4\ud2b8(Plague)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uac00\ub798, \uae30\uad00\uc9c0\uc138\ucc99\uc561, \ub9bc\ud504\uc808\ud761\uc778\ubb3c \ub4f1)\uc5d0\uc11c Y. pestis \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ud398\uc2a4\ud2b8,\uc0dd\ubb3c\ud14c\ub7ec\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309744/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-9\", \"group\": \"1\uae09\", \"name\": \"\ud0c4\uc800(Anthrax)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uc218\ud3ec\ub3c4\ub9d0, \ub300\ubcc0, \uac00\ub798, \ub1cc\ucc99\uc218\uc561 \ub4f1)\uc5d0\uc11c B. anthracis \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ud0c4\uc800,\uc0dd\ubb3c\ud14c\ub7ec\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309744/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-10\", \"group\": \"1\uae09\", \"name\": \"\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d(Botulism)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0, \uad6c\ud1a0\ubb3c, \uc704\ud761\uc778\ubb3c)\uc5d0\uc11c \ub3c5\uc18c \uc0dd\uc131 C. botulinum \ub4f1 \ubd84\ub9ac \ub3d9\uc815 \u25cb \uac80\uccb4(\ud608\uc561, \ub300\ubcc0, \uad6c\ud1a0\ubb3c, \uc704\ud761\uc778\ubb3c \ub4f1)\uc5d0\uc11c \ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ubcf4\ud234\ub9ac\ub214,\uc0dd\ubb3c\ud14c\ub7ec\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309744/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-11\", \"group\": \"1\uae09\", \"name\": \"\uc57c\ud1a0\ubcd1(Tularemia)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uac00\ub798, \uae30\uad00\uc9c0\uc138\ucc99\uc561, \ub9bc\ud504\uc808\ud761\uc778\ubb3c, \uc870\uc9c1 \ub4f1)\uc5d0\uc11c F. tularensis \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\uc57c\ud1a0\ubcd1,\uc0dd\ubb3c\ud14c\ub7ec\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1(\ub450\ucc3d\u00b7\ud398\uc2a4\ud2b8\u00b7\ud0c4\uc800\u00b7\ubcf4\ud234\ub9ac\ub214\ub3c5\uc18c\uc99d\u00b7\uc57c\ud1a0\ubcd1) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309744/artclView.do\", \"latestDate\": \"2026-01-15\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-12\", \"group\": \"1\uae09\", \"name\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\uc99d\ud6c4\uad70\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\uad6d\ub0b4 \ubbf8\ubc1c\uc0dd \uc2e0\uc885 \ubcd1\uc6d0\uccb4 \uc758\uc2ec \uc99d\ud6c4\uad70(\uae09\uc131\ucd9c\ud608\uc5f4\u00b7\ud638\ud761\uae30\u00b7\uc124\uc0ac\u00b7\ud669\ub2ec\u00b7\uc2e0\uacbd\uc99d\ud6c4\uad70) \uc784\uc0c1\u00b7\uc5ed\ud559 \ud310\uc815 \ubc0f KDCA \ud655\uc778\uac80\uc0ac\", \"dept\": \"\uc704\uae30\uad00\ub9ac\ucd1d\uad04\uacfc, \uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc\", \"guidelineGroup\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1 \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1 \ub300\uc751\uc9c0\uce68(4\uc885: \ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790\u00b7\ucd9c\ud608\uc5f4\u00b7\uc0dd\ubb3c\ud14c\ub7ec\u00b7MERS/SARS)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260126111805686&q_clsfNo=-1\", \"latestDate\": \"2026-01-26\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-13\", \"group\": \"1\uae09\", \"name\": \"\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70 (Severe Acute Respiratory Syndrome, SARS)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70\u00b7\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70(SARS) \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70,\uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70,\uba54\ub974\uc2a4\", \"latestTitle\": \"2026\ub144 \uc81c1\uae09\uac10\uc5fc\ubcd1 \uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70(MERS)\u00b7\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70(SARS) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309847/artclView.do\", \"latestDate\": \"2026-01-22\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-14\", \"group\": \"1\uae09\", \"name\": \"\uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70 (Middle East Respiratory Syndrome: MERS)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \ube44\uac15\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\ud761\uc778\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70\u00b7\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70(SARS) \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70,\uba54\ub974\uc2a4\", \"latestTitle\": \"2026\ub144 \uc81c1\uae09\uac10\uc5fc\ubcd1 \uc911\ub3d9\ud638\ud761\uae30\uc99d\ud6c4\uad70(MERS)\u00b7\uc911\uc99d\uae09\uc131\ud638\ud761\uae30\uc99d\ud6c4\uad70(SARS) \ub300\uc751\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309847/artclView.do\", \"latestDate\": \"2026-01-22\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-15\", \"group\": \"1\uae09\", \"name\": \"\ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790 \uc778\uccb4\uac10\uc5fc\uc99d (Human infection with zoonotic influenza)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561)\uc5d0\uc11c Animal influenza virus \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561, \uc548(\u773c) \uc810\ub9c9 \ub3c4\ub9d0\ubb3c)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790 \uc778\uccb4\uac10\uc5fc\uc99d \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790,\uc870\ub958\uc778\ud50c\ub8e8\uc5d4\uc790\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1 \ub300\uc751\uc9c0\uce68(4\uc885: \ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790\u00b7\ucd9c\ud608\uc5f4\u00b7\uc0dd\ubb3c\ud14c\ub7ec\u00b7MERS/SARS)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260126111805686&q_clsfNo=-1\", \"latestDate\": \"2026-01-26\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-16\", \"group\": \"1\uae09\", \"name\": \"\uc2e0\uc885\uc778\ud50c\ub8e8\uc5d4\uc790(Novel influenza)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561)\uc5d0\uc11c \ubc14\uc774\ub7ec\uc2a4 \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c * \uc0c8\ub85c\uc6b4 \ubc14\uc774\ub7ec\uc2a4\uc758 \uc720\ud589...\", \"dept\": \"\uc704\uae30\uad00\ub9ac\ucd1d\uad04\uacfc, \uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790 \uc778\uccb4\uac10\uc5fc\uc99d \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790,\uc2e0\uc885\uc778\ud50c\ub8e8\uc5d4\uc790\", \"latestTitle\": \"2026\ub144\ub3c4 \uc81c1\uae09\uac10\uc5fc\ubcd1 \ub300\uc751\uc9c0\uce68(4\uc885: \ub3d9\ubb3c\uc778\ud50c\ub8e8\uc5d4\uc790\u00b7\ucd9c\ud608\uc5f4\u00b7\uc0dd\ubb3c\ud14c\ub7ec\u00b7MERS/SARS)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260126111805686&q_clsfNo=-1\", \"latestDate\": \"2026-01-26\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-17\", \"group\": \"1\uae09\", \"name\": \"\ub514\ud504\ud14c\ub9ac\uc544(Diphtheria)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uac15\ub3c4\ub9d0, \uc704\ub9c9\uc870\uc9c1)\uc5d0\uc11c \ub3c5\uc18c\uc0dd\uc131 C. diphtheriae \ubd84\ub9ac \ub3d9\uc815 - \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uac15\ub3c4\ub9d0, \uc704\ub9c9\uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774(\ub3c5\uc18c\uc720\uc804\uc790 \ud3ec\ud568) \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uac15\ub3c4\ub9d0, \uc704\ub9c9\uc870\uc9c1 \ub4f1)\uc5d0\uc11c C. diphtheriae \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c1\uae09-18\", \"group\": \"1\uae09\", \"name\": \"\ub2c8\ud30c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d(Nipah virus infection)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"\uc989\uc2dc \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uc778\ud6c4\ub3c4\ub9d0, \ube44\uac15 \ubc0f \ube44\uc778\ub450 \ud761\uc778\ubb3c, \ub1cc\ucc99\uc218\uc561, \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc81c1\uae09\uac10\uc5fc\ubcd1 \ub2c8\ud30c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\ub2c8\ud30c\ubc14\uc774\ub7ec\uc2a4\", \"latestTitle\": \"2025\ub144 \uc81c1\uae09\uac10\uc5fc\ubcd1 \ub2c8\ud30c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d \ub300\uc751 \uc9c0\uce68(\uc81c1\ud310)\", \"latestYear\": \"2025\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20250908095956243&q_clsfNo=-1\", \"latestDate\": \"2025-09-08\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-1\", \"group\": \"2\uae09\", \"name\": \"\uacb0\ud575(Tuberculosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uac00\ub798, \uae30\uad00\uc9c0\uc138\ucc99\uc561, \uccb4\uc561, \uc870\uc9c1)\uc5d0\uc11c \ud56d\uc0b0\uade0\ub3c4\ub9d0 \uc591\uc131 \u25cb \uac80\uccb4(\uac00\ub798, \uae30\uad00\uc9c0\uc138\ucc99\uc561, \uccb4\uc561, \uc870\uc9c1)\uc5d0\uc11c \uacb0\ud575\uade0* \ubd84\ub9ac \ub3d9\uc815 * \ud2b9\ud788, M. bovis\ub294 \ubc30\uc591\uc5d0\uc11c \ub3d9\uc815\uc774 \ub418\uc5b4\uc57c \ud655\uc9c4 \u25cb \uac80\uccb4(\uac00\ub798, \uae30\uad00\uc9c0\uc138\ucc99\uc561, \uccb4\uc561, \uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uacb0\ud575\uc815\ucc45\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uad6d\uac00\uacb0\ud575\uad00\ub9ac\uc9c0\uce68 \u00b7 \uacb0\ud575\uc9c4\ub8cc\uc9c0\uce68\", \"guidelineKeyword\": \"\uad6d\uac00\uacb0\ud575\uad00\ub9ac\uc9c0\uce68,\uacb0\ud575\", \"latestTitle\": \"2026 \uad6d\uac00\uacb0\ud575\uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309608/artclView.do\", \"latestDate\": \"2026-01-05\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-2\", \"group\": \"2\uae09\", \"name\": \"\uc218\ub450(Varicella)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uc218\ud3ec\uc561, \uac00\ud53c, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \ub1cc\ucc99\uc218\uc561 \ub4f1)\uc5d0\uc11c Human alphaherpesvirus 3 \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\uc218\ud3ec\uc561, \uac00\ud53c, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \ub1cc\ucc99\uc218\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-3\", \"group\": \"2\uae09\", \"name\": \"\ud64d\uc5ed(Measles)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ube44\uac15\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c Measles morbillivirus \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ube44\uac15\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ud64d\uc5ed \ub300\uc751 \uc9c0\uce68\", \"guidelineKeyword\": \"\ud64d\uc5ed \ub300\uc751,\ud64d\uc5ed,\uc608\ubc29\uc811\uc885\ub300\uc0c1\uac10\uc5fc\ubcd1\", \"latestTitle\": \"\ud64d\uc5ed \ub300\uc751 \uc9c0\uce68(2026.5.)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311031/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-4\", \"group\": \"2\uae09\", \"name\": \"\ucf5c\ub808\ub77c(Cholera)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uad6c\ud1a0\ubb3c)\uc5d0\uc11c \ub3c5\uc18c\ud615 \ucf5c\ub808\ub77c\uade0(V. cholerae O1 \ub610\ub294 O139) \ubd84\ub9ac\ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uad6c\ud1a0\ubb3c \ub4f1)\uc5d0\uc11c \ub3c5\uc18c \uc720\ubb34\ub97c \uc54c \uc218 \uc5c6\ub294 V. cholerae O1 \ub610\ub294 O139 \ubd84\ub9ac \ub3d9\uc815 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uad6c\ud1a0\ubb3c \ub4f1)\uc5d0\uc11c \ub3c5\uc18c \uc720\uc804\uc790(ctxA) \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-5\", \"group\": \"2\uae09\", \"name\": \"\uc7a5\ud2f0\ud478\uc2a4(Typhoid fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uc18c\ubcc0, \ub2f4\uc999, \uace8\uc218, \ud608\uc561)\uc5d0\uc11c S. Typhi \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-6\", \"group\": \"2\uae09\", \"name\": \"\ud30c\ub77c\ud2f0\ud478\uc2a4(Paratyphoid fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uc18c\ubcc0, \ub2f4\uc999, \uace8\uc218, \ud608\uc561)\uc5d0\uc11c S. Paratyphi A, B, C \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-7\", \"group\": \"2\uae09\", \"name\": \"\uc138\uade0\uc131\uc774\uc9c8(Shigellosis, Bacillary dysentery)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \uc138\uade0\uc131\uc774\uc9c8\uade0 \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790(ipaH) \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-8\", \"group\": \"2\uae09\", \"name\": \"\uc7a5\ucd9c\ud608\uc131\ub300\uc7a5\uade0(Enterohemorrhagic Escherichia coli) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 \u00ad \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \ub3c5\uc18c \uc720\uc804\uc790(stx1, stx2)\ub97c \ubcf4\uc720\ud55c E. coli \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 \u00ad\uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \ub3c5\uc18c \uc720\uc804\uc790(stx1, stx2) \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-9\", \"group\": \"2\uae09\", \"name\": \"A\ud615\uac04\uc5fc(Viral hepatitis A)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561, \ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68 (A\u00b7B\u00b7C\u00b7E\ud615)\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc,A\ud615\uac04\uc5fc,\uac04\uc5fc\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68(A\ud615\u00b7B\ud615\u00b7C\ud615\u00b7E\ud615, 2026.5. \uc218\uc815\ubc1c\uac04)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309601/artclView.do\", \"latestDate\": \"2026-05-15\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-10\", \"group\": \"2\uae09\", \"name\": \"\ubc31\uc77c\ud574(Pertussis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798)\uc5d0\uc11c B. pertussis \ubd84\ub9ac \ub3d9\uc815 \u25cb \uac80\uccb4(\ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc778\ubb3c, \uac00\ub798)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-11\", \"group\": \"2\uae09\", \"name\": \"\uc720\ud589\uc131\uc774\ud558\uc120\uc5fc(Mumps)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud0c0\uc561, \ubcfc\uc810\ub9c9\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c Mumps orthorubulavirus \ubd84\ub9ac \u25cb\ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb\uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb\uac80\uccb4(\ud0c0\uc561, \ubcfc\uc810\ub9c9\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-12\", \"group\": \"2\uae09\", \"name\": \"\ud48d\uc9c4(Rubella)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u2460 \uc120\ucc9c\uc131 \ud48d\uc9c4 \u25cb \uac80\uccb4(\ube44\uac15\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c Rubivirus rubellae \ubd84\ub9ac \u25cb \ubaa8\uccb4 \ud56d\uccb4\uac00\uac00 \uc5c6\uc5b4\uc9c0\ub294 \uc2dc\uae30(\uc0dd\ud6c4 \uc57d 9\uac1c\uc6d4) \uc774\ud6c4\uc5d0\ub3c4 \ud56d\uccb4 \uc9c0\uc18d \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ube44\uac15\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80...\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-13\", \"group\": \"2\uae09\", \"name\": \"\ud3f4\ub9ac\uc624(Poliomyelitis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0, \ub1cc\ucc99\uc218\uc561, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ud608\uc561)\uc5d0\uc11c Poliovirus \ubd84\ub9ac\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-14\", \"group\": \"2\uae09\", \"name\": \"\uc218\ub9c9\uad6c\uade0 \uac10\uc5fc\uc99d(Meningococcal disease)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub1cc\ucc99\uc218\uc561, \ud608\uc561 \ub4f1)\uc5d0\uc11c N. meningitidis \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-15\", \"group\": \"2\uae09\", \"name\": \"b\ud615\ud5e4\ubaa8\ud544\ub8e8\uc2a4\uc778\ud50c\ub8e8\uc5d4\uc790(Haemophilus influenzae type b)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80\uc870\uc9c1 \ub4f1)\uc5d0\uc11c H. influenzae type b \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80\uc870\uc9c1 \ub4f1)\uc5d0\uc11c H. influenzae \ubd84\ub9ac \ub3d9\uc815 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80...\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-16\", \"group\": \"2\uae09\", \"name\": \"\ud3d0\ub834\uad6c\uade0 \uac10\uc5fc\uc99d(Pneumococcal disease)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80\uc870\uc9c1 \ub4f1)\uc5d0\uc11c S. pneumoniae \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80\uc870\uc9c1 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uad00\uc808\uc561, \ub291\ub9c9\uc561, \uc2ec\ub0ad\uc561, \ubcf5\uc218, \uc0dd\uac80\uc870\uc9c1 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-17\", \"group\": \"2\uae09\", \"name\": \"\ud55c\uc13c\ubcd1(Hansen\u2019s disease, Leprosy)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ubd84\uc790\uc0dd\ubb3c\ud559\uc801 \uac80\uc0ac\ubc95\uc5d0 \uc758\ud55c \ub098\uade0 \ud655\uc778 \u25cb \uc870\uc9c1\uac80\uc0ac\uc0c1 \ud55c\uc13c\ubcd1 \uc721\uc544\uc885 \uc18c\uacac \ud655\uc778 \u25cb \ubcd1\ubcc0\uc758 \ub3c4\ub9d0\uac80\uc0ac(\ud56d\uc0b0\uc131 \uc5fc\uc0c9)\uc5d0\uc11c \ud56d\uc0b0\uc131\uade0 \ud655\uc778\", \"dept\": \"\uacb0\ud575\uc815\ucc45\uacfc\", \"guidelineGroup\": \"\ud55c\uc13c\ubcd1 \uc0ac\uc5c5\uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ud55c\uc13c\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \ud55c\uc13c\ubcd1\uc0ac\uc5c5 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260224093345460&q_clsfNo=-1\", \"latestDate\": \"2026-02-24\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-18\", \"group\": \"2\uae09\", \"name\": \"\uc131\ud64d\uc5f4(Scarlet fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ud608\uc561)\uc5d0\uc11c S. pyogenes \ubd84\ub9ac \ub3d9\uc815 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68(\uc77c\ubd80\uac1c\uc815\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260130085033851&q_clsfNo=-1\", \"latestDate\": \"2026-01-30\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-19\", \"group\": \"2\uae09\", \"name\": \"\ubc18\ucf54\ub9c8\uc774\uc2e0\ub0b4\uc131\ud669\uc0c9\ud3ec\ub3c4\uc54c\uade0(VRSA) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 S. aureus \uc911 Vancomycin MIC 4~8(VISA) \ub610\ub294 \u226516 \u00b5g/mL(VRSA) \ud655\uc778 (\ud608\uc561=\ud658\uc790, \ud608\uc561\uc678=\ubcf4\uc720\uc790)\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-20\", \"group\": \"2\uae09\", \"name\": \"\uce74\ubc14\ud398\ub134\ub0b4\uc131\uc7a5\ub0b4\uc138\uade0\ubaa9(CRE) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 Enterobacterales \uc911 Doripenem/Imipenem/Meropenem MIC \u22654(\u226419mm) \ub610\ub294 Ertapenem MIC \u22652 \u00b5g/mL(\u226418mm) \ub0b4\uc131 \ud655\uc778 (CPE \uc0dd\uc131 \ubcc4\ub3c4 \uc2e0\uace0)\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c2\uae09-21\", \"group\": \"2\uae09\", \"name\": \"E\ud615\uac04\uc5fc(Viral hepatitis E)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561, \ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68 (A\u00b7B\u00b7C\u00b7E\ud615)\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc,E\ud615\uac04\uc5fc,\uac04\uc5fc\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68(A\ud615\u00b7B\ud615\u00b7C\ud615\u00b7E\ud615, 2026.5. \uc218\uc815\ubc1c\uac04)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309601/artclView.do\", \"latestDate\": \"2026-05-15\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-1\", \"group\": \"3\uae09\", \"name\": \"\ud30c\uc0c1\ud48d(Tetanus)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uc9c4\ub2e8 \uae30\uc900 \uc801\uc6a9 (\ud30c\uc0c1\ud48d\uc5d0 \ubd80\ud569\ud558\ub294 \ud2b9\uc9d5\uc801 \uc784\uc0c1\uc99d\uc0c1)\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc608\ubc29\uc811\uc885\", \"latestTitle\": \"2026\ub144\ub3c4 \uc608\ubc29\uc811\uc885 \ub300\uc0c1 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311032/artclView.do\", \"latestDate\": \"2026-05-11\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-2\", \"group\": \"3\uae09\", \"name\": \"B\ud615\uac04\uc5fc(Viral hepatitis B)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0(HBsAg) \ubc0f \ud2b9\uc774 \ud56d\uccb4(IgM anti-HBc) \uac80\ucd9c (\ub2e8, 6\uac1c\uc6d4 \uc804\uc5d0 B\ud615\uac04\uc5fc\uc744 \uc9c4\ub2e8\ubc1b\uc558\ub358 \uc790\ub294 \uc81c\uc678\ud568) \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4(IgM anti-HBc) \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68 (A\u00b7B\u00b7C\u00b7E\ud615)\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc,B\ud615\uac04\uc5fc,\uac04\uc5fc\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68(A\ud615\u00b7B\ud615\u00b7C\ud615\u00b7E\ud615, 2026.5. \uc218\uc815\ubc1c\uac04)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309601/artclView.do\", \"latestDate\": \"2026-05-15\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-3\", \"group\": \"3\uae09\", \"name\": \"\uc77c\ubcf8\ub1cc\uc5fc(Japanese encephalitis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c Japanese encephalitis virus \ubd84\ub9ac - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815\uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-4\", \"group\": \"3\uae09\", \"name\": \"C\ud615\uac04\uc5fc(Viral hepatitis C)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68 \u00b7 C\ud615\uac04\uc5fc \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc,C\ud615\uac04\uc5fc,\uac04\uc5fc\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4 \uac04\uc5fc \uad00\ub9ac\uc9c0\uce68(A\ud615\u00b7B\ud615\u00b7C\ud615\u00b7E\ud615, 2026.5. \uc218\uc815\ubc1c\uac04)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309601/artclView.do\", \"latestDate\": \"2026-05-15\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-5\", \"group\": \"3\uae09\", \"name\": \"\ub9d0\ub77c\ub9ac\uc544(Malaria)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ub3c4\ub9d0\uac80\uc0ac\ub85c \ub9d0\ub77c\ub9ac\uc544 \uc6d0\ucda9(\ucda9\uccb4) \ud655\uc778 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c(\uc2e0\uc18d\uc9c4\ub2e8\uac80\uc0ac \ud3ec\ud568)\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ub9d0\ub77c\ub9ac\uc544 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ub9d0\ub77c\ub9ac\uc544\", \"latestTitle\": \"2026\ub144\ub3c4 \ub9d0\ub77c\ub9ac\uc544 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310433/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-6\", \"group\": \"3\uae09\", \"name\": \"\ub808\uc9c0\uc624\ub12c\ub77c\uc99d(Legionellosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\uae30\uad00\uc9c0\uc138\ucc99\uc561, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc561, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uac00\ub798, \ud3d0\uc870\uc9c1, \ud749\uc218, \ud608\uc561 \ub4f1)\uc5d0\uc11c \ub808\uc9c0\uc624\ub12c\ub77c\uade0 \ubd84\ub9ac \ub3d9\uc815 - \uac80\uccb4(\uc18c\ubcc0)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\uae30\uad00\uc9c0\uc138\ucc99\uc561, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc561, \uae30\uad00\uc9c0\ud761\uc778\ubb3c, \uac00\ub798, \ud3d0\uc870\uc9c1, \ud749\uc218, \ud608\uc561 \ub4f1)\uc5d0\uc11c \uc9c1\uc811\ud615\uad11\ud56d...\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ub808\uc9c0\uc624\ub12c\ub77c\uc99d \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ub808\uc9c0\uc624\ub12c\ub77c\uc99d,\ud638\ud761\uae30\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2025\ub144 \ub808\uc9c0\uc624\ub12c\ub77c\uc99d \uad00\ub9ac\uc9c0\uce68 (2026\ud310 dportal \uc790\ub3d9\uc5f0\ub3d9)\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/228292/artclView.do\", \"latestDate\": \"2025-04-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-7\", \"group\": \"3\uae09\", \"name\": \"\ube44\ube0c\ub9ac\uc624\ud328\ud608\uc99d(Vibrio vulnificus sepsis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \ub300\ubcc0, \uc18c\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \uad6c\ud1a0\ubb3c, \uc218\ud3ec\uc561, \uc870\uc9c1)\uc5d0\uc11c V. vulnificus \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-8\", \"group\": \"3\uae09\", \"name\": \"\ubc1c\uc9c4\ud2f0\ud478\uc2a4(Epidemic typhus)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \ud53c\ubd80\uc870\uc9c1)\uc5d0\uc11c R. prowazekii \ubd84\ub9ac \ub3d9\uc815 \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561, \ud53c\ubd80\uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-9\", \"group\": \"3\uae09\", \"name\": \"\ubc1c\uc9c4\uc5f4(Murine typhus)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \ud53c\ubd80\uc870\uc9c1)\uc5d0\uc11c R. typhi \ubd84\ub9ac \ub3d9\uc815 \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561, \ud53c\ubd80\uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-10\", \"group\": \"3\uae09\", \"name\": \"\ucbd4\ucbd4\uac00\ubb34\uc2dc\uc99d(Scrub typhus)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \uc870\uc9c1, \uac00\ud53c)\uc5d0\uc11c O. tsutsugamushi \ubd84\ub9ac \ub3d9\uc815 - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \uc870\uc9c1, \uac00\ud53c)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-11\", \"group\": \"3\uae09\", \"name\": \"\ub819\ud1a0\uc2a4\ud53c\ub77c\uc99d(Leptospirosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561, \uc870\uc9c1)\uc5d0\uc11c \ub819\ud1a0\uc2a4\ud53c\ub77c\uade0 \ubd84\ub9ac \ub3d9\uc815 - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \ud604\ubbf8\uacbd\uc751\uc9d1\ubc95\uc73c\ub85c \ud608\uccad\uc758 \ub2e8\uc77c\ud56d\uccb4\uac00\uac00 1:800 \uc774\uc0c1 - \uac80\uccb4(\ud608\uc561, \uc18c\ubcc0, \ub1cc\ucc99\uc218\uc561, \uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \ud604\ubbf8\uacbd\uc751\uc9d1\ubc95\uc73c\ub85c \ud608\uccad\uc758 \ub2e8\uc77c\ud56d\uccb4\uac00\uac00 1:200 \uc774\uc0c1 \u223c 1...\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-12\", \"group\": \"3\uae09\", \"name\": \"\ube0c\ub8e8\uc140\ub77c\uc99d(Brucellosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \uace8\uc218, \uad00\uc808\uc561, \uc870\uc9c1 \ub4f1)\uc5d0\uc11c \ube0c\ub8e8\uc140\ub77c\uade0 \ubd84\ub9ac \ub3d9\uc815 - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uae09\uc131\uae30 \ud608\uccad\uc5d0\uc11c \ubbf8\uc138\uc751\uc9d1\ubc95\uc73c\ub85c \ud56d\uccb4\uac00 1:160 \uc774\uc0c1 - \uac80\uccb4(\ud608\uc561, \uace8\uc218, \uad00\uc808\uc561, \uc870\uc9c1 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc778\uc218\uacf5\ud1b5\", \"latestTitle\": \"2026\ub144\ub3c4 \uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1(\ube0c\ub8e8\uc140\ub77c\uc99d, \uacf5\uc218\ubcd1, \ud050\uc5f4, \uc720\ube44\uc800) \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310521/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-13\", \"group\": \"3\uae09\", \"name\": \"\uacf5\uc218\ubcd1(Rabies)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud0c0\uc561)\uc5d0\uc11c Lyssavirus rabies \ubd84\ub9ac \u25cb \uac80\uccb4(\ubaa9 \ud53c\ubd80\uc870\uc9c1, \ub1cc\uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud0c0\uc561, \ubaa9 \ud53c\ubd80\uc870\uc9c1, \ub1cc\uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc778\uc218\uacf5\ud1b5\", \"latestTitle\": \"2026\ub144\ub3c4 \uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1(\ube0c\ub8e8\uc140\ub77c\uc99d, \uacf5\uc218\ubcd1, \ud050\uc5f4, \uc720\ube44\uc800) \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310521/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-14\", \"group\": \"3\uae09\", \"name\": \"\uc2e0\uc99d\ud6c4\uad70\ucd9c\ud608\uc5f4(\uc720\ud589\uc131\ucd9c\ud608\uc5f4) (Hemorrhagic Fever with Renal Syndrome)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c - \uc608\ubc29\uc811\uc885\uc744 \ubc1b\uc9c0 \uc54a\uc740 \uc790 \uc911\uc5d0 \uac04\uc811\uba74\uc5ed\ud615\uad11\ud56d\uccb4\ubc95\uc73c\ub85c \ud56d\uccb4\uac00\uac00 1:512 \uc774\uc0c1 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uc608\ubc29\uc811\uc885\uc744 \ubc1b\uc740 \uc790 \uc911\uc5d0 \uac04\uc811\uba74\uc5ed\ud615\uad11\ud56d\uccb4\ubc95\uc73c\ub85c \ud2b9\uc774 IgG \ud56d\uccb4 \uac80\ucd9c -...\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-15\", \"group\": \"3\uae09\", \"name\": \"\ud6c4\ucc9c\uc131\uba74\uc5ed\uacb0\ud54d\uc99d(AIDS)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uc0dd\ud6c4 18\uac1c\uc6d4 \ubbf8\ub9cc\uc778 \uc790 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c p24 \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c\uc774\uba74\uc11c \ud56d\uc6d0\uc911\ud654\uac80\uc0ac \uc591\uc131 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \uc0dd\ud6c4 18\uac1c\uc6d4 \uc774\uc0c1\uc778 \uc790 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c(\uc6e8\uc2a4\ud134\ube14\ub86f\ubc95\uc73c\ub85c \uc591\uc131\uc778 \uacbd\uc6b0) - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c p24 \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c\uc774\uba74\uc11c \ud56d\uc6d0\uc911\ud654\uac80\uc0ac\uc5d0\uc11c \uc591\uc131 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"HIV/AIDS \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"HIV\", \"latestTitle\": \"2026\ub144 HIV/AIDS \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309970/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-16\", \"group\": \"3\uae09\", \"name\": \"\ud06c\ub85c\uc774\uce20\ud3a0\ud2b8-\uc57c\ucf65\ubcd1(Creutzfeldt-Jakob disease) \ubc0f \ubcc0\uc885\ud06c\ub85c\uc774\uce20\ud3a0\ud2b8-\uc57c\ucf65\ubcd1 (vCJD)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc2e0\uacbd\ubcd1\ub9ac\u00b7\uba74\uc5ed\uc870\uc9c1\ud654\ud559(IHC)\u00b7\ub1cc\ucc99\uc218\uc561 14-3-3/RT-QuIC\u00b7PRNP \uc720\uc804\uc790 \ubcc0\uc774\u00b7\ub1cc\ud30c/MRI \uae30\uc900\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ud06c\ub85c\uc774\uce20\ud3a0\ud2b8-\uc57c\ucf65\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ud06c\ub85c\uc774\uce20\ud3a0\ud2b8,\uc778\uc218\uacf5\ud1b5\", \"latestTitle\": \"2026\ub144 \ud06c\ub85c\uc774\uce20\ud3a0\ud2b8-\uc57c\ucf65\ubcd1(CJD) \uc548\ub0b4\uc11c (\uad00\ub9ac\uc9c0\uce68 2026.03 \uac1c\uc815)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/312872/artclView.do\", \"latestDate\": \"2026-10-07\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-17\", \"group\": \"3\uae09\", \"name\": \"\ud669\uc5f4(Yellow fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778\uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c Yellow fever virus \ubd84\ub9ac - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815\uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1,\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-18\", \"group\": \"3\uae09\", \"name\": \"\ub385\uae30\uc5f4(Dengue fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c Dengue virus \ubd84\ub9ac - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1,\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-19\", \"group\": \"3\uae09\", \"name\": \"\ud050\uc5f4(Q fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \uc870\uc9c1)\uc5d0\uc11c C. burnetii \ubd84\ub9ac\ub3d9\uc815 - \uae09\uc131 \ud050\uc5f4 : \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00(\ud050\uc5f4\uade0 phase \u2161 \ud56d\uc6d0\uc5d0 \ub300\ud55c \ud2b9\uc774\ud56d\uccb4) - \ub9cc\uc131 \ud050\uc5f4 : \uac04\uc811\uba74\uc5ed\ud615\uad11\ud56d\uccb4\ubc95\uc73c\ub85c \uce21\uc815\ud55c phase \u2160 \ud56d\uc6d0\uc5d0 \ub300\ud55c \ud2b9\uc774 IgG \ub2e8\uc77c\ud56d\uccb4\uac00\uac00 1:800 \uc774\uc0c1\uc774\uba74\uc11c phase \u2160\ud56d\uc6d0\uc5d0 \ub300\ud55c ...\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc778\uc218\uacf5\ud1b5\", \"latestTitle\": \"2026\ub144\ub3c4 \uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1(\ube0c\ub8e8\uc140\ub77c\uc99d, \uacf5\uc218\ubcd1, \ud050\uc5f4, \uc720\ube44\uc800) \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310521/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-20\", \"group\": \"3\uae09\", \"name\": \"\uc6e8\uc2a4\ud2b8\ub098\uc77c\uc5f4(West Nile fever)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uc870\uc9c1)\uc5d0\uc11c West Nile virus \ubd84\ub9ac - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \uc870\uc9c1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815\uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1,\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-21\", \"group\": \"3\uae09\", \"name\": \"\ub77c\uc784\ubcd1(Lyme Borreliosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \ud53c\ubd80\uc0dd\uac80\uc870\uc9c1)\uc5d0\uc11c \ubcf4\ub810\ub9ac\uc544\uc18d\uade0 \ubd84\ub9ac \ub3d9\uc815 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \uac04\uc811\uba74\uc5ed\ud615\uad11\ud56d\uccb4\ubc95(\ub610\ub294 ELISA)\uacfc \uc6e8\uc2a4\ud134\ube14\ub7ff\ubc95\uc73c\ub85c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561, \ud53c\ubd80\uc0dd\uac80\uc870\uc9c1)\uc5d0\uc11c \uac04\uc811\uba74\uc5ed\ud615\uad11\ud56d\uccb4\ubc95(\ub610\ub294 ELISA) \ub610\ub294 \uc6e8\uc2a4\ud134\ube14\ub7ff\ubc95\uc73c\ub85c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-22\", \"group\": \"3\uae09\", \"name\": \"\uc9c4\ub4dc\uae30\ub9e4\uac1c\ub1cc\uc5fc(Tick-borne Encephalitis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c Tick-borne encephalitis virus \ubd84\ub9ac - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud574 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-23\", \"group\": \"3\uae09\", \"name\": \"\uc720\ube44\uc800(Melioidosis)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561, \uc18c\ubcc0, \uac00\ub798, \ub18d\uc591 \ub4f1)\uc5d0\uc11c B. pseudomallei \ubd84\ub9ac \ub3d9\uc815\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc778\uc218\uacf5\ud1b5\", \"latestTitle\": \"2026\ub144\ub3c4 \uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1(\ube0c\ub8e8\uc140\ub77c\uc99d, \uacf5\uc218\ubcd1, \ud050\uc5f4, \uc720\ube44\uc800) \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310521/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-24\", \"group\": \"3\uae09\", \"name\": \"\uce58\ucfe4\uad6c\ub2c8\uc57c\uc5f4(Chikungunya fever)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c Chikungunya virus \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud574 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1,\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-25\", \"group\": \"3\uae09\", \"name\": \"\uc911\uc99d\uc5f4\uc131\ud608\uc18c\ud310\uac10\uc18c\uc99d\ud6c4\uad70 (Severe Fever with Thrombocytopenia Syndrome: SFTS)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \uc911\uc99d\uc5f4\uc131\ud608\uc18c\ud310\uac10\uc18c\uc99d\ud6c4\uad70 \ubc14\uc774\ub7ec\uc2a4 \ubd84\ub9ac \u25cb \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958 \ub9e4\uac1c \uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc9c4\ub4dc\uae30\", \"latestTitle\": \"2026\ub144\ub3c4 \uc9c4\ub4dc\uae30\u00b7\uc124\uce58\ub958\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310390/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-26\", \"group\": \"3\uae09\", \"name\": \"\uc9c0\uce74\ubc14\uc774\ub7ec\uc2a4 \uac10\uc5fc\uc99d(Zika virus infection)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \ud655\uc778 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561, \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c Zika virus \ubd84\ub9ac - \ud68c\ubcf5\uae30 \ud608\uccad\uc758 \ud56d\uccb4\uac00\uac00 \uae09\uc131\uae30\uc5d0 \ube44\ud558\uc5ec 4\ubc30 \uc774\uc0c1 \uc99d\uac00 - \uac80\uccb4(\ud608\uc561, \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c \u25cb \ucd94\uc815 \uc9c4\ub2e8 - \uac80\uccb4(\ud608\uc561)\uc5d0\uc11c \ud2b9\uc774 IgM \ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1,\ubaa8\uae30\ub9e4\uac1c\", \"latestTitle\": \"2026\ub144\ub3c4 \ubc14\uc774\ub7ec\uc2a4\uc131 \ubaa8\uae30\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310434/artclView.do\", \"latestDate\": \"2026-03-12\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-27\", \"group\": \"3\uae09\", \"name\": \"\uc5e0\ud3ed\uc2a4(MPOX)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud53c\ubd80\ubcd1\ubcc0\uc561, \ud53c\ubd80\ubcd1\ubcc0\uc870\uc9c1, \uac00\ud53c, \uad6c\uc778\ub450\ub3c4\ub9d0, \ud608\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc5e0\ud3ed\uc2a4 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc5e0\ud3ed\uc2a4\", \"latestTitle\": \"\uc5e0\ud3ed\uc2a4 \uad00\ub9ac\uc9c0\uce68(\uc81c3-1\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260601083147987&q_clsfNo=-1\", \"latestDate\": \"2026-06-01\", \"directVerified\": true}, {\"code\": \"\uc81c3\uae09-28\", \"group\": \"3\uae09\", \"name\": \"\ub9e4\ub3c5(Syphilis)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"24\uc2dc\uac04 \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb 1\uae30\u318d2\uae30\u318d3\uae30 \ub9e4\ub3c5 \ubc0f \uc870\uae30 \uc7a0\ubcf5 \ub9e4\ub3c5 - \uac80\uccb4(\uacbd\uc131\ud558\uac10 \ub610\ub294 \ud3b8\ud3c9\ucf58\ub51c\ub86c\uacfc \uac19\uc740 \ud53c\ubd80\ubcd1\ubcc0\uc870\uc9c1, \uada4\uc591\ubd80\uc704 \uc0bc\ucd9c\uc561(\ub18d, \uc9c4\ubb3c))\uc5d0\uc11c \uc554\uc2dc\uc57c\ud604\ubbf8\uacbd \uac80\uc0ac\ub85c \ub9e4\ub3c5\uade0 \uac80\ucd9c - \uac80\uccb4(\ud608\uc561, \ub1cc\ucc99\uc218\uc561)\uc5d0\uc11c \ud2b8\ub808\ud3ec\ub124\ub9c8\uac80\uc0ac \ubc0f \ube44\ud2b8\ub808\ud3ec\ub124\ub9c8\uac80\uc0ac \ubaa8\ub450 \uc591\uc131 - \uac80\uccb4(\uacbd\uc131\ud558\uac10 \ub610\ub294 \ud3b8\ud3c9\ucf58\ub51c\ub86c\uacfc \uac19\uc740 \ud53c\ubd80\ubcd1\ubcc0\uc870\uc9c1, \uada4\uc591\ubd80\uc704 \uc0bc\ucd9c\uc561(\ub18d, \uc9c4\ubb3c), \ud608\uc561)\uc5d0...\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-1\", \"group\": \"4\uae09\", \"name\": \"\uc778\ud50c\ub8e8\uc5d4\uc790(Influenza)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0(\ub9e4\uc8fc \ud654\uc694\uc77c\uae4c\uc9c0)\", \"lab\": \"\u25cb \uac80\uccb4(\uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ud761\uc785\ubb3c, \ube44\uac15\ud761\uc785\ubb3c, \uae30\uad00\uc9c0\ud3d0\ud3ec\uc138\ucc99\uc561, \uac00\ub798)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68 \u00b7 \uc778\ud50c\ub8e8\uc5d4\uc790 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc778\ud50c\ub8e8\uc5d4\uc790 \uad00\ub9ac\uc9c0\uce68,\ud638\ud761\uae30\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68(\uc77c\ubd80\uac1c\uc815\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260130085033851&q_clsfNo=-1\", \"latestDate\": \"2026-01-30\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-2\", \"group\": \"4\uae09\", \"name\": \"\ud68c\ucda9\uc99d(Ascariasis)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0)\uc5d0\uc11c \ucda9\ub780 \ud655\uc778 \u25cb \uac10\uc5fc\ubd80\uc704\uc5d0\uc11c \ucda9\uccb4 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-3\", \"group\": \"4\uae09\", \"name\": \"\ud3b8\ucda9\uc99d(Trichuriasis)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0)\uc5d0\uc11c \ucda9\ub780 \ud655\uc778 \u25cb \uac10\uc5fc\ubd80\uc704\uc5d0\uc11c \ucda9\uccb4 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-4\", \"group\": \"4\uae09\", \"name\": \"\uc694\ucda9\uc99d(Enterobiasis)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ud56d\ubb38\uc8fc\uc704\ub3c4\ub9d0)\uc5d0\uc11c \ucda9\ub780 \ubc0f \ucda9\uccb4 \ud655\uc778 \u25cb \ud56d\ubb38\uc8fc\uc704\uc640 \uc5ec\uc131\uc758 \uc9c8\uc5d0\uc11c \ucda9\uccb4 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-5\", \"group\": \"4\uae09\", \"name\": \"\uac04\ud761\ucda9\uc99d(Clonorchiasis)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0)\uc5d0\uc11c \ucda9\ub780 \ud655\uc778 \u25cb \uac10\uc5fc\ubd80\uc704\uc5d0\uc11c \ucda9\uccb4 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-6\", \"group\": \"4\uae09\", \"name\": \"\ud3d0\ud761\ucda9\uc99d(Paragonimiasis)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0, \uac00\ub798)\uc5d0\uc11c \ucda9\ub780 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-7\", \"group\": \"4\uae09\", \"name\": \"\uc7a5\ud761\ucda9\uc99d(Intestinal fluke infections)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0)\uc5d0\uc11c \ucda9\ub780 \ubc0f \ucda9\uccb4 \ud655\uc778\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ub9e4\uac1c\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-8\", \"group\": \"4\uae09\", \"name\": \"\uc218\uc871\uad6c\ubcd1(Hand, foot and mouth disease)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \ub1cc\ucc99\uc218\uc561, \ud608\uc561, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uac15\uc138\ucc99\uc561 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\u00b7\uc218\uc871\uad6c\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d\u00b7\uc218\uc871\uad6c\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311285/artclView.do\", \"latestDate\": \"2026-06-01\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-9\", \"group\": \"4\uae09\", \"name\": \"\uc784\uc9c8(Gonorrhea)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub0a8\uc131\uc758 \uc694\ub3c4\ub3c4\ub9d0)\uc5d0\uc11c \uc138\ud3ec \ub0b4 \uadf8\ub78c\uc74c\uc131 \uc30d\uc54c\uade0 \ud604\ubbf8\uacbd \uac80\uc0ac \u25cb \uac80\uccb4(\uc694\ub3c4\u00b7\uc790\uad81\uacbd\ubd80\u00b7\uc9c1\uc7a5\u00b7\uc778\ub450\ub3c4\ub9d0, \uacb0\ub9c9, \ud608\uc561, \uad00\uc808\uc561)\uc5d0\uc11c N. gonorrhoeae \ubd84\ub9ac \ub3d9\uc815 \u25cb \uac80\uccb4(\uc694\ub3c4\u00b7\uc790\uad81\uacbd\ubd80\u00b7\uc9c1\uc7a5\u00b7\uc778\ub450\ub3c4\ub9d0, \uccab \uc18c\ubcc0, \ucc99\uc218\uc561, \uad00\uc808\uc561)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c \u25cb \uac80\uccb4(\uc790\uad81\uacbd\ubd80\u00b7\uc9c8\ub3c4\ub9d0, \uccab \uc18c\ubcc0)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-10\", \"group\": \"4\uae09\", \"name\": \"\ud074\ub77c\ubbf8\ub514\uc544\uac10\uc5fc\uc99d(Chlamydial infection)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uc694\ub3c4\u00b7\uc790\uad81\uacbd\ubd80\u00b7\uc9c1\uc7a5\u00b7\uc778\ub450\ub3c4\ub9d0)\uc5d0\uc11c C. trachomatis \ubd84\ub9ac \ub3d9\uc815 \u25cb \uac80\uccb4(\uc694\ub3c4\u00b7\uc790\uad81\uacbd\ubd80\u00b7\uc9c1\uc7a5\u00b7\uc778\ub450\ub3c4\ub9d0, \uccab \uc18c\ubcc0)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uc6d0 \uac80\ucd9c \u25cb \uac80\uccb4(\uc9c8\ub3c4\ub9d0, \uccab \uc18c\ubcc0 \ub4f1)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-11\", \"group\": \"4\uae09\", \"name\": \"\uc5f0\uc131\ud558\uac10(Chancroid)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ubcd1\ubcc0\uc758 \ubd84\ube44\ubb3c, \uada4\uc591\ubd80\uc704 \uc0bc\ucd9c\ubb3c)\uc5d0\uc11c H. ducreyi \ubd84\ub9ac \ub3d9\uc815 \u25cb \uac80\uccb4(\ubcd1\ubcc0\uc758 \ubd84\ube44\ubb3c, \uada4\uc591\ubd80\uc704 \uc0bc\ucd9c\ubb3c)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-12\", \"group\": \"4\uae09\", \"name\": \"\uc131\uae30\ub2e8\uc21c\ud3ec\uc9c4(Genital herpes)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uc218\ud3ec\ub098 \uada4\uc591 \ubcd1\ubcc0 \ub4f1\uc5d0\uc11c \ub098\uc624\ub294 \ubd84\ube44\ubb3c, \uada4\uc591\ubd80\uc704\ub3c4\ub9d0)\uc5d0\uc11c Simplexvirus humanalpha2 \ub610\ub294 Simplexvirus humanalpha1 \ubd84\ub9ac \u25cb \uac80\uccb4(\ud608\uc561, \uc218\ud3ec\ub098 \uada4\uc591 \ubcd1\ubcc0 \ub4f1\uc5d0\uc11c \ub098\uc624\ub294 \ubd84\ube44\ubb3c, \uada4\uc591\ubd80\uc704\ub3c4\ub9d0)\uc5d0\uc11c \ud2b9\uc774 \ud56d\uccb4 \uac80\ucd9c \u25cb \uac80\uccb4(\ud608\uc561, \uc218\ud3ec\ub098 \uada4\uc591 \ubcd1\ubcc0 \ub4f1\uc5d0\uc11c \ub098\uc624\ub294 \ubd84\ube44\ubb3c, \uada4\uc591\ubd80\uc704\ub3c4\ub9d0)\uc5d0\uc11c ...\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-13\", \"group\": \"4\uae09\", \"name\": \"\ucca8\uaddc\ucf58\ub51c\ub86c(Condyloma acuminata)\", \"scope\": \"\ud658\uc790, \uc758\uc0ac\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ubcd1\ubcc0\uc870\uc9c1, \uc790\uad81\uacbd\ubd80\uc138\ud3ec)\uc5d0\uc11c Human papillomavirus \uac10\uc5fc\uc5d0 \ud569\ub2f9\ud55c \uc870\uc9c1\u00b7\ubcd1\ub9ac\ud559\uc801 \ubcc0\ud654 \ud655\uc778 \u25cb \uac80\uccb4(\ubcd1\ubcc0\uc870\uc9c1, \uc790\uad81\uacbd\ubd80\uc138\ud3ec)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-14\", \"group\": \"4\uae09\", \"name\": \"\ubc18\ucf54\ub9c8\uc774\uc2e0\ub0b4\uc131\uc7a5\uc54c\uade0(VRE) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 Enterococcus \uc911 Vancomycin MIC \u226532 \u00b5g/mL(\u226414mm) \ub0b4\uc131 \ud655\uc778 \ub610\ub294 vanA/vanB \uc720\uc804\uc790 \uac80\ucd9c (\ud608\uc561=\ud658\uc790, \ud608\uc561\uc678=\ubcf4\uc720\uc790)\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-15\", \"group\": \"4\uae09\", \"name\": \"\uba54\ud2f0\uc2e4\ub9b0\ub0b4\uc131\ud669\uc0c9\ud3ec\ub3c4\uc54c\uade0(MRSA) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 S. aureus \uc911 Oxacillin MIC \u22654 \ub610\ub294 Cefoxitin MIC \u22658 \u00b5g/mL(\u226421mm) \ub0b4\uc131 \ud655\uc778 \ub610\ub294 mecA \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-16\", \"group\": \"4\uae09\", \"name\": \"\ub2e4\uc81c\ub0b4\uc131\ub179\ub18d\uade0(MRPA) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 P. aeruginosa \uc911 \uce74\ubc14\ud398\ub134\uacc4\u00b7\uc544\ubbf8\ub178\uae00\ub9ac\ucf54\uc0ac\uc774\ub4dc\uacc4\u00b7\ud50c\ub8e8\uc624\ub85c\ud034\ub180\ub860\uacc4 3\uac1c \uacc4\uc5f4 \ubaa8\ub450 \ub0b4\uc131 \ud655\uc778\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-17\", \"group\": \"4\uae09\", \"name\": \"\ub2e4\uc81c\ub0b4\uc131\uc544\uc2dc\ub124\ud1a0\ubc15\ud130\ubc14\uc6b0\ub9c8\ub2c8\uade0(MRAB) \uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc784\uc0c1\uac80\uccb4 A. baumannii \uc911 \uce74\ubc14\ud398\ub134\uacc4\u00b7\uc544\ubbf8\ub178\uae00\ub9ac\ucf54\uc0ac\uc774\ub4dc\uacc4\u00b7\ud50c\ub8e8\uc624\ub85c\ud034\ub180\ub860\uacc4 3\uac1c \uacc4\uc5f4 \ubaa8\ub450 \ub0b4\uc131 \ud655\uc778\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309968/artclView.do\", \"latestDate\": \"2026-02-02\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-18\", \"group\": \"4\uae09\", \"name\": \"\uc7a5\uad00\uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uc138\uade0 11\uc885(\uc0b4\ubaa8\ub12c\ub77c\u00b7\uc7a5\uc5fc\ube44\ube0c\ub9ac\uc624\u00b7ETEC/EIEC/EPEC\u00b7\uce84\ud544\ub85c\ubc15\ud130 \ub4f1)\u00b7\ubc14\uc774\ub7ec\uc2a4 5\uc885(\ub85c\ud0c0\u00b7\ub178\ub85c \ub4f1)\u00b7\uc6d0\ucda9 4\uc885 \ubd84\ub9ac\ub3d9\uc815 \ub610\ub294 \ud2b9\uc774\uc720\uc804\uc790/\ud56d\uc6d0 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uc7a5\uad00\uac10\uc5fc\uc99d \u00b7 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc7a5\uad00\uac10\uc5fc,\uc218\uc778\uc131\", \"latestTitle\": \"2026\ub144\ub3c4 \uc218\uc778\uc131 \ubc0f \uc2dd\ud488\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/309571/artclView.do\", \"latestDate\": \"2026-01-02\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-19\", \"group\": \"4\uae09\", \"name\": \"\uae09\uc131\ud638\ud761\uae30\uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\ubc14\uc774\ub7ec\uc2a4 7\uc885(\uc544\ub370\ub178\u00b7\ubcf4\uce74\u00b7\ud30c\ub77c\uc778\ud50c\ub8e8\uc5d4\uc790\u00b7RSV\u00b7\ub9ac\ub178\u00b7hMPV\u00b7\uc0ac\ub78c\ucf54\ub85c\ub098) \ubc0f \uc138\uade0 2\uc885(\ub9c8\uc774\ucf54\ud50c\ub77c\uc2a4\ub9c8\u00b7\ud074\ub77c\ubbf8\ub514\uc544 \ud3d0\ub834\uade0) \ubd84\ub9ac \ub610\ub294 \ud2b9\uc774\uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68(\uc77c\ubd80\uac1c\uc815\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260130085033851&q_clsfNo=-1\", \"latestDate\": \"2026-01-30\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-20\", \"group\": \"4\uae09\", \"name\": \"\ud574\uc678\uc720\uc785\uae30\uc0dd\ucda9\uac10\uc5fc\uc99d\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\ud574\uc678\uc720\uc785 \uae30\uc0dd\ucda9 11\uc885(\ub9ac\uc288\ub9cc\ud3b8\ubaa8\ucda9\u00b7\ubc14\ubca0\uc2a4\uc5f4\uc6d0\ucda9\u00b7\uc544\ud504\ub9ac\uce74\uc218\uba74\ubcd1\u00b7\uc8fc\ud608\ud761\ucda9\u00b7\uc0e4\uac00\uc2a4\ubcd1\u00b7\ud3ec\ucda9\u00b7\ud1a1\uc18c\ud3ec\uc790\ucda9 \ub4f1) \ucda9\uccb4\u00b7\ucda9\ub780 \ud655\uc778 \ub610\ub294 \ud2b9\uc774\uc720\uc804\uc790/\ud56d\uccb4 \uac80\ucd9c\", \"dept\": \"\uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uae30\uc0dd\ucda9\", \"latestTitle\": \"2026\ub144\ub3c4 \uae30\uc0dd\ucda9\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310522/artclView.do\", \"latestDate\": \"2026-03-23\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-21\", \"group\": \"4\uae09\", \"name\": \"\uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d(Enterovirus infection)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ub300\ubcc0, \uc9c1\uc7a5\ub3c4\ub9d0, \ub1cc\ucc99\uc218\uc561, \ud608\uc561, \uad6c\uc778\ub450\ub3c4\ub9d0, \ube44\uc778\ub450\ub3c4\ub9d0, \ube44\uac15\uc138\ucc99\uc561)\uc5d0\uc11c \ud2b9\uc774 \uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\u00b7\uc218\uc871\uad6c\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\", \"latestTitle\": \"2026\ub144\ub3c4 \uc5d4\ud14c\ub85c\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d\u00b7\uc218\uc871\uad6c\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311285/artclView.do\", \"latestDate\": \"2026-06-01\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-22\", \"group\": \"4\uae09\", \"name\": \"\uc0ac\ub78c\uc720\ub450\uc885\ubc14\uc774\ub7ec\uc2a4 \uac10\uc5fc\uc99d(Human papillomavirus infection)\", \"scope\": \"\ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\uc790\uad81\uacbd\ubd80 \ub610\ub294 \uc131\uae30 \ubd80\uc704\uc758 \ubcd1\ubcc0\uc870\uc9c1\uc774\ub098 \ub3c4\ub9d0\ubb3c)\uc5d0\uc11c HPV \ud2b9\uc774 \uc720\uc804\ud615 \uac80\ucd9c - HPV \ud2b9\uc774 \uc720\uc804\ud615 : 16, 18, 31, 33, 35, 39, 45, 51, 52, 56, 58, 59, 68 \u203b HPV \ud2b9\uc774 \uc720\uc804\ud615\uc5d0 \ub300\ud55c \uae30\uc900\uc740 IARC \ubd84\ub958 \uae30\uc900(Group 1, Group 2A)\uc5d0 \uadfc\uac70\", \"dept\": \"\uc5d0\uc774\uc988\uad00\ub9ac\uacfc, \ubc14\uc774\ub7ec\uc2a4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \uc131\ub9e4\uac1c\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260303173721749&q_clsfNo=-1\", \"latestDate\": \"2026-03-03\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-23\", \"group\": \"4\uae09\", \"name\": \"\ucf54\ub85c\ub098\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d-19(COVID-19)\", \"scope\": \"\ud658\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\u25cb \uac80\uccb4(\ube44\uc778\ub450\ub3c4\ub9d0, \uad6c\uc778\ub450\ub3c4\ub9d0, \uac00\ub798 \ub4f1)\uc5d0\uc11c \ud2b9\uc774\uc720\uc804\uc790 \uac80\ucd9c\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc, \uc2e0\uc885\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\ud638\ud761\uae30\uac10\uc5fc\ubcd1 \uad00\ub9ac\uc9c0\uce68 \u00b7 \ucf54\ub85c\ub09819 \uad00\ub9ac\uc9c0\uce68\", \"guidelineKeyword\": \"\ucf54\ub85c\ub09819,\ud638\ud761\uae30\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144 \ucf54\ub85c\ub098\ubc14\uc774\ub7ec\uc2a4\uac10\uc5fc\uc99d-19 \uad00\ub9ac\uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/312381/artclView.do\", \"latestDate\": \"2026-08-27\", \"directVerified\": true}, {\"code\": \"\uc81c4\uae09-24\", \"group\": \"4\uae09\", \"name\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uac10\uc5fc\uc99d(Candida auris infection)\", \"scope\": \"\ud658\uc790, \ubcd1\uc6d0\uccb4\ubcf4\uc720\uc790\", \"deadline\": \"7\uc77c \uc774\ub0b4 \uc2e0\uace0\", \"lab\": \"\uac80\uccb4(\uadc0 \uac80\uccb4 \uc81c\uc678, \ud608\uc561\u00b7\uc18c\ubcc0\u00b7\uac10\uc2dc\ubc30\uc591 \ud3ec\ud568)\uc5d0\uc11c Candida auris \ubd84\ub9ac \ub3d9\uc815 (2026.3.29. \uc81c4\uae09 \uc2e0\uaddc \uc9c0\uc815)\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc, \uc138\uade0\ubd84\uc11d\uacfc(\uc9c4\ub2e8)\", \"guidelineGroup\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uad00\ub9ac\uc9c0\uce68 (2026 \uc81c\uc815\ud310)\", \"guidelineKeyword\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4,\uc758\ub8cc\uad00\ub828\uac10\uc5fc\ubcd1\", \"latestTitle\": \"2026\ub144\ub3c4 \uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uad00\ub9ac\uc9c0\uce68 (\uc81c\uc815\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310488/artclView.do\", \"latestDate\": \"2026-03-19\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-1\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\u00b7\uc2e0\uace0\uae30\uc900 (2026.5 \uace0\uc2dc \uc81c2026-4\ud638)\", \"scope\": \"\uc804 \uc758\ub8cc\uae30\uad00 \u00b7 \ubcd1\uc6d0\uccb4 \ud655\uc778\uae30\uad00\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900 (\ub9e4\ub144/\uace0\uc2dc \uac1c\uc815 \uc2dc)\", \"lab\": \"1~4\uae09 \ubc95\uc815\uac10\uc5fc\ubcd1 91\uc885(2026.3.29. \uc81c4\uae09 \uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uac10\uc5fc\uc99d \uc2e0\uc124 \ubc18\uc601) \uc2e0\uace0\ubc94\uc704\u00b7\uc9c4\ub2e8\uac80\uc0ac\uae30\uc900\u00b7\ud45c\ubcf8\uac10\uc2dc \uae30\uc900 \ucd1d\uad04\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\u00b7\uc2e0\uace0\uae30\uc900 \uc9c0\uce68\", \"guidelineKeyword\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8,\uc2e0\uace0\uae30\uc900\", \"latestTitle\": \"2026\ub144 \ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\u00b7\uc2e0\uace0\uae30\uc900 \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311073/artclView.do\", \"latestDate\": \"2026-05-14\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-2\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\uac80\uc0ac \ud1b5\ud569\uc9c0\uce68\", \"scope\": \"\uc9c4\ub2e8\uac80\uc0ac\uc758\ud559\uacfc \u00b7 \ubcf4\uac74\ud658\uacbd\uc5f0\uad6c\uc6d0\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900 (\uac1c\uc815 \uc2dc \uac31\uc2e0)\", \"lab\": \"\ubc95\uc815\uac10\uc5fc\ubcd1\ubcc4 \uac80\uccb4 \ucc44\ucde8\u00b7\ubcf4\uad00\u00b7\uc6b4\uc1a1, \ud655\uc778\uc9c4\ub2e8 \ud45c\uc900\uc2dc\ud5d8\ubc95(\ubc30\uc591\u00b7PCR\u00b7\ud608\uccad\uac80\uc0ac) \ubc0f \ud655\uc778\uae30\uad00 \uc758\ub8b0 \uc808\ucc28\", \"dept\": \"\uc9c4\ub2e8\ubd84\uc11d\uad6d(\uc138\uade0\u00b7\ubc14\uc774\ub7ec\uc2a4\u00b7\uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4\ubd84\uc11d\uacfc)\", \"guidelineGroup\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\uac80\uc0ac \ud1b5\ud569\uc9c0\uce68 (\uc81c5\ud310)\", \"guidelineKeyword\": \"\uc9c4\ub2e8\uac80\uc0ac \ud1b5\ud569\uc9c0\uce68\", \"latestTitle\": \"\ubc95\uc815\uac10\uc5fc\ubcd1 \uc9c4\ub2e8\uac80\uc0ac \ud1b5\ud569\uc9c0\uce68(\uc81c5\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311901/artclView.do\", \"latestDate\": \"2026-07-23\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-3\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc \ud45c\uc900\uc608\ubc29\uc9c0\uce68 \ubc0f \uc9d1\ub2e8\ubc1c\uc0dd \ub300\uc751 \uc548\ub0b4\", \"scope\": \"\uc804 \uc758\ub8cc\uae30\uad00 \u00b7 \uac10\uc5fc\uad00\ub9ac\uc2e4\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900 (\uac1c\uc815 \uc2dc \uac31\uc2e0)\", \"lab\": \"\uc190\uc704\uc0dd\u00b7\ud45c\uc900\uc8fc\uc758\u00b7\uc804\ud30c\uacbd\ub85c\ubcc4 \uc8fc\uc758\u00b7\ud658\uacbd\uc18c\ub3c5\u00b7\uc0bd\uc785\uae30\uad6c \uac10\uc5fc\uc608\ubc29 \ubc0f \ub2e4\uc81c\ub0b4\uc131\uade0(CRE\u00b7C. auris \ub4f1) \uc9d1\ub2e8\ubc1c\uc0dd \ub300\uc751 \ub9e4\ub274\uc5bc\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc \ud45c\uc900\uc608\ubc29\uc9c0\uce68\", \"guidelineKeyword\": \"\ud45c\uc900\uc608\ubc29\uc9c0\uce68\", \"latestTitle\": \"\uc758\ub8cc\uad00\ub828\uac10\uc5fc \ud45c\uc900\uc608\ubc29\uc9c0\uce68 (\ud604\ud589)\", \"latestYear\": \"2017\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/227584/artclView.do\", \"latestDate\": \"2017-07-28\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-4\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4(Candida auris) \uc758\ub8cc\uae30\uad00 \uac10\uc5fc\uad00\ub9ac \uad8c\uace0\uc548\", \"scope\": \"\uc885\ud569\ubcd1\uc6d0 \u00b7 \uac10\uc5fc\uad00\ub9ac\uc2e4 \u00b7 \ubbf8\uc0dd\ubb3c\uac80\uc0ac\uc2e4\", \"deadline\": \"\uc218\uc2dc \uac1c\uc815 (2026\ub144 4\uae09 \uc2e0\uc124)\", \"lab\": \"\uc120\uc81c \uac10\uc2dc\ubc30\uc591\u00b7\uc811\ucd09\uc790 \uac80\uc0ac\u00b7\uaca9\ub9ac\ud574\uc81c \uae30\uc900\u00b7\ud658\uacbd\uc18c\ub3c5 \ubc0f MALDI-TOF/\ubd84\uc790\ub3d9\uc815 \ud655\uc778 \uad8c\uace0\uc0ac\ud56d\", \"dept\": \"\uc758\ub8cc\uac10\uc5fc\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uad00\ub9ac\uc9c0\uce68 (2026 \uc81c\uc815\ud310)\", \"guidelineKeyword\": \"\uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4\", \"latestTitle\": \"2026\ub144\ub3c4 \uce78\ub514\ub2e4 \uc624\ub9ac\uc2a4 \uad00\ub9ac\uc9c0\uce68 (\uc81c\uc815\ud310)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310488/artclView.do\", \"latestDate\": \"2026-03-19\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-5\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uac10\uc5fc\uc131\ubb3c\uc9c8 \uc548\uc804\uc218\uc1a1 \uc9c0\uce68\", \"scope\": \"\uac80\uccb4 \ucc44\ucde8\u00b7\uc218\uc1a1\u00b7\uc758\ub8b0 \uae30\uad00\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900\", \"lab\": \"\uce74\ud14c\uace0\ub9ac A(UN2814)\u00b7\uce74\ud14c\uace0\ub9ac B(UN3373) 3\uc911 \ud3ec\uc7a5 \uae30\uc900, 4\u2103 \ub0c9\uc7a5\u00b7-70\u2103 \ub4dc\ub77c\uc774\uc544\uc774\uc2a4 \uc218\uc1a1 \ubc0f \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4 \uc774\ub3d9 \uc2e0\uace0 \uc808\ucc28\", \"dept\": \"\uc0dd\ubb3c\uc548\uc804\ud3c9\uac00\uacfc\", \"guidelineGroup\": \"\uac10\uc5fc\uc131\ubb3c\uc9c8 \uc548\uc804\uc218\uc1a1 \uc9c0\uce68\", \"guidelineKeyword\": \"\uc548\uc804\uc218\uc1a1\", \"latestTitle\": \"\uac10\uc5fc\uc131\ubb3c\uc9c8 \uc548\uc804\uc218\uc1a1 \uc9c0\uce68(2023)\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/34/207652/artclView.do\", \"latestDate\": \"2025-12-16\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-6\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uc2e4\ud5d8\uc2e4 \uc0dd\ubb3c\uc548\uc804\uc9c0\uce68 \ubc0f \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4 \ucde8\uae09\u00b7\ubcf4\uc874 \uad00\ub9ac\uc9c0\uce68\", \"scope\": \"BL2/BL3 \uac80\uc0ac\uc2e4 \u00b7 \uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4 \ucde8\uae09\uae30\uad00\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900\", \"lab\": \"\uace0\uc704\ud5d8\ubcd1\uc6d0\uccb4(\uc138\uade0\u00b7\uc9c4\uade0 14\uc885, \ubc14\uc774\ub7ec\uc2a4\u00b7\ud504\ub9ac\uc628 24\uc885) \ubd84\ub9ac\uc2e0\uace0\u00b7\uc774\ub3d9\uc2e0\uace0\u00b7\ubcf4\uc874\uad00\ub9ac \ubc0f \uc0dd\ubb3c\uc548\uc804\uc791\uc5c5\ub300(BSC) \uc6b4\uc601 \uae30\uc900\", \"dept\": \"\uc0dd\ubb3c\uc548\uc804\ud3c9\uac00\uacfc\", \"guidelineGroup\": \"\uc2e4\ud5d8\uc2e4 \uc0dd\ubb3c\uc548\uc804\uc9c0\uce68\", \"guidelineKeyword\": \"\uc0dd\ubb3c\uc548\uc804\", \"latestTitle\": \"2025 \uc2e4\ud5d8\uc2e4 \uc0dd\ubb3c\uc548\uc804\uc9c0\uce68(\uc81c4\ud310)\", \"latestYear\": \"2025\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/34/207665/artclView.do\", \"latestDate\": \"2025-12-16\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-7\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uc758\ub8cc\uae30\uad00 \ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uad00\ub9ac(ASP) \uc9c0\uce68\", \"scope\": \"\uc885\ud569\ubcd1\uc6d0 \u00b7 ASP\ud300 \u00b7 \uc9c4\ub2e8\uac80\uc0ac\uc758\ud559\uacfc\", \"deadline\": \"\uc0c1\uc2dc \uae30\uc900\", \"lab\": \"\uc6d0\ub0b4 \ud56d\uc0dd\uc81c \uc0ac\uc6a9\ub7c9(DOT/DDD) \ubaa8\ub2c8\ud130\ub9c1, \ub204\uc801 \ud56d\uc0dd\uc81c \uac10\uc218\uc131(Antibiogram) \uc791\uc131\u00b7\ud658\ub958 \ubc0f \uc120\ud0dd\uc801 \ubcf4\uace0(Cascade reporting) \uc5f0\uacc4\", \"dept\": \"\ud56d\uc0dd\uc81c\ub0b4\uc131\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uc2e4\ubb34\uc9c0\uce68 \ud578\ub4dc\ubd81 \u00b7 \uc9c8\ud658\ubcc4 \uc9c0\uce68\", \"guidelineKeyword\": \"\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uc2e4\ubb34\uc9c0\uce68,\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9\", \"latestTitle\": \"2026\ub144 \ub2e4\ube48\ub3c4 \uac10\uc5fc\uc99d 5\uc885 \ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9\uc744 \uc704\ud55c \uc2e4\ubb34\uc9c0\uce68 (\ud578\ub4dc\ubd81 \ud3ec\ud568)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310916/artclView.do\", \"latestDate\": \"2026-04-29\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-8\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uac10\uc5fc\ubcd1 \uac10\uc2dc\uccb4\uacc4(\uc804\uc218\u00b7\ud45c\ubcf8\uac10\uc2dc) \uc6b4\uc601 \uc9c0\uce68\", \"scope\": \"\uc804 \uc758\ub8cc\uae30\uad00 \u00b7 \ud45c\ubcf8\uac10\uc2dc\uae30\uad00\", \"deadline\": \"\ub9e4\ub144 \uac1c\uc815\", \"lab\": \"\uc804\uc218\uac10\uc2dc(1~3\uae09) \ubc0f \ud45c\ubcf8\uac10\uc2dc(4\uae09 24\uc885\u00b7\ud638\ud761\uae30/\uc7a5\uad00/HAI) \uc2e0\uace0 \uc11c\uc2dd, \ubc29\uc5ed\ud1b5\ud569\uc815\ubcf4\uc2dc\uc2a4\ud15c(eid.kdca.go.kr) \ubcf4\uace0 \uc808\ucc28\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc \u00b7 \uac10\uc5fc\ubcd1\uc815\ubcf4\uacfc\", \"guidelineGroup\": \"\uac10\uc5fc\ubcd1 \uac10\uc2dc\uccb4\uacc4 \uc6b4\uc601 \uc9c0\uce68\", \"guidelineKeyword\": \"\ud45c\ubcf8\uac10\uc2dc \uc6b4\uc601\", \"latestTitle\": \"2026\ub144 \uac10\uc5fc\ubcd1 \uad00\ub9ac \uc0ac\uc5c5 \uc548\ub0b4 \uc9c0\uce68\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/311242/artclView.do\", \"latestDate\": \"2026-05-29\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-9\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uc2e4\ubb34\uc9c0\uce68 (\ud578\ub4dc\ubd81 \ubc0f 5\ub300 \uac10\uc5fc\uc99d: \ud3d0\ub834\u00b7\uc694\ub85c\u00b7\ud53c\ubd80\uc5f0\uc870\uc9c1\u00b7\ubcf5\uac15\ub0b4\u00b7\uae09\uc131\uc0c1\uae30\ub3c4)\", \"scope\": \"\uc885\ud569\ubcd1\uc6d0 \u00b7 \uc784\uc0c1\uc758 \u00b7 ASP\ud300 \u00b7 \uc9c4\ub2e8\uac80\uc0ac\uc758\ud559\uacfc\", \"deadline\": \"2026.04 \uac1c\uc815\", \"lab\": \"\uc9c0\uc5ed\uc0ac\ud68c\ud68d\ub4dd \ud3d0\ub834, \uc694\ub85c\uac10\uc5fc, \ud53c\ubd80\u00b7\uc5f0\uc870\uc9c1 \uac10\uc5fc, \ubcf5\uac15\ub0b4 \uac10\uc5fc, \uae09\uc131 \uc0c1\uae30\ub3c4 \uac10\uc5fc\uc758 \uc6d0\uc778\uade0\u00b7\uc9c4\ub2e8\uac80\uc0ac \ubc0f \uacbd\ud5d8\uc801\u00b7\ud45c\uc801 \ud56d\uc0dd\uc81c \uc2e4\ubb34\uc9c0\uce68\", \"dept\": \"\ud56d\uc0dd\uc81c\ub0b4\uc131\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uc2e4\ubb34\uc9c0\uce68\", \"guidelineKeyword\": \"\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9 \uc2e4\ubb34\uc9c0\uce68,\ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9\", \"latestTitle\": \"2026\ub144 \ub2e4\ube48\ub3c4 \uac10\uc5fc\uc99d 5\uc885 \ud56d\uc0dd\uc81c \uc801\uc815\uc0ac\uc6a9\uc744 \uc704\ud55c \uc2e4\ubb34\uc9c0\uce68 (\ud578\ub4dc\ubd81 \ud3ec\ud568)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/310916/artclView.do\", \"latestDate\": \"2026-04-29\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-10\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uc6d0\uc778\ubd88\uba85 \uc9c8\ubcd1 \uc9d1\ub2e8\uac10\uc5fc \ub300\uc751\uc9c0\uce68 (\uc81c9\ud310)\", \"scope\": \"\uc804 \uc758\ub8cc\uae30\uad00 \u00b7 \uc9c4\ub2e8\uac80\uc0ac\uc758\ud559\uacfc \u00b7 \ubcf4\uac74\uc18c\", \"deadline\": \"\uc218\uc2dc \uac1c\uc815\", \"lab\": \"\uc6d0\uc778 \ubbf8\uc0c1 \ud638\ud761\uae30\u00b7\ucd9c\ud608\uc5f4\u00b7\ubc1c\uc9c4\u00b7\uc2e0\uacbd\u00b7\uc124\uc0ac \uc99d\ud6c4\uad70 \uc9d1\ub2e8\ubc1c\uc0dd \uc2dc \uac80\uccb4 \ucc44\ucde8, \ubc30\uc81c\uc9c4\ub2e8 \ubc0f KDCA \ud655\uc778\uac80\uc0ac \uc758\ub8b0 \uc808\ucc28\", \"dept\": \"\uc2e0\uc885\uac10\uc5fc\ubcd1\ub300\uc751\uacfc \u00b7 \uc9c4\ub2e8\ubd84\uc11d\uad6d\", \"guidelineGroup\": \"\uc6d0\uc778\ubd88\uba85 \uc9c8\ubcd1 \uc9d1\ub2e8\uac10\uc5fc \ub300\uc751\uc9c0\uce68\", \"guidelineKeyword\": \"\uc6d0\uc778\ubd88\uba85 \uc9c8\ubcd1 \uc9d1\ub2e8\uac10\uc5fc,\uc6d0\uc778\ubd88\uba85\", \"latestTitle\": \"\uc6d0\uc778\ubd88\uba85 \uc9c8\ubcd1 \uc9d1\ub2e8\uac10\uc5fc \ub300\uc751\uc9c0\uce68(\uc81c3\ud310)\", \"latestYear\": \"2024\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/228253/artclView.do\", \"latestDate\": \"2024-12-23\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-11\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\uacb0\ud575\uc9c4\ub8cc\uc9c0\uce68 (\uc81c5\ud310) \ubc0f \ub9d0\ub77c\ub9ac\uc544 \uc9c4\ub8cc\uac00\uc774\ub4dc (\uc81c3\ud310)\", \"scope\": \"\uc9c4\ub8cc\uacfc \u00b7 \uc9c4\ub2e8\uac80\uc0ac\uc758\ud559\uacfc\", \"deadline\": \"\uac1c\uc815 \uc2dc \uac31\uc2e0\", \"lab\": \"\uacb0\ud575 \ub3c4\ub9d0\u00b7\ubc30\uc591\u00b7\uc2e0\uc18d\ubd84\uc790\uc9c4\ub2e8(Xpert)\u00b7\uc57d\uc81c\uac10\uc218\uc131\uac80\uc0ac(DST) \ubc0f \ub9d0\ub77c\ub9ac\uc544 \ud608\uc561\ub3c4\ub9d0\u00b7\uc2e0\uc18d\ud56d\uc6d0\u00b7PCR \uc784\uc0c1\uc9c4\ub8cc \uc5f0\uacc4 \uae30\uc900\", \"dept\": \"\uacb0\ud575\uc815\ucc45\uacfc \u00b7 \uc778\uc218\uacf5\ud1b5\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\uacb0\ud575\uc9c4\ub8cc\uc9c0\uce68 \u00b7 \ub9d0\ub77c\ub9ac\uc544 \uc9c4\ub8cc\uac00\uc774\ub4dc\", \"guidelineKeyword\": \"\ub9d0\ub77c\ub9ac\uc544 \uc9c4\ub8cc\uac00\uc774\ub4dc,\uacb0\ud575\uc9c4\ub8cc\uc9c0\uce68\", \"latestTitle\": \"\uacb0\ud575 \uc9c4\ub8cc\uc9c0\uce68(\uc81c5\ud310, 2024)\", \"latestYear\": \"2024\", \"latestUrl\": \"https://www.kdca.go.kr/bbs/kdca/55/228178/artclView.do\", \"latestDate\": \"2024-02-07\", \"directVerified\": true}, {\"code\": \"\uacf5\ud1b5-12\", \"group\": \"\uacf5\ud1b5\u00b7\uac80\uc0ac\uc2e4\", \"name\": \"\ub9cc\uc131 \ucf54\ub85c\ub09819\uc99d\ud6c4\uad70 \uc784\uc0c1\uc9c4\ub8cc\uc9c0\uce68 \uad8c\uace0\uc548\", \"scope\": \"\uc758\ub8cc\uae30\uad00 \u00b7 \ud638\ud761\uae30\u00b7\uac10\uc5fc\ub0b4\uacfc\", \"deadline\": \"2026.02 \ubc1c\uac04\", \"lab\": \"\ucf54\ub85c\ub09819 \uac10\uc5fc \ud6c4 \uc9c0\uc18d \uc99d\uc0c1 \ud3c9\uac00 \ubc0f \uac10\ubcc4\uc9c4\ub2e8 \uad8c\uace0\uc548\", \"dept\": \"\uac10\uc5fc\ubcd1\uad00\ub9ac\uacfc\", \"guidelineGroup\": \"\ub9cc\uc131 \ucf54\ub85c\ub09819\uc99d\ud6c4\uad70 \uc784\uc0c1\uc9c4\ub8cc\uc9c0\uce68\", \"guidelineKeyword\": \"\ub9cc\uc131 \ucf54\ub85c\ub09819\uc99d\ud6c4\uad70\", \"latestTitle\": \"\ub9cc\uc131 \ucf54\ub85c\ub09819\uc99d\ud6c4\uad70 \uc784\uc0c1\uc9c4\ub8cc\uc9c0\uce68 \uad8c\uace0\uc548(2025)\", \"latestYear\": \"2026\", \"latestUrl\": \"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo=20260210094545190&q_clsfNo=-1\", \"latestDate\": \"2026-02-10\", \"directVerified\": true}]"
REFERENCE_STANDARDS_MASTER_JSON = "{\"updatedNote\": \"\uc784\uc0c1\ubbf8\uc0dd\ubb3c\u00b7\uac10\uc5fc\ubd84\uc790\uc9c4\ub2e8 \ud575\uc2ec \ud604\ud589 \ud45c\uc900 \ucd1d 44\uc885 (CLSI M 30\uc885 \u00b7 MM 9\uc885 \u00b7 EUCAST 5\uc885 \u00b7 \uc720\uc804\uc9c8\ud658/\uc554 \ubd84\uc790\ubcd1\ub9ac MM \ubc0f \uc218\uc758\ud559 VET \uc81c\uc678)\", \"clsi_m\": [{\"code\": \"CLSI M100\", \"series\": \"M\", \"title\": \"Performance Standards for Antimicrobial Susceptibility Testing\", \"titleKo\": \"\ud56d\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac(AST) \uc218\ud589 \ud45c\uc900 \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8\u00b7QC \ud45c\", \"edition\": \"36th Ed (M100-Ed36)\", \"publishedDate\": \"2026-01-26\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"\ub9e4\ub144 1~2\uc6d4 \uac1c\uc815. M02\u00b7M07\u00b7M11\uacfc \uc5f0\uacc4 \uc801\uc6a9. Salmonella/Shigella \ud45c \uc7ac\ubc30\uce58 \ub4f1 \ubc18\uc601.\", \"url\": \"https://clsi.org/shop/standards/m100/\"}, {\"code\": \"CLSI M27M44S\", \"series\": \"M\", \"title\": \"Performance Standards for Antifungal Susceptibility Testing of Yeasts\", \"titleKo\": \"\ud6a8\ubaa8\uade0(Candida \ub4f1) \ud56d\uc9c4\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8\u00b7QC \ud45c\", \"edition\": \"4th Ed (M27M44S-Ed4)\", \"publishedDate\": \"2026-03-04\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"M27(\uc561\uccb4\ubc30\uc9c0\ud76c\uc11d\ubc95)\u00b7M44(\ub514\uc2a4\ud06c\ud655\uc0b0\ubc95) \uc5f0\uacc4 \ud6a8\ubaa8\uade0 \ucd5c\uc2e0 \ud310\uc815\ud45c.\", \"url\": \"https://clsi.org/shop/standards/m27m44s/\"}, {\"code\": \"CLSI M38M51S\", \"series\": \"M\", \"title\": \"Performance Standards for Antifungal Susceptibility Testing of Filamentous Fungi\", \"titleKo\": \"\uc0ac\uc0c1\uade0(Aspergillus \ub4f1) \ud56d\uc9c4\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8\u00b7QC \ud45c\", \"edition\": \"4th Ed (M38M51S-Ed4)\", \"publishedDate\": \"2026-03-04\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"M38\u00b7M51 \uc5f0\uacc4 \uc0ac\uc0c1\uade0 MIC/ECV\u00b7QC \ucd5c\uc2e0 \ud310\uc815\ud45c.\", \"url\": \"https://clsi.org/shop/standards/m38m51s/\"}, {\"code\": \"CLSI M02\", \"series\": \"M\", \"title\": \"Performance Standards for Antimicrobial Disk Susceptibility Tests\", \"titleKo\": \"\ud56d\uade0\uc81c \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"14th Ed (M02-Ed14)\", \"publishedDate\": \"2024-03-19\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud638\uae30\uc131 \uc138\uade0 \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ud45c\uc900 \uc808\ucc28 (M100\uacfc \ubcd1\ud589 \uc801\uc6a9).\", \"url\": \"https://clsi.org/shop/standards/m02/\"}, {\"code\": \"CLSI M07\", \"series\": \"M\", \"title\": \"Methods for Dilution Antimicrobial Susceptibility Tests for Bacteria That Grow Aerobically\", \"titleKo\": \"\ud638\uae30\uc131 \uc138\uade0 \uc561\uccb4\u00b7\ud55c\ucc9c\ubc30\uc9c0 \ud76c\uc11d\ubc95(MIC) \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"12th Ed (M07-Ed12)\", \"publishedDate\": \"2024-03-19\", \"status\": \"\ud604\ud589 \uc720\ud6a8 (2025.05 \uc815\uc624\ud45c)\", \"note\": \"\uc561\uccb4\ubc30\uc9c0\ubbf8\ub7c9\ud76c\uc11d\ubc95(BMD) \ud45c\uc900 \uae30\uc900. 2025\ub144 5\uc6d4 \uc815\uc624\ud45c(Correction Notice) \ud655\uc778 \ud544\uc694.\", \"url\": \"https://clsi.org/shop/standards/m07/\"}, {\"code\": \"CLSI M11\", \"series\": \"M\", \"title\": \"Methods for Antimicrobial Susceptibility Testing of Anaerobic Bacteria\", \"titleKo\": \"\ud610\uae30\uc131 \uc138\uade0 \ud56d\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"9th Ed (M11-Ed9)\", \"publishedDate\": \"2018-10-16\", \"status\": \"\ud604\ud589 \uc720\ud6a8 (2025.04 \uc7ac\ud655\uc778)\", \"note\": \"2025\ub144 4\uc6d4 Reaffirmed. Bacteroides/Parabacteroides BMD \ub300\uc548\ubc95 \ud3ec\ud568.\", \"url\": \"https://clsi.org/shop/standards/m11/\"}, {\"code\": \"CLSI M24S\", \"series\": \"M\", \"title\": \"Performance Standards for Susceptibility Testing of Mycobacteria, Nocardia spp., and Other Aerobic Actinomycetes\", \"titleKo\": \"\ud56d\uc0b0\uade0(\uacb0\ud575\u00b7NTM)\u00b7Nocardia \uac10\uc218\uc131 \uac80\uc0ac \ud310\uc815\ud45c\", \"edition\": \"2nd Ed (M24S-Ed2)\", \"publishedDate\": \"2023-02-27\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"M24 \uc5f0\uacc4 \uacb0\ud575\uade0\u00b7\ube44\uacb0\ud575\ud56d\uc0b0\uade0(NTM)\u00b7Nocardia \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \ubc0f QC \ud45c.\", \"url\": \"https://clsi.org/shop/standards/m24s/\"}, {\"code\": \"CLSI M45\", \"series\": \"M\", \"title\": \"Methods for Antimicrobial Dilution and Disk Susceptibility Testing of Infrequently Isolated or Fastidious Bacteria\", \"titleKo\": \"\ub4dc\ubb38 \ubd84\ub9ac\uade0\u00b7\uae4c\ub2e4\ub85c\uc6b4 \uc138\uade0 \uac10\uc218\uc131 \uac80\uc0ac \uc9c0\uce68\", \"edition\": \"3rd Ed (M45-Ed3)\", \"publishedDate\": \"2016-08-15\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"Aeromonas, Bacillus, Campylobacter, Corynebacterium, Pasteurella, Vibrio \ub4f1 \ud2b9\uc218 \uade0\uc885 AST \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m45/\"}, {\"code\": \"CLSI M58\", \"series\": \"M\", \"title\": \"Methods for the Identification of Cultured Microorganisms Using MALDI-TOF MS\", \"titleKo\": \"MALDI-TOF MS\ub97c \uc774\uc6a9\ud55c \ubc30\uc591 \ubbf8\uc0dd\ubb3c \ub3d9\uc815 \uc9c0\uce68\", \"edition\": \"1st Ed (M58-Ed1)\", \"publishedDate\": \"2017-04-26\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"MALDI-TOF MS \ub3d9\uc815 \uac80\uc99d, \uc2a4\ucf54\uc5b4 \ud574\uc11d \ubc0f \uc6b4\uc601\u00b7QC \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m58/\"}, {\"code\": \"CLSI M64\", \"series\": \"M\", \"title\": \"Implementation of Taxonomy Nomenclature Changes\", \"titleKo\": \"\ubbf8\uc0dd\ubb3c \ubd84\ub958\ud559 \uba85\uba85\ubc95 \ubcc0\uacbd\uc758 \uac80\uc0ac\uc2e4 \uc801\uc6a9 \uc9c0\uce68\", \"edition\": \"1st Ed (M64-Ed1)\", \"publishedDate\": \"2024-12-05\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"Candidozyma(Candida) auris \ub4f1 \uac1c\uc815 \uade0\uc885\uba85 LIS \ubc18\uc601\u00b7\uc784\uc0c1 \uacf5\uc9c0 \uac00\uc774\ub4dc.\", \"url\": \"https://clsi.org/shop/standards/m64/\"}, {\"code\": \"CLSI M39\", \"series\": \"M\", \"title\": \"Analysis and Presentation of Cumulative Antimicrobial Susceptibility Test Data\", \"titleKo\": \"\ub204\uc801 \ud56d\uade0\uc81c \uac10\uc218\uc131 \ud1b5\uacc4\ud45c(Antibiogram) \uc791\uc131 \ubc0f \ubcf4\uace0 \uc9c0\uce68\", \"edition\": \"5th Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ubcd1\uc6d0\u00b7\ubcd1\ub3d9\u00b7\uac80\uccb4\ubcc4 \ub204\uc801 \ud56d\uade0\uc81c \uac10\uc218\uc131\ud45c(Antibiogram) \uc791\uc131 \ud1b5\uacc4 \uae30\uc900 \ubc0f ASP \ud658\ub958 \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/m39/\"}, {\"code\": \"CLSI M47\", \"series\": \"M\", \"title\": \"Principles and Procedures for Blood Cultures\", \"titleKo\": \"\ud608\uc561\ubc30\uc591 \uc6d0\uce59 \ubc0f \ud45c\uc900 \uc808\ucc28 \uc9c0\uce68\", \"edition\": \"2nd Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud608\uc561\ubc30\uc591 \ucc44\ud608\ub7c9\u00b7\uc138\ud2b8 \uc218\u00b7\uc624\uc5fc\ub960 \uc9c0\ud45c\u00b7\ubc30\uc591\uae30\uac04\u00b7\uc591\uc131 \ud608\uc561\ubc30\uc591 \uc2e0\uc18d \ub3d9\uc815/\ubcf4\uace0 \ud575\uc2ec \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m47/\"}, {\"code\": \"CLSI M52\", \"series\": \"M\", \"title\": \"Verification of Commercial Microbial Identification and Antimicrobial Susceptibility Testing Systems\", \"titleKo\": \"\uc0c1\uc6a9 \ubbf8\uc0dd\ubb3c \ub3d9\uc815 \ubc0f \ud56d\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \uc2dc\uc2a4\ud15c \uac80\uc99d \uc9c0\uce68\", \"edition\": \"1st Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc2e0\uaddc \uc790\ub3d9\ud654 AST\u00b7\ub3d9\uc815 \uc7a5\ube44 \ubc0f \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \uc5c5\ub370\uc774\ud2b8 \uc2dc \uac80\uc0ac\uc2e4 \uc790\uccb4 \uac80\uc99d(Verification) \ud504\ub85c\ud1a0\ucf5c.\", \"url\": \"https://clsi.org/shop/standards/m52/\"}, {\"code\": \"CLSI M48\", \"series\": \"M\", \"title\": \"Laboratory Detection and Identification of Mycobacteria\", \"titleKo\": \"\ud56d\uc0b0\uade0(\uacb0\ud575\uade0\u00b7\ube44\uacb0\ud575\ud56d\uc0b0\uade0) \uac80\ucd9c \ubc0f \ub3d9\uc815 \uc9c0\uce68\", \"edition\": \"2nd Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud56d\uc0b0\uade0 \uac80\uccb4 \ucc98\ub9ac\u00b7\ub3c4\ub9d0\u00b7\uc561\uccb4/\uace0\uccb4 \ubc30\uc591\u00b7MALDI-TOF \ubc0f \ubd84\uc790\ub3d9\uc815 \ud45c\uc900 \uc808\ucc28.\", \"url\": \"https://clsi.org/shop/standards/m48/\"}, {\"code\": \"CLSI M24\", \"series\": \"M\", \"title\": \"Susceptibility Testing of Mycobacteria, Nocardia spp., and Other Aerobic Actinomycetes\", \"titleKo\": \"\ud56d\uc0b0\uade0\u00b7Nocardia\u00b7\ud638\uae30\uc131 \ubc29\uc120\uade0 \uac10\uc218\uc131 \uac80\uc0ac \ubcf8\ubb38 \ud45c\uc900\", \"edition\": \"3rd Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"M24S \ud310\uc815\ud45c\uc640 \uc5f0\uacc4\ub418\ub294 \uacb0\ud575\u00b7NTM\u00b7Nocardia \uc57d\uc81c\uac10\uc218\uc131 \uc2dc\ud5d8\ubc95 \ubcf8\ubb38.\", \"url\": \"https://clsi.org/shop/standards/m24/\"}, {\"code\": \"CLSI M29\", \"series\": \"M\", \"title\": \"Protection of Laboratory Workers From Occupationally Acquired Infections\", \"titleKo\": \"\uac80\uc0ac\uc2e4 \uc885\uc0ac\uc790\uc758 \uc9c1\uc5c5\uc131 \uac10\uc5fc \uc608\ubc29 \ubc0f \uc0dd\ubb3c\uc548\uc804 \uc9c0\uce68\", \"edition\": \"4th Ed / M29-A4 (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud608\uc561\ub9e4\uac1c\u00b7\uc5d0\uc5b4\ub85c\uc878 \ubcd1\uc6d0\uccb4 \ub178\ucd9c \uc608\ubc29, \uac1c\uc778\ubcf4\ud638\uad6c(PPE) \ubc0f \uc0dd\ubb3c\uc548\uc804\uc0ac\uace0 \ub300\uc751 \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/m29/\"}, {\"code\": \"CLSI M22\", \"series\": \"M\", \"title\": \"Quality Control for Commercially Prepared Microbiological Culture Media\", \"titleKo\": \"\uc0c1\uc6a9 \ubbf8\uc0dd\ubb3c \ubc30\uc591\ubc30\uc9c0 \uc815\ub3c4\uad00\ub9ac(QC) \uc9c0\uce68\", \"edition\": \"3rd Ed / M22-A3 (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc0c1\uc6a9 \ubc30\uc9c0 \uba74\uc81c(Exempt) \ub300\uc0c1 \ubc0f \uac80\uc0ac\uc2e4 \ud544\uc218 \uc131\ub2a5 \uac80\uc99d \ubc30\uc9c0(\uc120\ud0dd\ubc30\uc9c0 \ub4f1) QC \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m22/\"}, {\"code\": \"CLSI M35\", \"series\": \"M\", \"title\": \"Abbreviated Identification of Bacteria and Yeast\", \"titleKo\": \"\uc138\uade0 \ubc0f \ud6a8\ubaa8\uade0\uc758 \uc57d\uc2dd(\uc2e0\uc18d) \ub3d9\uc815 \uc9c0\uce68\", \"edition\": \"2nd Ed / M35-A2 (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"E. coli, S. aureus, P. aeruginosa, C. albicans \ub4f1 \uc8fc\uc694 \uade0\uc885\uc758 \uac04\uc774\u00b7\uc2e0\uc18d \ub3d9\uc815 \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m35/\"}, {\"code\": \"CLSI M40\", \"series\": \"M\", \"title\": \"Quality Control of Microbiological Transport Systems\", \"titleKo\": \"\ubbf8\uc0dd\ubb3c \uac80\uccb4 \uc218\uc1a1\ubc30\uc9c0\u00b7\uc218\uc1a1\uc2dc\uc2a4\ud15c \uc815\ub3c4\uad00\ub9ac \uc9c0\uce68\", \"edition\": \"2nd Ed / M40-A2 (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uba74\ubd09(Swab)\u00b7\uc561\uccb4 \uc218\uc1a1\ubc30\uc9c0\uc758 \ud638\uae30\uc131\u00b7\ud610\uae30\uc131\u00b7\uae4c\ub2e4\ub85c\uc6b4 \uade0\uc885 \uc0dd\uc874\ub825 \ud3c9\uac00 \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m40/\"}, {\"code\": \"CLSI M67\", \"series\": \"M\", \"title\": \"CLSI M67 (Clinical Microbiology Automation Verification Standard)\", \"titleKo\": \"\uc784\uc0c1\ubbf8\uc0dd\ubb3c \uac80\uc0ac\uc2e4 \uc790\ub3d9\ud654 \uc2dc\uc2a4\ud15c \uac80\uc99d \uc9c0\uce68\", \"edition\": \"1st Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\uc2e0\uaddc \ud45c\uc900\", \"note\": \"\ubbf8\uc0dd\ubb3c \ubc30\uc591\u00b7\uc811\uc885\u00b7\ub514\uc9c0\ud138 \ud310\ub3c5 \ub4f1 \uac80\uc0ac\uc2e4 \uc790\ub3d9\ud654(TLA) \ub3c4\uc785 \ubc0f \uac80\uc99d \uc9c0\uce68.\", \"url\": \"https://clsi.org/shop/standards/m67/\"}, {\"code\": \"CLSI M27\", \"series\": \"M\", \"title\": \"Reference Method for Broth Dilution Antifungal Susceptibility Testing of Yeasts\", \"titleKo\": \"\ud6a8\ubaa8\uade0 \ud56d\uc9c4\uade0\uc81c \uc561\uccb4\ubc30\uc9c0 \ud76c\uc11d\ubc95(BMD) \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"4th Ed (M27-Ed4)\", \"publishedDate\": \"2017-11-30\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"Candida \ub4f1 \ud6a8\ubaa8\uade0 \ud56d\uc9c4\uade0\uc81c BMD \ubcf8\ubb38 \ud45c\uc900 (\ud310\uc815\ud45c\ub294 M27M44S \uc5f0\uacc4).\", \"url\": \"https://clsi.org/shop/standards/m27/\"}, {\"code\": \"CLSI M44\", \"series\": \"M\", \"title\": \"Method for Antifungal Disk Diffusion Susceptibility Testing of Yeasts\", \"titleKo\": \"\ud6a8\ubaa8\uade0 \ud56d\uc9c4\uade0\uc81c \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"3rd Ed (M44-Ed3)\", \"publishedDate\": \"2018-12-10\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"Candida spp. \ud50c\ub8e8\ucf54\ub098\uc878\u00b7\ubcf4\ub9ac\ucf54\ub098\uc878\u00b7\uc5d0\ud0a4\ub178\uce78\ub518 \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ubcf8\ubb38 (\ud310\uc815\ud45c\ub294 M27M44S \uc5f0\uacc4).\", \"url\": \"https://clsi.org/shop/standards/m44/\"}, {\"code\": \"CLSI M38\", \"series\": \"M\", \"title\": \"Reference Method for Broth Dilution Antifungal Susceptibility Testing of Filamentous Fungi\", \"titleKo\": \"\uc0ac\uc0c1\uade0 \ud56d\uc9c4\uade0\uc81c \uc561\uccb4\ubc30\uc9c0 \ud76c\uc11d\ubc95(BMD) \ud45c\uc900 \uc2dc\ud5d8\ubc95\", \"edition\": \"3rd Ed (M38-Ed3)\", \"publishedDate\": \"2017-11-30\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"Aspergillus \ub4f1 \uc0ac\uc0c1\uade0 \ud56d\uc9c4\uade0\uc81c BMD \ubcf8\ubb38 \ud45c\uc900 (\ud310\uc815\ud45c\ub294 M38M51S \uc5f0\uacc4).\", \"url\": \"https://clsi.org/shop/standards/m38/\"}, {\"code\": \"CLSI M51\", \"series\": \"M\", \"title\": \"Method for Antifungal Disk Diffusion Susceptibility Testing of Nondermatophyte Filamentous Fungi\", \"titleKo\": \"\ube44\ud53c\ubd80\uc0ac\uc0c1\uade0\uc131 \uc0ac\uc0c1\uade0 \ud56d\uc9c4\uade0\uc81c \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ud45c\uc900\", \"edition\": \"1st Ed (M51-A)\", \"publishedDate\": \"2010-05-25\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc0ac\uc0c1\uade0 \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \uc2dc\ud5d8 \ubcf8\ubb38 (\ud310\uc815\ud45c\ub294 M38M51S \uc5f0\uacc4).\", \"url\": \"https://clsi.org/shop/standards/m51/\"}, {\"code\": \"CLSI M54\", \"series\": \"M\", \"title\": \"Principles and Procedures for Detection and Culture of Fungi in Clinical Specimens\", \"titleKo\": \"\uc784\uc0c1\uac80\uccb4 \uc9c4\uade0 \uac80\ucd9c \ubc0f \ubc30\uc591 \uc6d0\uce59\u00b7\uc808\ucc28 \uc9c0\uce68\", \"edition\": \"2nd Ed (M54-Ed2)\", \"publishedDate\": \"2021-01-06\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc9c4\uade0\uac80\uc0ac\uc2e4 \uc548\uc804\uc218\uce59, \ubc30\uc9c0 \uc120\ud0dd, \ubc30\uc591 \uc870\uac74 \ubc0f \ud310\ub3c5\u00b7\uc758\ub8b0 \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m54/\"}, {\"code\": \"CLSI M57\", \"series\": \"M\", \"title\": \"Principles and Procedures for the Development of Epidemiological Cutoff Values for Antifungal Susceptibility Testing\", \"titleKo\": \"\ud56d\uc9c4\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \uc5ed\ud559\uc801 \uc808\ub2e8\uac12(ECV) \uc218\ub9bd \uc6d0\uce59\u00b7\uc808\ucc28\", \"edition\": \"1st Ed (M57-Ed1)\", \"publishedDate\": \"2016-04-30\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \ubbf8\ud655\ub9bd \uc9c4\uade0\uc758 \uc57c\uc0dd\ud615/\ube44\uc57c\uc0dd\ud615 \uad6c\ubd84 ECV \uac1c\ubc1c \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m57/\"}, {\"code\": \"CLSI M57S\", \"series\": \"M\", \"title\": \"Epidemiological Cutoff Values for Antifungal Susceptibility Testing\", \"titleKo\": \"\ud56d\uc9c4\uade0\uc81c \uac10\uc218\uc131 \uac80\uc0ac \uc5ed\ud559\uc801 \uc808\ub2e8\uac12(ECV) \ubcf4\ucda9\ud45c\", \"edition\": \"\ud604\ud589 \ubcf4\ucda9\ud45c (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"M57 \uc5f0\uacc4 \ud6a8\ubaa8\uade0\u00b7\uc0ac\uc0c1\uade0 \ud56d\uc9c4\uade0\uc81c ECV \uc218\uce58\ud45c.\", \"url\": \"https://clsi.org/shop/standards/m57s/\"}, {\"code\": \"CLSI M15\", \"series\": \"M\", \"title\": \"Laboratory Diagnosis of Blood-borne Parasitic Diseases\", \"titleKo\": \"\ud608\uc561\ub9e4\uac1c \uae30\uc0dd\ucda9 \uc9c8\ud658(\ub9d0\ub77c\ub9ac\uc544 \ub4f1) \uc2e4\ud5d8\uc2e4 \uc9c4\ub2e8 \uc9c0\uce68\", \"edition\": \"1st Ed (M15-A)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ub9d0\ub77c\ub9ac\uc544\u00b7\ubc14\ubca0\uc2dc\uc544 \ub4f1 \ud608\uc561 \ub3c4\ub9d0 \uc81c\uc791\u00b7\uc5fc\uc0c9\u00b7\uac80\uacbd \ubc0f \uc6d0\ucda9 \ubc00\ub3c4 \uc0b0\ucd9c \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/m15/\"}, {\"code\": \"CLSI M50\", \"series\": \"M\", \"title\": \"Quality Control for Commercial Microbial Identification Systems\", \"titleKo\": \"\uc0c1\uc6a9 \ubbf8\uc0dd\ubb3c \ub3d9\uc815 \uc2dc\uc2a4\ud15c \uc815\ub3c4\uad00\ub9ac(QC) \uc9c0\uce68\", \"edition\": \"1st Ed (M50-A)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc0c1\uc6a9 \uc138\uade0\u00b7\uc9c4\uade0 \ub3d9\uc815 \ud328\ub110\uc758 \uac04\uc18c\ud654(Streamlined) QC \ubc0f \uc815\ub3c4\uad00\ub9ac \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/m50/\"}, {\"code\": \"CLSI M23\", \"series\": \"M\", \"title\": \"Development of In Vitro Susceptibility Test Methods, Breakpoints, and Quality Control Parameters\", \"titleKo\": \"\uccb4\uc678 \uac10\uc218\uc131 \uc2dc\ud5d8\ubc95\u00b7\ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8\u00b7QC \ubc94\uc704 \uac1c\ubc1c \uc9c0\uce68 (\ubb34\ub8cc \uacf5\uac1c)\", \"edition\": \"6th Ed (M23-Ed6)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8 (\ubb34\ub8cc)\", \"note\": \"CLSI AST \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8(MIC/PK-PD/\uc784\uc0c1) \uc124\uc815 \uae30\uc900 \ubc0f CLSI-EUCAST \uc870\ud654 \ubb38\uc11c.\", \"url\": \"https://clsi.org/shop/standards/m23/\"}], \"clsi_mm\": [{\"code\": \"CLSI MM06\", \"series\": \"MM\", \"title\": \"Quantitative Molecular Methods for Infectious Diseases\", \"titleKo\": \"\uac10\uc5fc\ubcd1 \uc815\ub7c9 \ubd84\uc790\uc9c4\ub2e8 \uac80\uc0ac \uc9c0\uce68 (HIV\u00b7HBV\u00b7HCV\u00b7CMV \ub4f1)\", \"edition\": \"3rd Ed (MM06-Ed3)\", \"publishedDate\": \"2026-07-22\", \"status\": \"2026 \uc2e0\uaddc \uac1c\uc815\", \"note\": \"2010\ub144 Ed2 \uc774\ud6c4 16\ub144 \ub9cc\uc758 \uc804\uba74 \uac1c\uc815(2026.07.22 \ubc1c\uac04). \ubc14\uc774\ub7ec\uc2a4 \uc815\ub7c9 PCR \uac80\uc99d\u00b7\uc815\ub3c4\uad00\ub9ac \ubc18\uc601.\", \"url\": \"https://clsi.org/shop/standards/mm06/\"}, {\"code\": \"CLSI MM03\", \"series\": \"MM\", \"title\": \"Molecular Diagnostic Methods for Infectious Diseases\", \"titleKo\": \"\uac10\uc5fc\ubcd1 \ubd84\uc790\uc9c4\ub2e8 \uac80\uc0ac \ud45c\uc900 \uc9c0\uce68\", \"edition\": \"\ud604\ud589 \ucd5c\uc2e0\ud310\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uac10\uc5fc\ubcd1 \ud575\uc0b0\uc99d\ud3ed\uac80\uc0ac(NAAT) \uc804\ubc18\uc758 \uc124\uacc4\u00b7\uac80\uc99d\u00b7\uc624\uc5fc\ubc29\uc9c0\u00b7\uc815\ub3c4\uad00\ub9ac \ud575\uc2ec \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm03/\"}, {\"code\": \"CLSI MM17\", \"series\": \"MM\", \"title\": \"Validation and Verification of Multiplex Nucleic Acid Assays\", \"titleKo\": \"\ub2e4\uc911(Multiplex) \ud575\uc0b0\uc99d\ud3ed \ud328\ub110 \uac80\uc0ac \uc131\ub2a5 \uac80\uc99d \uc9c0\uce68\", \"edition\": \"2nd Ed (MM17-Ed2)\", \"publishedDate\": \"2018-05-31\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud638\ud761\uae30\u00b7\uc704\uc7a5\uad00\u00b7\ub1cc\uc218\ub9c9\uc5fc \ub2e4\uc911 PCR \ud328\ub110(FilmArray, QIAstat-Dx \ub4f1) \ubd84\uc11d\uc801 \uac80\uc99d\u00b7QC \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm17/\"}, {\"code\": \"CLSI MM18\", \"series\": \"MM\", \"title\": \"Interpretive Criteria for Identification of Bacteria and Fungi by Targeted DNA Sequencing\", \"titleKo\": \"\ud45c\uc801 DNA \uc5fc\uae30\uc11c\uc5f4\ubd84\uc11d(16S rRNA\u00b7ITS \ub4f1) \uc138\uade0\u00b7\uc9c4\uade0 \ub3d9\uc815 \ud310\ub3c5 \uae30\uc900\", \"edition\": \"2nd Ed (MM18-Ed2)\", \"publishedDate\": \"2018-07-16\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"16S rRNA\u00b7ITS\u00b7D1/D2 \uc2dc\ud000\uc2f1 \ub3d9\uc815 \uc77c\uce58\uc728 \ud310\uc815 \ubc0f \ubcf4\uace0 \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm18/\"}, {\"code\": \"CLSI MM24\", \"series\": \"MM\", \"title\": \"Molecular Methods for Genotyping and Strain Typing of Infectious Organisms\", \"titleKo\": \"\uac10\uc5fc\uc131 \ubcd1\uc6d0\uccb4 \uc720\uc804\ud615 \ubd84\uc11d \ubc0f \uade0\uc8fc \ud615\ubcc4(WGS\u00b7MLST) \uc9c0\uce68\", \"edition\": \"1st Ed (MM24-Ed1)\", \"publishedDate\": \"2021-09-30\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uc6d0\ub0b4 \uc720\ud589 \uc5ed\ud559 \ucd94\uc801\u00b7cgMLST\u00b7\uc720\uc804\ud615 \ubd84\uc11d \ubd84\uc790\uac80\uc0ac \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm24/\"}, {\"code\": \"CLSI MM13\", \"series\": \"MM\", \"title\": \"Collection, Transport, Preparation, and Storage of Specimens for Molecular Methods\", \"titleKo\": \"\ubd84\uc790\uc9c4\ub2e8 \uac80\uccb4 \ucc44\ucde8\u00b7\uc218\uc1a1\u00b7\uc804\ucc98\ub9ac \ubc0f \ubcf4\uad00 \uc9c0\uce68\", \"edition\": \"2nd Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ud575\uc0b0(DNA/RNA) \ubd84\ud574 \ubc0f PCR \uc5b5\uc81c\uc81c \ubc29\uc9c0\ub97c \uc704\ud55c \ubd84\uc790\uc9c4\ub2e8 \uac80\uccb4 \ucde8\uae09 \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm13/\"}, {\"code\": \"CLSI MM19\", \"series\": \"MM\", \"title\": \"Establishing Molecular Testing in Clinical Laboratory Environments\", \"titleKo\": \"\uc784\uc0c1\uac80\uc0ac\uc2e4 \ubd84\uc790\uc9c4\ub2e8 \uac80\uc0ac \ub3c4\uc785 \ubc0f \uc6b4\uc601 \uc9c0\uce68\", \"edition\": \"1st Ed (\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\ubd84\uc790\uc9c4\ub2e8 \uad6c\uc5ed \ubd84\ub9ac(\uc624\uc5fc \ubc29\uc9c0), \uc778\ub825 \uad50\uc721, \uc2dc\uc57d \uac80\uc99d \ubc0f \uc77c\uc0c1 QA \uc6b4\uc601 \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm19/\"}, {\"code\": \"CLSI MM14\", \"series\": \"MM\", \"title\": \"Design of Molecular Proficiency Testing / External Quality Assessment\", \"titleKo\": \"\ubd84\uc790\uc9c4\ub2e8 \uc678\ubd80\uc815\ub3c4\uad00\ub9ac(PT/EQA) \uc124\uacc4 \ubc0f \ud3c9\uac00 \uc9c0\uce68\", \"edition\": \"3rd Ed (MM14-Ed3)\", \"publishedDate\": \"2025-07-17\", \"status\": \"\ud604\ud589 \uc720\ud6a8 (2025 \uac1c\uc815)\", \"note\": \"2025\ub144 7\uc6d4 Ed3 \uac1c\uc815. \ubd84\uc790\uc9c4\ub2e8 \uc678\ubd80\uc815\ub3c4\uad00\ub9ac \uc6b4\uc601 \ubc0f \uacb0\uacfc \ud574\uc11d \uae30\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm14/\"}, {\"code\": \"CLSI MM22\", \"series\": \"MM\", \"title\": \"Microarrays for Diagnosis and Monitoring of Infectious Diseases\", \"titleKo\": \"\uac10\uc5fc\ubcd1 \uc9c4\ub2e8 \ubc0f \ubaa8\ub2c8\ud130\ub9c1\uc6a9 \ub9c8\uc774\ud06c\ub85c\uc5b4\ub808\uc774 \uc9c0\uce68\", \"edition\": \"1st Ed (MM22-Ed1)\", \"publishedDate\": \"\uacf5\uc2dd \ud398\uc774\uc9c0 \ud655\uc778\", \"status\": \"\ud604\ud589 \uc720\ud6a8\", \"note\": \"\uac10\uc5fc\ubcd1 \ub2e4\uc911 \ud45c\uc801 \ub9c8\uc774\ud06c\ub85c\uc5b4\ub808\uc774 \ubd84\uc11d \uac80\uc99d\u00b7QC\u00b7\ubcf4\uace0 \ud45c\uc900.\", \"url\": \"https://clsi.org/shop/standards/mm22/\"}], \"eucast\": [{\"code\": \"EUCAST Bacteria Breakpoints\", \"series\": \"EUCAST\", \"title\": \"Clinical Breakpoint Tables for Bacteria (MICs & Zone Diameters)\", \"titleKo\": \"EUCAST \uc138\uade0 \uc784\uc0c1 \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \ud45c (\ubb34\ub8cc \uc6d0\ubb38 PDF/XLS)\", \"edition\": \"v16.1 (Dosages v16.0)\", \"publishedDate\": \"2026-06-24 ~ 2026-12-31\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"2026\ub144 6\uc6d4 v16.1 \uac1c\uc815 \ubc18\uc601. S / I(Increased exposure) / R \uc815\uc758 \ubc0f ATU \uc8fc\uc758.\", \"url\": \"https://www.eucast.org/bacteria/clinical-breakpoints-and-interpretation/clinical-breakpoint-tables/\"}, {\"code\": \"EUCAST Fungi (AFST)\", \"series\": \"EUCAST\", \"title\": \"Clinical Breakpoints & ECOFFs for Fungi (Yeasts, Moulds, Dermatophytes)\", \"titleKo\": \"EUCAST \uc9c4\uade0 \ud56d\uc9c4\uade0\uc81c \uc784\uc0c1 \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \ubc0f ECOFF \ud45c\", \"edition\": \"Breakpoints v12.1 / ECOFFs v6.1\", \"publishedDate\": \"2026-04-10 / 2026-04-18\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"Candida auris \uc2e0\uaddc \ube0c\ub808\uc774\ud06c\ud3ec\uc778\ud2b8 \ud574\uc11d \uc9c0\uce68 \ubc0f 2026\ub144 Voriconazole/Azole rationale \ud3ec\ud568.\", \"url\": \"https://www.eucast.org/fungi-afst/clinical-breakpoints-and-interpretation/clinical-breakpoint-table/\"}, {\"code\": \"EUCAST QC Tables\", \"series\": \"EUCAST\", \"title\": \"Quality Control Criteria for Disk Diffusion & MIC (Aerobic/Anaerobic)\", \"titleKo\": \"EUCAST \ub0b4\ubd80\uc815\ub3c4\uad00\ub9ac(QC) \uae30\uc900\ud45c\", \"edition\": \"v16.0\", \"publishedDate\": \"2026-01-01\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"\ud638\uae30\uc131\u00b7\ud610\uae30\uc131 \uc138\uade0 ATCC/NCTC QC \uade0\uc8fc \ud5c8\uc6a9 \ubc94\uc704.\", \"url\": \"https://www.eucast.org/bacteria/methodology-and-instructions/disk-diffusion-and-quality-control/\"}, {\"code\": \"EUCAST RAST\", \"series\": \"EUCAST\", \"title\": \"Rapid Antimicrobial Susceptibility Testing (RAST) directly from Positive Blood Cultures\", \"titleKo\": \"\uc591\uc131 \ud608\uc561\ubc30\uc591 \uc9c1\uc811 \uc2e0\uc18d \ub514\uc2a4\ud06c\ud655\uc0b0\ubc95(RAST) \uc9c0\uce68\", \"edition\": \"Method v7.0 / QC v8.0\", \"publishedDate\": \"2026-01-16\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"\uc591\uc131 \ud608\uc561\ubc30\uc591\ubcd1\uc5d0\uc11c 4\u00b76\u00b78\u00b716-20\uc2dc\uac04 \ud310\ub3c5\uc6a9 \uc2e0\uc18d AST \uae30\uc900.\", \"url\": \"https://www.eucast.org/bacteria/rapid-ast-in-blood-culture/methods-and-quality-control/\"}, {\"code\": \"EUCAST Anaerobic Disk\", \"series\": \"EUCAST\", \"title\": \"EUCAST Disk Diffusion Manual for Rapidly Growing Anaerobic Bacteria\", \"titleKo\": \"\uc2e0\uc18d \ubc1c\uc721 \ud610\uae30\uc131 \uc138\uade0 \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ub9e4\ub274\uc5bc\", \"edition\": \"v3.0 (Aerobic Manual v13.0)\", \"publishedDate\": \"2026-06\", \"status\": \"2026 \ucd5c\uc2e0\ud310\", \"note\": \"2026\ub144 6\uc6d4 \ud610\uae30\uc131 \uc138\uade0 \ub514\uc2a4\ud06c \ud655\uc0b0\ubc95 \ub9e4\ub274\uc5bc v3.0 \uac1c\uc815 \ubc18\uc601.\", \"url\": \"https://www.eucast.org/bacteria/methodology-and-instructions/disk-diffusion-and-quality-control/\"}]}"

ROOT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_JSON_PATH = ROOT_DIR / "docs" / "data" / "dashboard.json"

HTTP = requests.Session()
HTTP.headers.update({"User-Agent": "ClinicalMicroBriefing/8.4.3 (Academic Surveillance Bot)"})


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
    source_code: str = "",
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
    review_cap_exempt = False
    # 공식 학회 가이드라인은 유형 Cap 예외(uncapped), 옵션 2 최상위 6개 저널은 서술형 리뷰(78점) Cap만 해제
    if content_type == "guideline":
        cap_val = None
        study_type = "uncapped"
        cap_label = STUDY_TYPE_CAPS["uncapped"][1]
    elif study_type == "narrative_review" and safe_str(source_code).upper() in TOP_REVIEW_JOURNALS_NO_REVIEW_CAP:
        cap_val = None
        review_cap_exempt = True
        cap_label = "최상위 리뷰 Cap 해제(옵션2)"
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
    elif review_cap_exempt:
        breakdown += f" | 유형:{cap_label}"
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
    source_map: Dict[str, str] = {}
    for idx, it in enumerate(chunk, 1):
        raw = str(it.get("rawText") or "").strip()
        has_abs = len(raw) >= 40
        has_abs_map[it["uid"]] = has_abs
        ctype_map[it["uid"]] = it["contentType"]
        source_map[it["uid"]] = it.get("source", "")
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
                item, ctype_map.get(uid, "paper"), has_abs_map.get(uid, False), cutoffs, source_code=source_map.get(uid, "")
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

            pub_types = [
                (pt.text or "").strip()
                for pt in article.findall("./PublicationTypeList/PublicationType")
                if (pt.text or "").strip()
            ]
            results.append(
                {
                    "pmid": pmid,
                    "journal": j_title,
                    "title": title,
                    "abstractText": abstract_text,
                    "pubDate": pub_date,
                    "doi": doi,
                    "pubTypes": ", ".join(pub_types),
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



# =====================================================================
# Official Standards & KDCA Guidelines Live Auto-Detector (0-Won, Fail-Safe)
# =====================================================================
def _parse_version_tuple(v_str: str) -> Tuple[int, ...]:
    nums = re.findall(r"\d+", str(v_str or ""))
    return tuple(int(x) for x in nums[:3]) if nums else ()


def check_eucast_live(eucast_items: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], str, int]:
    """EUCAST 공식 페이지에서 최신 버전 문자열 자동 감지 (실패 시 기존 데이터 안전 보존)"""
    updated = [dict(x) for x in eucast_items]
    ok_cnt = 0
    new_cnt = 0
    checked_urls: Dict[str, str] = {}

    for item in updated:
        url = safe_str(item.get("url"))
        if not url:
            continue
        try:
            if url not in checked_urls:
                res = HTTP.get(url, timeout=(10, 20))
                if res.status_code == 200:
                    checked_urls[url] = strip_html(res.text)
                else:
                    checked_urls[url] = ""
                time.sleep(0.2)
            body = checked_urls[url]
            if not body:
                item["liveStatus"] = "접속 보류 (검수본 유지)"
                continue
            ok_cnt += 1
            # Extract highest plausible version number on page (e.g., v16.1, v12.1, v7.0)
            code = item.get("code", "")
            stored_ed = safe_str(item.get("edition"))
            stored_vs = re.findall(r"v\s*(\d+\.\d+)", stored_ed, re.IGNORECASE)
            found_vs = re.findall(r"(?:v\.?|version\s*)(\d{1,2}\.\d{1,2})", body, re.IGNORECASE)
            # Filter plausible range near stored major version so unrelated numbers aren't picked
            best_v = ""
            if stored_vs and found_vs:
                base_major = int(stored_vs[0].split(".")[0])
                plausible = [v for v in found_vs if base_major <= int(v.split(".")[0]) <= base_major + 2]
                if plausible:
                    best_v = max(plausible, key=_parse_version_tuple)
            if best_v and stored_vs:
                if _parse_version_tuple(best_v) > _parse_version_tuple(stored_vs[0]):
                    item["liveStatus"] = f"신규 개정 감지 (v{best_v})"
                    item["liveDetected"] = f"v{best_v}"
                    new_cnt += 1
                else:
                    item["liveStatus"] = f"자동 점검 일치 (v{best_v})"
            else:
                item["liveStatus"] = "페이지 정상 응답 (수동 판본 확인)"
        except Exception:
            item["liveStatus"] = "접속 보류 (검수본 유지)"

    summary = f"EUCAST {ok_cnt}/{len(updated)} 페이지 정상 (신규감지 {new_cnt}건)"
    return updated, summary, new_cnt


def check_clsi_live(clsi_m: List[Dict[str, Any]], clsi_mm: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, int]:
    """CLSI 공식 페이지(clsi.org) 접속 가능 여부 확인 및 판본 자동 감지 (403 차단 시 즉시 안전 종료)"""
    out_m = [dict(x) for x in clsi_m]
    out_mm = [dict(x) for x in clsi_mm]
    probe_url = "https://clsi.org/shop/standards/m100/"
    try:
        res = HTTP.get(probe_url, timeout=(10, 20))
        if res.status_code != 200:
            msg = f"CLSI 자동 접속 제한(HTTP {res.status_code}) — 검수본 안전 유지"
            for it in out_m + out_mm:
                it["liveStatus"] = "공식 페이지 링크 제공"
            return out_m, out_mm, msg, 0
    except Exception as e:
        msg = f"CLSI 접속 보류({str(e)[:60]}) — 검수본 안전 유지"
        for it in out_m + out_mm:
            it["liveStatus"] = "공식 페이지 링크 제공"
        return out_m, out_mm, msg, 0

    ok_cnt = 0
    new_cnt = 0
    for it in out_m + out_mm:
        url = safe_str(it.get("url"))
        if not url:
            continue
        try:
            r = HTTP.get(url, timeout=(10, 20))
            if r.status_code != 200:
                it["liveStatus"] = f"조회 제한(HTTP {r.status_code})"
                continue
            ok_cnt += 1
            txt = strip_html(r.text)[:8000]
            # Look for Edition pattern like "36th Edition" or "M100-Ed36"
            ed_m = re.search(r"\b(\d{1,2})(?:st|nd|rd|th)\s+Edition\b", txt, re.IGNORECASE) or re.search(
                r"\bEd(?:ition)?\s*(\d{1,2})\b", txt, re.IGNORECASE
            )
            date_m = re.search(
                r"(?:Published|Publication Date|Effective)\s*[:\-]?\s*([A-Z][a-z]+\s+\d{1,2},?\s+20\d{2}|20\d{2}-\d{2}-\d{2})",
                txt,
                re.IGNORECASE,
            )
            if ed_m:
                live_ed_num = int(ed_m.group(1))
                stored_ed_m = re.search(r"(\d{1,2})(?:st|nd|rd|th)|Ed(\d{1,2})|A(\d{1,2})", safe_str(it.get("edition")), re.IGNORECASE)
                stored_ed_num = int(next(g for g in stored_ed_m.groups() if g)) if stored_ed_m else 0
                if "공식 페이지 확인" in safe_str(it.get("edition")):
                    it["edition"] = f"{live_ed_num}th Ed (자동 확인)"
                if date_m and "공식 페이지 확인" in safe_str(it.get("publishedDate")):
                    it["publishedDate"] = date_m.group(1)
                if stored_ed_num and live_ed_num > stored_ed_num:
                    it["liveStatus"] = f"신규 개정 감지 (Ed{live_ed_num})"
                    it["liveDetected"] = f"Ed{live_ed_num}"
                    new_cnt += 1
                else:
                    it["liveStatus"] = f"자동 점검 일치 (Ed{live_ed_num})"
            else:
                it["liveStatus"] = "페이지 정상 응답"
            time.sleep(0.2)
        except Exception:
            it["liveStatus"] = "접속 보류 (검수본 유지)"

    summary = f"CLSI {ok_cnt}/{len(out_m) + len(out_mm)} 페이지 정상 (신규감지 {new_cnt}건)"
    return out_m, out_mm, summary, new_cnt


def _fetch_dportal_guidelines_1001(max_pages: int = 15) -> List[Dict[str, str]]:
    """
    질병관리청 감염병포털 > 감염병지침 게시판 (dportal.kdca.go.kr, q_bbsSn=1001) 실시간 크롤링.
    - 표본감시 주간소식지(q_bbsSn=1010)와 동일한 정적 HTML 서버 엔진 사용.
    - q_bbsDocNo(YYYYMMDDHHMMSS...)에서 정확한 등록일(YYYY-MM-DD)과 개별 게시글 직링크를 추출하며,
      최신순(q_bbsDocNo 내림차순) 정렬로 구버전 지침 매칭을 원천 차단함.
    """
    posts: List[Dict[str, str]] = []
    seen_docs: Set[str] = set()

    for page in range(1, max_pages + 1):
        url = f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbsList.do?q_bbsSn=1001&q_currPage={page}"
        res = HTTP.get(url, timeout=(12, 25))
        if res.status_code != 200:
            if page == 1:
                raise RuntimeError(f"dportal 감염병지침 HTTP {res.status_code}")
            break

        html_body = res.text
        added = 0

        # Pattern 1: href에 q_bbsDocNo=...가 직접 들어있는 링크
        for m in re.finditer(r'<a[^>]+href="[^"]*q_bbsDocNo=(\d{12,20})[^"]*"[^>]*>([\s\S]{3,350}?)</a>', html_body, re.IGNORECASE):
            doc_no = m.group(1)
            if doc_no in seen_docs:
                continue
            title = re.sub(r"\s*(?:새글|NEW|첨부파일)\s*$", "", strip_html(m.group(2)), flags=re.IGNORECASE).strip()
            if len(title) < 4:
                continue
            seen_docs.add(doc_no)
            date_str = f"{doc_no[:4]}-{doc_no[4:6]}-{doc_no[6:8]}" if re.match(r"^20[12]\d[01]\d[0-3]\d", doc_no) else ""
            yr_m = re.search(r"(20[12]\d)", title)
            yr = date_str[:4] if date_str else (yr_m.group(1) if yr_m else "2026")
            posts.append({
                "docNo": doc_no,
                "title": title,
                "pubDate": date_str,
                "year": yr,
                "url": f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo={doc_no}&q_clsfNo=-1",
                "source": "KDCA 감염병포털 지침",
            })
            added += 1

        # Pattern 2: onclick="...('20260723171016500')..." 형식의 링크
        if added == 0:
            for m in re.finditer(r'<a[^>]+onclick="[^"]*?(20[12]\d{12,17})[^"]*"[^>]*>([\s\S]{3,350}?)</a>', html_body, re.IGNORECASE):
                doc_no = m.group(1)
                if doc_no in seen_docs:
                    continue
                title = re.sub(r"\s*(?:새글|NEW|첨부파일)\s*$", "", strip_html(m.group(2)), flags=re.IGNORECASE).strip()
                if len(title) < 4:
                    continue
                seen_docs.add(doc_no)
                date_str = f"{doc_no[:4]}-{doc_no[4:6]}-{doc_no[6:8]}"
                posts.append({
                    "docNo": doc_no,
                    "title": title,
                    "pubDate": date_str,
                    "year": date_str[:4],
                    "url": f"https://dportal.kdca.go.kr/pot/bbs/BD_selectBbs.do?q_bbsSn=1001&q_bbsDocNo={doc_no}&q_clsfNo=-1",
                    "source": "KDCA 감염병포털 지침",
                })
                added += 1

        if added == 0:
            break
        # 2023년 이전 페이지까지 내려갔으면 중단 (구버전 불필요)
        oldest_on_page = min((p["pubDate"] for p in posts[-added:] if p["pubDate"]), default="2099-01-01")
        if oldest_on_page < "2023-01-01":
            break
        time.sleep(0.18)

    # 최신 문서 최우선 정렬 (docNo/pubDate 내림차순)
    posts.sort(key=lambda p: (p.get("pubDate", ""), p.get("docNo", "")), reverse=True)
    return posts


def check_kdca_guidelines_live(notifiable_list: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, int]:
    """
    KDCA 감염병포털(dportal q_bbsSn=1001) 실시간 크롤링으로:
    1) 법정감염병 103건의 최신 세부지침 직링크 자동 갱신 (구버전 역행 차단 가드 적용)
    2) 탭 1 주간 브리핑용 2026년 KDCA 지침 출간·개정 소식 목록(kdcaGuidelineReleases) 생성
    """
    updated = [dict(x) for x in notifiable_list]
    live_posts: List[Dict[str, str]] = []
    live_err = ""
    try:
        live_posts = _fetch_dportal_guidelines_1001(max_pages=15)
    except Exception as e:
        live_err = str(e)[:80]

    def norm_ko(s: str) -> str:
        return re.sub(r"[\s·ㆍ\-_()\[\],.]", "", str(s or "")).lower()

    def title_year_of(s: str, fallback_date: str = "") -> int:
        m = re.search(r"(20[12]\d)", str(s or ""))
        if m:
            return int(m.group(1))
        if fallback_date and len(fallback_date) >= 4 and fallback_date[:4].isdigit():
            return int(fallback_date[:4])
        return 0

    matched_live = 0
    if live_posts:
        # 정식 지침/가이드/권고안 우선, 단순 정오표·안내문은 후순위
        def is_main_guide(t: str) -> bool:
            return any(k in t for k in ("지침", "가이드", "권고안", "매뉴얼", "기준", "안내서"))

        prioritized_posts = sorted(
            live_posts,
            key=lambda p: (
                1 if is_main_guide(p["title"]) else 0,
                p.get("pubDate", ""),
                p.get("docNo", ""),
            ),
            reverse=True,
        )
        # Re-sort primarily by pubDate desc so a newer 2026 guide always beats an older 2024 guide!
        prioritized_posts.sort(
            key=lambda p: (
                title_year_of(p["title"], p.get("pubDate", "")),
                p.get("pubDate", ""),
                1 if is_main_guide(p["title"]) else 0,
                p.get("docNo", ""),
            ),
            reverse=True,
        )

        for d in updated:
            raw_kw = safe_str(d.get("guidelineKeyword"))
            if not raw_kw:
                continue
            kw_candidates = [k.strip() for k in raw_kw.split(",") if k.strip()]
            hit = None
            for cand_kw in kw_candidates:
                cand_tokens = [norm_ko(t) for t in re.split(r"\s+", cand_kw) if len(norm_ko(t)) >= 2]
                if not cand_tokens:
                    continue
                for post in prioritized_posts:
                    nt = norm_ko(post["title"])
                    if all(tok in nt for tok in cand_tokens):
                        hit = post
                        break
                if hit:
                    break
            if hit:
                existing_date = safe_str(d.get("latestDate"))
                hit_date = safe_str(hit.get("pubDate"))
                existing_yr = title_year_of(d.get("latestTitle", ""), existing_date)
                hit_yr = title_year_of(hit.get("title", ""), hit_date)
                # 구버전 역행 차단 가드: 검증된 2026 직링크가 이미 있는 항목은 더 최신 날짜/연도일 때만 갱신
                if d.get("directVerified") and existing_date:
                    if (hit_yr < existing_yr) or (hit_yr == existing_yr and hit_date and hit_date < existing_date):
                        continue
                d["latestTitle"] = hit["title"]
                d["latestDate"] = hit_date or existing_date
                d["latestYear"] = str(hit_yr) if hit_yr else (hit["year"] or "2026")
                d["latestUrl"] = hit["url"]
                d["liveMatched"] = True
                matched_live += 1

    # 탭 1 주간 브리핑용 2026년 KDCA 지침 출간 소식 통합 (실시간 수집분 + 2026 검증 마스터 병합, 중복 제거)
    releases_by_key: Dict[str, Dict[str, Any]] = {}
    for p in live_posts:
        pd = safe_str(p.get("pubDate"))
        if pd >= "2026-01-01":
            rk = f"{pd}||{norm_ko(p['title'])[:24]}"
            releases_by_key[rk] = {
                "title": p["title"],
                "pub_date": pd,
                "week_key": get_iso_week_key(pd),
                "url": p["url"],
                "source": "KDCA 감염병지침",
                "dept": "",
            }

    # 실시간 목록에 없는 검증 완료 2026 KDCA 지침도 보완 병합
    seen_release_titles = {norm_ko(r["title"])[:18] for r in releases_by_key.values()}
    seen_release_urls = {r["url"] for r in releases_by_key.values()}
    for d in updated:
        pd = safe_str(d.get("latestDate"))
        u = safe_str(d.get("latestUrl"))
        t = safe_str(d.get("latestTitle"))
        if pd >= "2026-01-01" and u and u not in seen_release_urls and norm_ko(t)[:18] not in seen_release_titles:
            seen_release_urls.add(u)
            seen_release_titles.add(norm_ko(t)[:18])
            releases_by_key[f"{pd}||{norm_ko(t)[:24]}"] = {
                "title": t,
                "pub_date": pd,
                "week_key": get_iso_week_key(pd),
                "url": u,
                "source": "KDCA 공식 지침",
                "dept": safe_str(d.get("dept")),
            }

    releases_list = sorted(releases_by_key.values(), key=lambda x: (x["pub_date"], x["title"]), reverse=True)
    direct_total = sum(1 for d in updated if d.get("liveMatched") or d.get("directVerified"))

    if live_posts:
        msg = f"KDCA 감염병포털(1001) {len(live_posts)}건 실시간 수집 · 세부지침 직링크 {direct_total}/{len(updated)}건 (실시간 {matched_live}건, 2026 출간 {len(releases_list)}건)"
    else:
        msg = f"KDCA 포털 접속 보류({live_err or '오프라인'}) — 검증 2026 직링크 {direct_total}/{len(updated)}건 유지"
    return updated, releases_list, msg, direct_total


def build_reference_and_notifiable_live(run_live_probe: bool = True) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]], str]:
    rs = json.loads(REFERENCE_STANDARDS_MASTER_JSON)
    notifiable = json.loads(NOTIFIABLE_DISEASES_MASTER_JSON)
    if not run_live_probe:
        # 오프라인 빌드에서도 검증 완료 2026 직링크 기반 주간 브리핑 출간 소식 목록 생성
        _, releases_offline, kdca_offline_msg, _ = check_kdca_guidelines_live_offline(notifiable)
        rs["liveCheckSummary"] = {
            "checkedAt": "",
            "eucastStatus": "오프라인 빌드 (검수본)",
            "clsiStatus": "오프라인 빌드 (검수본)",
            "kdcaStatus": kdca_offline_msg,
            "updatesDetectedCount": 0,
        }
        return rs, notifiable, releases_offline, "오프라인 빌드(검수본)"

    eu_items, eu_msg, eu_new = check_eucast_live(rs.get("eucast", []))
    m_items, mm_items, clsi_msg, clsi_new = check_clsi_live(rs.get("clsi_m", []), rs.get("clsi_mm", []))
    noti_items, kdca_releases, kdca_msg, kdca_direct_cnt = check_kdca_guidelines_live(notifiable)

    rs["eucast"] = eu_items
    rs["clsi_m"] = m_items
    rs["clsi_mm"] = mm_items
    rs["liveCheckSummary"] = {
        "checkedAt": now_kst_str(),
        "eucastStatus": eu_msg,
        "clsiStatus": clsi_msg,
        "kdcaStatus": kdca_msg,
        "kdcaMatchedCount": kdca_direct_cnt,
        "updatesDetectedCount": eu_new + clsi_new,
    }
    summary = f"{eu_msg} | {clsi_msg} | {kdca_msg}"
    return rs, noti_items, kdca_releases, summary


def check_kdca_guidelines_live_offline(notifiable_list: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, int]:
    """네트워크 호출 없이 검증 2026 직링크만으로 kdcaGuidelineReleases 생성 (build_from_csv용)"""
    def norm_ko(s: str) -> str:
        return re.sub(r"[\s·ㆍ\-_()\[\],.]", "", str(s or "")).lower()

    seen_urls: Set[str] = set()
    releases: List[Dict[str, Any]] = []
    for d in notifiable_list:
        pd = safe_str(d.get("latestDate"))
        u = safe_str(d.get("latestUrl"))
        t = safe_str(d.get("latestTitle"))
        if pd >= "2026-01-01" and u and u not in seen_urls:
            seen_urls.add(u)
            releases.append({
                "title": t,
                "pub_date": pd,
                "week_key": get_iso_week_key(pd),
                "url": u,
                "source": "KDCA 공식 지침",
                "dept": safe_str(d.get("dept")),
            })
    releases.sort(key=lambda x: (x["pub_date"], x["title"]), reverse=True)
    direct_cnt = sum(1 for d in notifiable_list if d.get("directVerified"))
    return notifiable_list, releases, f"검증 2026 직링크 {direct_cnt}/{len(notifiable_list)}건 · 출간 {len(releases)}건", direct_cnt


def upsert_config_value(ss, key: str, value: str, description: str = "") -> None:
    """config 시트에서 해당 key 행을 찾아 value를 갱신하거나 없으면 새 행으로 추가"""
    ws = ss.worksheet(SHEET_CONFIG)
    rows = ws.get_all_values()
    for idx, r in enumerate(rows[1:], start=2):
        if r and r[0].strip() == key:
            ws.update(range_name=f"B{idx}", values=[[value]], value_input_option="RAW")
            return
    ws.append_row([key, value, description], value_input_option="RAW")


def _jacobi_top2_eigs(B: List[List[float]]) -> Tuple[List[float], List[List[float]]]:
    """순수 Python Jacobi 고유값 분해 (대칭 행렬 B의 상위 2개 고유값·고유벡터 반환)"""
    n = len(B)
    A = [row[:] for row in B]
    V = [[1.0 if i == j else 0.0 for j in range(n)] for i in range(n)]
    for _ in range(120):
        p, q, max_val = 0, 1, 0.0
        for i in range(n):
            for j in range(i + 1, n):
                av = abs(A[i][j])
                if av > max_val:
                    max_val = av
                    p, q = i, j
        if max_val < 1e-10:
            break
        diff = A[q][q] - A[p][p]
        if abs(A[p][q]) < 1e-12 * abs(diff):
            t = A[p][q] / diff
        else:
            phi = diff / (2.0 * A[p][q])
            t = 1.0 / (abs(phi) + math.sqrt(phi * phi + 1.0))
            if phi < 0:
                t = -t
        c = 1.0 / math.sqrt(t * t + 1.0)
        s = t * c
        tau = s / (1.0 + c)
        h = t * A[p][q]
        A[p][p] -= h
        A[q][q] += h
        A[p][q] = A[q][p] = 0.0
        for r in range(n):
            if r != p and r != q:
                arp, arq = A[r][p], A[r][q]
                A[r][p] = A[p][r] = arp - s * (arq + tau * arp)
                A[r][q] = A[q][r] = arq + s * (arp - tau * arq)
            vrp, vrq = V[r][p], V[r][q]
            V[r][p] = vrp - s * (vrq + tau * vrp)
            V[r][q] = vrq + s * (vrp - tau * vrq)
    eigs = [(A[i][i], i) for i in range(n)]
    eigs.sort(reverse=True)
    vals = [max(eigs[0][0], 1e-6), max(eigs[1][0], 1e-6)]
    vecs = [[V[r][eigs[0][1]], V[r][eigs[1][1]]] for r in range(n)]
    return vals, vecs


def _date_to_quarter_key(pub_date: str) -> str:
    """고정 캘린더 분기 키 반환 (1~3월=Q1, 4~6월=Q2, 7~9월=Q3, 10~12월=Q4)"""
    s = safe_str(pub_date)[:7]
    if len(s) < 7 or not s[:4].isdigit() or not s[5:7].isdigit():
        return ""
    yr = s[:4]
    mo = int(s[5:7])
    q = (mo - 1) // 3 + 1
    return f"{yr}-Q{q}"


def _quarter_meta(qk: str) -> Tuple[str, str, Optional[datetime.date], str]:
    """분기 키 -> (한글 라벨, 월 범위, 종료일 date, 공개 예정일 문자열)"""
    m = re.match(r"^(\d{4})-Q([1-4])$", safe_str(qk))
    if not m:
        return qk, "", None, ""
    yr = int(m.group(1))
    q = int(m.group(2))
    span_map = {1: ("1~3월", 3, 31), 2: ("4~6월", 6, 30), 3: ("7~9월", 9, 30), 4: ("10~12월", 12, 31)}
    span_lbl, end_m, end_d = span_map[q]
    end_date = datetime.date(yr, end_m, end_d)
    avail_date = (end_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    return f"{yr}년 {q}분기 ({span_lbl})", span_lbl, end_date, avail_date


def _is_quarter_completed(qk: str, ref_date: Optional[datetime.date] = None) -> bool:
    """해당 분기가 완전히 종료되었는지 판정 (종료일 다음 날부터 True: 예 11월 기준 Q3=True, Q4=False)"""
    _, _, end_date, _ = _quarter_meta(qk)
    if not end_date:
        return False
    today = ref_date or now_kst().date()
    return today > end_date


def _load_quarterly_trend_cache(cfg_map: Dict[str, str]) -> Dict[str, Dict[str, Any]]:
    """config의 paper_macro_trend_v843에서 완료 분기별 AI 브리핑 캐시 딕셔너리 로드"""
    raw = safe_str(cfg_map.get(TREND_CONFIG_KEY))
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            return {k: v for k, v in parsed.items() if isinstance(v, dict) and v.get("headline")}
    except Exception:
        pass
    return {}


def build_fallback_macro_trend_for_quarter(
    quarter_papers: List[Dict[str, Any]],
    quarter_key: str,
) -> Dict[str, Any]:
    """[AI 비용 0원] 완료된 특정 분기의 필독·주목 논문 기반 간결·사실 중심 구조화 브리핑"""
    if quarter_key == "total":
        period_label = "완료 분기 전체 누적"
    else:
        period_label, _, _, _ = _quarter_meta(quarter_key)

    pool = [
        p for p in quarter_papers
        if safe_str(p.get("importance")) in ("Must-Read", "High")
    ]
    pool.sort(key=lambda x: (1 if x.get("critical_exception") else 0, x.get("importance_score") or 0), reverse=True)
    ce_papers = [p for p in pool if p.get("critical_exception")]
    must_papers = [p for p in pool if p.get("importance") == "Must-Read"]

    def _clean_sentence(txt: str, limit: int = 135) -> str:
        s = re.sub(r"\s+", " ", safe_str(txt))
        s = re.sub(r"^[•\-\s]*\[[^\]]+\]\s*", "", s)
        s = re.sub(r"^\[AI[^\]]*\]\s*", "", s)
        return s[:limit].rstrip() + ("..." if len(s) > limit else "")

    def build_pillar(title: str, cats: Tuple[str, ...]) -> Dict[str, Any]:
        sub = [p for p in pool if safe_str(p.get("topic_category")) in cats]
        if not sub:
            return {
                "title": title,
                "trend_summary": f"{period_label}에는 해당 영역의 필독·주목 선별 논문이 없습니다.",
                "key_evidence": [],
                "lab_checkpoint": "해당 분기 특이 점검 사항 없음.",
            }
        top_titles = [safe_str(p.get("title_ko")) for p in sub[:2] if safe_str(p.get("title_ko"))]
        summary_txt = f"{period_label} 해당 영역 핵심 논문 {len(sub)}편 중 주요 주제는 '{' / '.join(top_titles)}'입니다."
        ev_list = []
        for p in sub[:3]:
            src = safe_str(p.get("source"))
            pd = safe_str(p.get("pub_date"))[5:]
            sc = p.get("importance_score") or 0
            ce_tag = f" · 예외 {p['critical_exception']}" if p.get("critical_exception") else ""
            ev_list.append(f"[{src} {pd} · {sc}점{ce_tag}] {safe_str(p.get('title_ko'))} — {_clean_sentence(p.get('fact_summary_ko'), 120)}")
        cp = _clean_sentence(sub[0].get("ai_clinical_note"), 170) or "원문 기준을 확인해 원내 검사 프로토콜과 대조하십시오."
        return {
            "title": title,
            "trend_summary": summary_txt,
            "key_evidence": ev_list,
            "lab_checkpoint": cp,
        }

    pillars = {
        "diagnostics": build_pillar("분자진단 · 검사실 워크플로우", ("분자진단·신기술",)),
        "amr_ast": build_pillar("항균제 내성 · 감수성 시험(AST)", ("항균제 내성·AST",)),
        "therapy_asp": build_pillar("임상 감염 치료 · 스튜어드십(ASP)", ("임상 감염·치료",)),
        "epidemiology_ic": build_pillar("병원체 역학 · 감염관리", ("감염관리·역학",)),
    }

    top_sources = [f"{safe_str(p.get('source'))}({safe_str(p.get('title_ko'))[:28]})" for p in pool[:3]]
    watch_pts = [
        f"[{safe_str(p.get('source'))} {safe_str(p.get('pub_date'))[5:]}] {safe_str(p.get('title_ko'))}"
        for p in (ce_papers + must_papers)[:3]
    ]
    if not watch_pts:
        watch_pts = [f"[{safe_str(p.get('source'))}] {safe_str(p.get('title_ko'))}" for p in pool[:3]]

    return {
        "quarter_key": quarter_key,
        "period_label": period_label,
        "headline": f"{period_label} 핵심 논문 {len(pool)}편(필독 {len(must_papers)} · 예외 {len(ce_papers)}) 요약입니다.",
        "executive_overview": (
            f"{period_label} 전체 논문 {len(quarter_papers)}편 중 필독·주목 {len(pool)}편을 집계했습니다. "
            + (f"대표 근거는 {', '.join(top_sources)} 등이며, collect 실행 시 AI 심층 분기 브리핑이 영구 생성됩니다." if top_sources else "")
        ),
        "pillars": pillars,
        "watch_points": watch_pts[:3],
        "paper_count": len(pool),
        "total_paper_count": len(quarter_papers),
        "updated_at": now_kst_str(),
        "model": "structured-digest(0원 · collect 시 AI 영구 생성)",
    }


def generate_macro_trend_briefing_ai(
    quarter_papers: List[Dict[str, Any]],
    quarter_key: str,
    deep_model: str,
    fallback_model: str,
    api_key: str,
) -> Dict[str, Any]:
    """완료된 특정 분기의 필독·주목 논문 기반 고퀄리티 AI 분기 브리핑 생성 (분기당 1회 영구 캐싱, ~₩95)"""
    period_label, _, _, _ = _quarter_meta(quarter_key)
    pool = [
        p for p in quarter_papers
        if safe_str(p.get("importance")) in ("Must-Read", "High")
    ]
    pool.sort(key=lambda x: (1 if x.get("critical_exception") else 0, x.get("importance_score") or 0), reverse=True)
    if not pool:
        return build_fallback_macro_trend_for_quarter(quarter_papers, quarter_key)

    sampled = pool[:110]
    lines = []
    for idx, p in enumerate(sampled, 1):
        fact_short = re.sub(r"\s+", " ", safe_str(p.get("fact_summary_ko")))[:320]
        note_short = re.sub(r"\s+", " ", re.sub(r"^\[[^\]]+\]\s*", "", safe_str(p.get("ai_clinical_note"))))[:170]
        ce = f" [예외 {p['critical_exception']}]" if p.get("critical_exception") else ""
        lines.append(
            f"{idx}. [{p.get('source')}|{p.get('pub_date')}|{p.get('importance')} {p.get('importance_score')}점{ce}|{p.get('topic_category')}] "
            f"{p.get('title_ko')}\n   - 사실: {fact_short}\n   - 검사실포인트: {note_short}"
        )

    prompt = "\n".join([
        "당신은 대한민국 대학병원 진단검사의학과(임상미생물·분자진단·항균제 내성) 전문의를 돕는 학술 에디터입니다.",
        f"아래는 {period_label}({quarter_key}) 동안 수집된 필독(Must-Read)·주목(High) 논문 {len(sampled)}편입니다.",
        f"이 자료를 바탕으로 {period_label} 전체 흐름을 한눈에 파악할 수 있는 분기 브리핑을 JSON으로 작성하십시오.",
        "",
        "[문체 및 작성 원칙 — 반드시 지킬 것]",
        "1. 쉽게, 정확하게: 전문 용어(CRE, MBL, AST, mNGS, ddPCR, ASP 등)는 정확히 쓰되, 문장은 짧고 명확하게 풀어 써서 바쁜 전문의가 한 번에 이해할 수 있게 작성하십시오.",
        "2. 상용어구·미사여구·반복 금지:",
        "   - 금지 표현: '주목할 만한', '패러다임 전환', '가속화되고 있습니다', '중요성이 부각되고 있습니다', '참고가 될 것으로 사료됩니다', '시사합니다', '획기적인', '괄목할 만한', '다각도로', '심층적으로' 등 공허한 수식어와 상투적 어미.",
        "   - 같은 논문이나 같은 내용을 총평·4개 영역·관전 포인트에 반복해서 적지 마십시오.",
        "3. 사실 기반·구체성: 원문에 없는 수치·균종·기관명을 절대 지어내지 말고, '무엇이 문제였고 → 무엇이 확인됐으며 → 검사실에서 무엇을 챙겨야 하는지'만 간결히 적으십시오.",
        "",
        "[필드별 작성 기준]",
        "- headline (1문장, 65자 이내): 이번 분기 가장 큰 변화나 실무 이슈를 꾸밈말 없이 직관적으로 압축.",
        "- executive_overview (2~3문장): 이번 분기 핵심 흐름을 진단·내성·치료 관점에서 군더더기 없이 요약.",
        "- pillars (4개 영역: diagnostics, amr_ast, therapy_asp, epidemiology_ic):",
        "  · trend_summary (2~3문장): 이번 분기 해당 영역에서 실제로 나타난 변화와 핵심 이유를 쉽게 설명.",
        "  · key_evidence (2~3개): '[저널명 MM-DD] 핵심 결과 + 구체적 수치(n수, 민감도/특이도, OR/HR, 사망률 등)' 형식으로 간결히 인용.",
        "  · lab_checkpoint (1~2문장): 대학병원 미생물·분자진단 검사실에서 바로 점검하거나 적용할 구체적 행동 1~2가지.",
        "- watch_points (3개): 다음 분기에 검사실·임상에서 계속 지켜봐야 할 구체적 체크 항목 3가지(각 1문장).",
        "",
        "\n".join(lines),
    ])

    pillar_schema = {
        "type": "OBJECT",
        "properties": {
            "trend_summary": {"type": "STRING"},
            "key_evidence": {"type": "ARRAY", "items": {"type": "STRING"}},
            "lab_checkpoint": {"type": "STRING"},
        },
        "required": ["trend_summary", "key_evidence", "lab_checkpoint"],
    }

    schema = {
        "type": "OBJECT",
        "properties": {
            "headline": {"type": "STRING"},
            "executive_overview": {"type": "STRING"},
            "pillars": {
                "type": "OBJECT",
                "properties": {
                    "diagnostics": pillar_schema,
                    "amr_ast": pillar_schema,
                    "therapy_asp": pillar_schema,
                    "epidemiology_ic": pillar_schema,
                },
                "required": ["diagnostics", "amr_ast", "therapy_asp", "epidemiology_ic"],
            },
            "watch_points": {"type": "ARRAY", "items": {"type": "STRING"}},
        },
        "required": ["headline", "executive_overview", "pillars", "watch_points"],
    }

    parsed, used_model = call_gemini_with_fallback(
        prompt, schema, deep_model, fallback_model, api_key, 0.15, "high", timeout_read=135
    )

    raw_pillars = parsed.get("pillars") or {}
    pillar_titles = {
        "diagnostics": "분자진단 · 검사실 워크플로우",
        "amr_ast": "항균제 내성 · 감수성 시험(AST)",
        "therapy_asp": "임상 감염 치료 · 스튜어드십(ASP)",
        "epidemiology_ic": "병원체 역학 · 감염관리",
    }
    clean_pillars: Dict[str, Any] = {}
    for pk, ptitle in pillar_titles.items():
        pobj = raw_pillars.get(pk) or {}
        clean_pillars[pk] = {
            "title": ptitle,
            "trend_summary": safe_str(pobj.get("trend_summary")),
            "key_evidence": [safe_str(x) for x in (pobj.get("key_evidence") or []) if safe_str(x)][:3],
            "lab_checkpoint": safe_str(pobj.get("lab_checkpoint")),
        }

    return {
        "quarter_key": quarter_key,
        "period_label": period_label,
        "headline": safe_str(parsed.get("headline")),
        "executive_overview": safe_str(parsed.get("executive_overview")),
        "pillars": clean_pillars,
        "watch_points": [safe_str(x) for x in (parsed.get("watch_points") or []) if safe_str(x)][:3],
        "paper_count": len(sampled),
        "total_paper_count": len(quarter_papers),
        "updated_at": now_kst_str(),
        "model": f"{used_model}(thinking:high)",
    }


def _mine_novel_clinical_terms(
    docs_text: List[str],
    covered_regexes: List[Any],
    min_df: int = 4,
    max_novel: int = 3,
) -> List[Tuple[str, str, str, str]]:
    """
    [AI 비용 0원] 사전(44개 핵심 개념)에 아직 등록되지 않은 신약(-bactam/-penem/-conazole/-vir/-mab/-cycline 등)이나
    신종 병원체·장비 용어가 해당 분기 논문에서 min_df편 이상 등장하면 자동으로 후보 키워드로 승격함.
    """
    drug_rx = re.compile(
        r"\b([a-z]{4,18}(?:bactam|penem|conazole|cycline|mycin|floxacin|cillin|vir|mab|fungin|vancin|zolid|micin))\b",
        re.IGNORECASE,
    )
    term_docs: Dict[str, Set[int]] = {}
    for idx, txt in enumerate(docs_text):
        for m in drug_rx.finditer(txt):
            w = m.group(1).lower()
            if w in NOVEL_TERM_STOPWORDS or len(w) < 6:
                continue
            if any(rx.search(w) for rx in covered_regexes):
                continue
            term_docs.setdefault(w, set()).add(idx)

    picked = [(len(dset), w) for w, dset in term_docs.items() if len(dset) >= min_df]
    picked.sort(reverse=True)
    out: List[Tuple[str, str, str, str]] = []
    for _, w in picked[:max_novel]:
        kid = "novel_" + re.sub(r"[^a-z0-9]+", "_", w)
        label = w.capitalize()
        pat = r"\b" + re.escape(w) + r"\b"
        out.append((kid, label, "drug", pat))
    return out


def _align_procrustes_2d(
    raw_coords: List[List[float]],
    node_ids: List[str],
    prev_positions: Dict[str, Tuple[float, float]],
    w_px: float = 1000.0,
    h_px: float = 610.0,
) -> List[List[float]]:
    """
    2D Orthogonal Procrustes 정렬:
    이전 분기에도 존재했던 공통 키워드들의 좌표를 기준 앵커로 삼아,
    새 분기 MDS 좌표의 최적 회전·반전·축척·이동을 계산하여 좌표계 뒤집힘(sign-flip)을 방지함.
    """
    n = len(node_ids)
    shared = [i for i, kid in enumerate(node_ids) if kid in prev_positions]
    if len(shared) < 2:
        # 첫 기준 분기: 정규화 후 캔버스 중앙 배치
        mx = sum(c[0] for c in raw_coords) / max(n, 1)
        my = sum(c[1] for c in raw_coords) / max(n, 1)
        centered = [[c[0] - mx, c[1] - my] for c in raw_coords]
        span = max((max(abs(c[0]), abs(c[1])) for c in centered), default=1.0) or 1.0
        return [[w_px * 0.5 + (c[0] / span) * 375.0, h_px * 0.5 + (c[1] / span) * 215.0] for c in centered]

    # Centroids of shared nodes
    pmx = sum(raw_coords[i][0] for i in shared) / len(shared)
    pmy = sum(raw_coords[i][1] for i in shared) / len(shared)
    qmx = sum(prev_positions[node_ids[i]][0] for i in shared) / len(shared)
    qmy = sum(prev_positions[node_ids[i]][1] for i in shared) / len(shared)

    ps = [(raw_coords[i][0] - pmx, raw_coords[i][1] - pmy) for i in shared]
    qs = [(prev_positions[node_ids[i]][0] - qmx, prev_positions[node_ids[i]][1] - qmy) for i in shared]

    # Proper rotation (det = +1)
    a_pos = sum(px * qx + py * qy for (px, py), (qx, qy) in zip(ps, qs))
    b_pos = sum(px * qy - py * qx for (px, py), (qx, qy) in zip(ps, qs))
    score_pos = math.hypot(a_pos, b_pos)

    # Improper rotation / reflection (det = -1, mirror x)
    a_neg = sum(-px * qx + py * qy for (px, py), (qx, qy) in zip(ps, qs))
    b_neg = sum(-px * qy - py * qx for (px, py), (qx, qy) in zip(ps, qs))
    score_neg = math.hypot(a_neg, b_neg)

    use_flip = score_neg > score_pos
    theta = math.atan2(b_neg, a_neg) if use_flip else math.atan2(b_pos, a_pos)
    ct, st = math.cos(theta), math.sin(theta)

    norm_p = math.sqrt(sum(px * px + py * py for px, py in ps))
    norm_q = math.sqrt(sum(qx * qx + qy * qy for qx, qy in qs))
    scale = (norm_q / norm_p) if norm_p > 1e-6 else 260.0
    scale = max(140.0, min(520.0, scale))

    aligned: List[List[float]] = []
    for i in range(n):
        rx = raw_coords[i][0] - pmx
        ry = raw_coords[i][1] - pmy
        if use_flip:
            rx = -rx
        tx = qmx + scale * (ct * rx - st * ry)
        ty = qmy + scale * (st * rx + ct * ry)
        aligned.append([tx, ty])
    return aligned


def _layout_period_mds(
    period_papers: List[Dict[str, Any]],
    selected_concepts: List[Tuple[str, str, str, str]],
    prev_positions: Dict[str, Tuple[float, float]],
    prev_active_ids: Set[str],
    all_seen_ids: Set[str],
    prev_count_map: Dict[str, int],
    total_count_map: Dict[str, int],
    is_base_or_total: bool = False,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Tuple[float, float]]]:
    """
    해당 기간(분기 또는 전체 누적)의 동적 키워드 집합에 대해:
    1) 동시출현(Ochiai) 유사도 행렬 -> Classical MDS
    2) 이전 분기 위치(prev_positions)를 참조한 Procrustes 정렬 + 앵커링
    3) 겹침 해소 후 노드·연관선·최종 좌표 반환
    """
    n_kw = len(selected_concepts)
    if n_kw == 0:
        return [], [], {}

    compiled = [(kid, lbl, cat, pat, re.compile(pat, re.IGNORECASE)) for kid, lbl, cat, pat in selected_concepts]
    n_docs = len(period_papers)
    masks = [[0] * n_docs for _ in range(n_kw)]
    counts = [0] * n_kw

    for d_idx, p in enumerate(period_papers):
        txt = f"{p.get('title_orig', '')} {p.get('title_ko', '')} {p.get('fact_summary_ko', '')}"
        for k_idx, (_, _, _, _, rx) in enumerate(compiled):
            if rx.search(txt):
                masks[k_idx][d_idx] = 1
                counts[k_idx] += 1

    cooc = [[0] * n_kw for _ in range(n_kw)]
    sim = [[0.0] * n_kw for _ in range(n_kw)]
    for i in range(n_kw):
        sim[i][i] = 1.0
        cooc[i][i] = counts[i]
        for j in range(i + 1, n_kw):
            c_ij = sum(1 for d in range(n_docs) if masks[i][d] and masks[j][d])
            cooc[i][j] = cooc[j][i] = c_ij
            denom = math.sqrt(counts[i] * counts[j])
            s_ij = (c_ij / denom) if denom > 0 else 0.0
            sim[i][j] = sim[j][i] = s_ij

    if n_kw >= 3:
        D2 = [[max(0.0, 1.0 - sim[i][j]) for j in range(n_kw)] for i in range(n_kw)]
        row_means = [sum(D2[i]) / n_kw for i in range(n_kw)]
        grand_mean = sum(row_means) / n_kw
        B = [
            [-0.5 * (D2[i][j] - row_means[i] - row_means[j] + grand_mean) for j in range(n_kw)]
            for i in range(n_kw)
        ]
        eigvals, eigvecs = _jacobi_top2_eigs(B)
        raw_coords = [
            [eigvecs[i][0] * math.sqrt(eigvals[0]), eigvecs[i][1] * math.sqrt(eigvals[1])]
            for i in range(n_kw)
        ]
    else:
        raw_coords = [[math.cos(i * 2.1), math.sin(i * 2.1)] for i in range(n_kw)]

    W, H_px = 1000.0, 610.0
    node_ids = [c[0] for c in selected_concepts]
    aligned = _align_procrustes_2d(raw_coords, node_ids, prev_positions, W, H_px)

    # Anchor continuing nodes tightly to previous quarter positions; let new nodes settle near aligned MDS / co-occurring partners
    pos: List[List[float]] = []
    anchor: List[List[float]] = []
    is_anchored: List[bool] = []
    for i, kid in enumerate(node_ids):
        if kid in prev_positions:
            px_prev, py_prev = prev_positions[kid]
            # 85% previous position + 15% aligned new MDS so position remains recognizable across quarters
            ix = 0.85 * px_prev + 0.15 * aligned[i][0]
            iy = 0.85 * py_prev + 0.15 * aligned[i][1]
            pos.append([ix, iy])
            anchor.append([px_prev, py_prev])
            is_anchored.append(True)
        else:
            # New node: blend Procrustes-aligned MDS with weighted centroid of anchored co-occurring neighbors
            nb_w, nb_x, nb_y = 0.0, 0.0, 0.0
            for j, other_id in enumerate(node_ids):
                if i != j and other_id in prev_positions and sim[i][j] > 0.08:
                    nb_w += sim[i][j]
                    nb_x += sim[i][j] * prev_positions[other_id][0]
                    nb_y += sim[i][j] * prev_positions[other_id][1]
            if nb_w > 0.0:
                nx = 0.55 * aligned[i][0] + 0.45 * (nb_x / nb_w)
                ny = 0.55 * aligned[i][1] + 0.45 * (nb_y / nb_w)
            else:
                nx, ny = aligned[i][0], aligned[i][1]
            pos.append([nx, ny])
            anchor.append([nx, ny])
            is_anchored.append(False)

    max_cnt = max(counts) if counts and max(counts) > 0 else 1
    radii = [17.0 + 27.0 * math.sqrt(c / max_cnt) for c in counts]
    lbl_hw = [36.0 + 4.3 * len(lbl) for _, lbl, _, _ in selected_concepts]
    lbl_hh = 12.0

    # Deterministic overlap resolution (continuing nodes resist drifting far from their previous-quarter anchor)
    for step in range(1400):
        moved = False
        for i in range(n_kw):
            for j in range(i + 1, n_kw):
                dx = pos[i][0] - pos[j][0]
                dy = pos[i][1] - pos[j][1]
                dist = math.hypot(dx, dy)
                min_d = radii[i] + radii[j] + 24.0
                if dist < min_d:
                    ux, uy = (dx / dist, dy / dist) if dist > 1e-4 else (math.cos(i + j), math.sin(i + j))
                    push = 0.5 * (min_d - dist) * 0.42
                    wi = 0.35 if (is_anchored[i] and not is_anchored[j]) else (0.65 if (not is_anchored[i] and is_anchored[j]) else 0.5)
                    wj = 1.0 - wi
                    pos[i][0] += ux * push * 2.0 * wi; pos[i][1] += uy * push * 2.0 * wi
                    pos[j][0] -= ux * push * 2.0 * wj; pos[j][1] -= uy * push * 2.0 * wj
                    moved = True

                lix, liy = pos[i][0], pos[i][1] + radii[i] + 15.0
                ljx, ljy = pos[j][0], pos[j][1] + radii[j] + 15.0
                ldx, ldy = abs(lix - ljx), abs(liy - ljy)
                req_x = lbl_hw[i] + lbl_hw[j] + 10.0
                req_y = lbl_hh * 2.0 + 7.0
                if ldx < req_x and ldy < req_y:
                    if (req_x - ldx) < (req_y - ldy):
                        sx = 1.0 if lix >= ljx else -1.0
                        pos[i][0] += sx * (req_x - ldx) * 0.32
                        pos[j][0] -= sx * (req_x - ldx) * 0.32
                    else:
                        sy = 1.0 if liy >= ljy else -1.0
                        pos[i][1] += sy * (req_y - ldy) * 0.32
                        pos[j][1] -= sy * (req_y - ldy) * 0.32
                    moved = True

                for a, b in ((i, j), (j, i)):
                    lx, ly = pos[a][0], pos[a][1] + radii[a] + 15.0
                    cx, cy, cr = pos[b][0], pos[b][1], radii[b]
                    clx = max(lx - lbl_hw[a], min(cx, lx + lbl_hw[a]))
                    cly = max(ly - lbl_hh, min(cy, ly + lbl_hh))
                    if math.hypot(cx - clx, cy - cly) < cr + 9.0:
                        sy = 1.0 if pos[a][1] >= pos[b][1] else -1.0
                        pos[a][1] += sy * 4.5
                        pos[b][1] -= sy * 4.5
                        moved = True

        for i in range(n_kw):
            pull = 0.035 if is_anchored[i] else 0.012
            pos[i][0] = (1.0 - pull) * pos[i][0] + pull * anchor[i][0]
            pos[i][1] = (1.0 - pull) * pos[i][1] + pull * anchor[i][1]
            pad_x = max(radii[i], lbl_hw[i]) + 18.0
            pos[i][0] = max(pad_x, min(W - pad_x, pos[i][0]))
            pos[i][1] = max(radii[i] + 24.0, min(H_px - radii[i] - 32.0, pos[i][1]))
        if not moved and step > 200:
            break

    # If first base quarter (no prev anchors), spread X cleanly across [115, 885]
    if not prev_positions and n_kw >= 3:
        xs = [p[0] for p in pos]
        min_x, max_x = min(xs), max(xs)
        if max_x - min_x > 10.0:
            for i in range(n_kw):
                pad_x = max(radii[i], lbl_hw[i]) + 18.0
                sx = 120.0 + ((pos[i][0] - min_x) / (max_x - min_x)) * 760.0
                pos[i][0] = max(pad_x, min(W - pad_x, sx))

    nodes_out: List[Dict[str, Any]] = []
    final_pos: Dict[str, Tuple[float, float]] = {}
    for i, (kid, lbl, cat, pat) in enumerate(selected_concepts):
        fx, fy = round(pos[i][0], 1), round(pos[i][1], 1)
        final_pos[kid] = (fx, fy)
        neighbors = []
        for j in range(n_kw):
            if i != j and cooc[i][j] > 0:
                neighbors.append((sim[i][j], cooc[i][j], selected_concepts[j][1]))
        neighbors.sort(reverse=True)

        if is_base_or_total:
            status = "base"
        elif kid in prev_active_ids:
            status = "continuing"
        elif kid in all_seen_ids:
            status = "reemerged"
        else:
            status = "new"

        nodes_out.append({
            "id": kid,
            "label": lbl,
            "category": cat,
            "pattern": pat,
            "x": fx,
            "y": fy,
            "count": counts[i],
            "prevCount": prev_count_map.get(kid, 0) if (not is_base_or_total and prev_active_ids) else None,
            "total": total_count_map.get(kid, counts[i]),
            "status": status,
            "topPartners": [{"label": nb[2], "shared": nb[1], "sim": round(nb[0], 2)} for nb in neighbors[:3]],
        })

    edges_out: List[Dict[str, Any]] = []
    for i in range(n_kw):
        for j in range(i + 1, n_kw):
            if sim[i][j] >= 0.14 and cooc[i][j] >= 2:
                edges_out.append({
                    "source": selected_concepts[i][0],
                    "target": selected_concepts[j][0],
                    "sim": round(sim[i][j], 3),
                    "shared": cooc[i][j],
                })

    return nodes_out, edges_out, final_pos


def build_paper_trends_payload(items: List[Dict[str, Any]], cfg_map: Dict[str, str]) -> Dict[str, Any]:
    """
    [AI 비용 0원] v8.4.3 완료 분기 1:1 동기화 동적 MDS 맵 + 분기별 심층 거시 브리핑:
    1) 고정 캘린더 분기(Q1:1~3, Q2:4~6, Q3:7~9, Q4:10~12) 중 완전히 종료된 분기만 분석·노출
       (예: 11월 기준 Q4는 12월 31일 종료 전까지 pendingQuarters로 보류하고 Q3를 기본 표시).
    2) 상단 MDS 클러스터 맵과 하단 심층 브리핑을 분기별로 1:1 동기화(periods[qk].macroBriefing 탑재).
    3) 완료 분기별 AI 심층 브리핑은 config(paper_macro_trend_v843)에 분기당 1회 영구 캐싱.
    """
    active_papers = [
        x for x in items
        if x.get("content_type") == "paper" and safe_str(x.get("source")).upper() != "NRM"
    ]
    by_quarter: Dict[str, List[Dict[str, Any]]] = {}
    for p in active_papers:
        qk = _date_to_quarter_key(p.get("pub_date", ""))
        if qk:
            by_quarter.setdefault(qk, []).append(p)

    today_d = now_kst().date()
    all_quarters = sorted(by_quarter.keys())
    completed_quarters: List[str] = []
    pending_quarters: List[Dict[str, Any]] = []

    for qk in all_quarters:
        cnt_q = len(by_quarter[qk])
        q_lbl, _, _, avail_str = _quarter_meta(qk)
        if _is_quarter_completed(qk, today_d):
            if cnt_q >= 5:
                completed_quarters.append(qk)
        else:
            if cnt_q > 0:
                pending_quarters.append({
                    "key": qk,
                    "label": q_lbl,
                    "paperCount": cnt_q,
                    "availableAfter": avail_str,
                })

    years = sorted({q[:4] for q in completed_quarters}, reverse=True)
    completed_papers = [p for qk in completed_quarters for p in by_quarter.get(qk, [])]

    pool_compiled = [(kid, lbl, cat, pat, re.compile(pat, re.IGNORECASE)) for kid, lbl, cat, pat in DYNAMIC_CONCEPT_POOL_V842]
    pool_regexes = [rx for _, _, _, _, rx in pool_compiled]

    # 완료 분기 전체 누적 카운트
    all_completed_texts = [f"{p.get('title_orig', '')} {p.get('title_ko', '')} {p.get('fact_summary_ko', '')}" for p in completed_papers]
    total_count_map: Dict[str, int] = {}
    for kid, _, _, _, rx in pool_compiled:
        total_count_map[kid] = sum(1 for t in all_completed_texts if rx.search(t))

    cached_by_quarter = _load_quarterly_trend_cache(cfg_map)

    periods_payload: Dict[str, Dict[str, Any]] = {}
    prev_positions: Dict[str, Tuple[float, float]] = {}
    prev_active_ids: Set[str] = set()
    prev_active_meta: Dict[str, Tuple[str, str, str, int, float, float]] = {}
    prev_count_map_all: Dict[str, int] = {}
    all_seen_ids: Set[str] = set()
    prev_q_key = ""

    for q_idx, qk in enumerate(completed_quarters):
        q_papers = by_quarter[qk]
        q_texts = [f"{p.get('title_orig', '')} {p.get('title_ko', '')} {p.get('fact_summary_ko', '')}" for p in q_papers]
        min_df = 2 if len(q_papers) < 100 else 3

        # Layer 1: 개념 풀 집계
        q_counts_map: Dict[str, int] = {}
        scored_cands: List[Tuple[int, str, str, str, str]] = []
        for kid, lbl, cat, pat, rx in pool_compiled:
            c = sum(1 for t in q_texts if rx.search(t))
            q_counts_map[kid] = c
            if c >= min_df:
                scored_cands.append((c, kid, lbl, cat, pat))

        # Layer 2: 미등록 신종 약제 자동 발굴
        novel_cands = _mine_novel_clinical_terms(q_texts, pool_regexes, min_df=max(4, min_df + 1), max_novel=2)
        for kid, lbl, cat, pat in novel_cands:
            rx_n = re.compile(pat, re.IGNORECASE)
            c = sum(1 for t in q_texts if rx_n.search(t))
            q_counts_map[kid] = c
            total_count_map[kid] = sum(1 for t in all_completed_texts if rx_n.search(t))
            scored_cands.append((c, kid, lbl, cat, pat))

        scored_cands.sort(key=lambda x: (x[0], 1 if x[1] in prev_active_ids else 0), reverse=True)
        top_selected = [(kid, lbl, cat, pat) for _, kid, lbl, cat, pat in scored_cands[:18]]

        nodes_q, edges_q, final_pos_q = _layout_period_mds(
            period_papers=q_papers,
            selected_concepts=top_selected,
            prev_positions=prev_positions,
            prev_active_ids=prev_active_ids,
            all_seen_ids=all_seen_ids,
            prev_count_map=prev_count_map_all,
            total_count_map=total_count_map,
            is_base_or_total=(q_idx == 0),
        )

        cur_ids = {nd["id"] for nd in nodes_q}
        faded_list = []
        if prev_q_key:
            for pid, (plbl, pcat, ppat, pcnt, px, py) in prev_active_meta.items():
                if pid not in cur_ids:
                    faded_list.append({
                        "id": pid,
                        "label": plbl,
                        "category": pcat,
                        "pattern": ppat,
                        "prevCount": pcnt,
                        "count": q_counts_map.get(pid, 0),
                        "x": px,
                        "y": py,
                    })
            faded_list.sort(key=lambda x: x["prevCount"], reverse=True)

        max_c = 1
        for nd in nodes_q:
            max_c = max(max_c, nd["count"], nd["prevCount"] or 0)

        q_lbl, _, _, _ = _quarter_meta(qk)
        q_briefing = cached_by_quarter.get(qk) or build_fallback_macro_trend_for_quarter(q_papers, qk)

        periods_payload[qk] = {
            "key": qk,
            "label": q_lbl,
            "prevKey": prev_q_key,
            "paperCount": len(q_papers),
            "maxCount": max_c,
            "nodes": nodes_q,
            "edges": edges_q,
            "faded": faded_list,
            "macroBriefing": q_briefing,
        }

        for kid, xy in final_pos_q.items():
            prev_positions[kid] = xy
        all_seen_ids.update(cur_ids)
        prev_active_ids = cur_ids
        prev_active_meta = {
            nd["id"]: (nd["label"], nd["category"], nd["pattern"], nd["count"], nd["x"], nd["y"])
            for nd in nodes_q
        }
        prev_count_map_all = q_counts_map
        prev_q_key = qk

    # 완료 분기 전체 누적("total") 뷰
    tot_cands = []
    for kid, lbl, cat, pat, _ in pool_compiled:
        c = total_count_map.get(kid, 0)
        if c >= 3:
            tot_cands.append((c, kid, lbl, cat, pat))
    tot_cands.sort(key=lambda x: (x[0], 1 if x[1] in prev_active_ids else 0), reverse=True)
    top_total = [(kid, lbl, cat, pat) for _, kid, lbl, cat, pat in tot_cands[:18]]

    nodes_tot, edges_tot, _ = _layout_period_mds(
        period_papers=completed_papers,
        selected_concepts=top_total,
        prev_positions=prev_positions,
        prev_active_ids=set(),
        all_seen_ids=all_seen_ids,
        prev_count_map={},
        total_count_map=total_count_map,
        is_base_or_total=True,
    )
    default_period = completed_quarters[-1] if completed_quarters else "total"
    latest_briefing = (
        periods_payload[default_period]["macroBriefing"]
        if default_period in periods_payload
        else build_fallback_macro_trend_for_quarter(completed_papers, "total")
    )
    periods_payload["total"] = {
        "key": "total",
        "label": "완료 분기 전체 누적",
        "prevKey": "",
        "paperCount": len(completed_papers),
        "maxCount": max([1] + [nd["count"] for nd in nodes_tot]),
        "nodes": nodes_tot,
        "edges": edges_tot,
        "faded": [],
        "macroBriefing": build_fallback_macro_trend_for_quarter(completed_papers, "total"),
    }

    return {
        "mode": "completed_quarter_synced_v843",
        "years": years,
        "quarters": completed_quarters,
        "pendingQuarters": pending_quarters,
        "defaultPeriod": default_period,
        "viewport": {"width": 1000, "height": 610},
        "periods": periods_payload,
        "macroBriefing": latest_briefing,
    }


def write_dashboard_json(
    item_rows: List[Dict[str, Any]],
    surv_rows: List[Dict[str, Any]],
    hai_rows: List[Dict[str, Any]],
    log_rows: List[Dict[str, Any]],
    cfg_map: Dict[str, str],
    run_live_standards_probe: bool = False,
    precomputed_live: Optional[Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]], str]] = None,
) -> str:
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
        "notifiableDiseases": [],
        "kdcaGuidelineReleases": [],
        "referenceStandards": {},
        "paperTrends": build_paper_trends_payload(items, cfg_map),
    }
    if precomputed_live is not None:
        rs_live, noti_live, kdca_rel_live, live_std_msg = precomputed_live
    else:
        rs_live, noti_live, kdca_rel_live, live_std_msg = build_reference_and_notifiable_live(run_live_probe=run_live_standards_probe)
    payload["referenceStandards"] = rs_live
    payload["notifiableDiseases"] = noti_live
    payload["kdcaGuidelineReleases"] = kdca_rel_live

    OUTPUT_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"[OK] dashboard.json 생성 완료: {OUTPUT_JSON_PATH} ({len(items)}건, {OUTPUT_JSON_PATH.stat().st_size:,} bytes)")
    return live_std_msg


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


def sync_sheet_to_dashboard_json(ss, run_live_standards_probe: bool = True) -> str:
    cfg_map = load_config_map(ss)
    items = read_sheet_as_dicts(ss, SHEET_ITEMS)
    surv = read_sheet_as_dicts(ss, SHEET_KDCA_SURV)
    hai = read_sheet_as_dicts(ss, SHEET_HAI_STATS)
    logs = read_sheet_as_dicts(ss, SHEET_LOGS)
    return write_dashboard_json(items, surv, hai, logs, cfg_map, run_live_standards_probe=run_live_standards_probe)


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
    """
    [AI API 비용 0원] API 연결 진단 + v8.3 신규 7개 저널(2026-07-01 이후) 사전 Dry-run + 공식 지침 자동 감지 실측.
    - 본 DB(items)는 전혀 건드리지 않으며, 신규 7개 저널 후보 목록은 별도 시트 탭(dryrun_v83)에만 저장합니다.
    - CLSI(M 20종·MM 7종)·EUCAST(5종) 확장 현황판 및 자동 점검 결과를 dashboard.json에 함께 반영합니다.
    """
    print(f"=== 임상미생물 브리핑 {CODE_VERSION} 0원 Dry-run · 공식 지침 자동 점검 ===")
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    lite_model, deep_model, fallback_model, thinking_p1, thinking_p2 = get_gemini_models(cfg)
    cutoffs = get_score_cutoffs(cfg)
    eff_start, guard_active = resolve_effective_paper_start(cfg, allow_h1)
    now_str = now_kst_str()
    run_id = "GH_DRYRUN_" + now_kst().strftime("%Y%m%d_%H%M%S")

    print(f"[OK] Google Sheets 연결 성공: {ss.title}")
    print(f"   Dry-run 조회 시작일: {eff_start} (상반기 소급 가드: {'ON - 2026-07-01 이후' if guard_active else 'OFF'})")
    print(f"   2-Pass 모델 설정: P1={deep_model}({thinking_p1}) / P2={deep_model}({thinking_p2}) / 우회={fallback_model} / 경보={lite_model}")
    print(f"   4단계 컷오프: Must-Read>={cutoffs['must_read']}, High>={cutoffs['high']}, Medium>={cutoffs['medium']}")

    ncbi_key = os.environ.get("NCBI_API_KEY", "").strip()
    kdca_key = os.environ.get("KDCA_API_KEY", "").strip()
    pubmed_delay = 0.15 if ncbi_key else 0.4

    # 1) 기존 items 탭의 PMID·DOI 목록 확보 (중복 제외 산출용)
    existing_items = read_sheet_as_dicts(ss, SHEET_ITEMS)
    existing_pmids = {clean_pmid(r.get("pmid")) for r in existing_items if clean_pmid(r.get("pmid"))}
    existing_dois = {normalize_doi(r.get("doi")) for r in existing_items if normalize_doi(r.get("doi"))}

    # 2) 신규 7개 저널 PubMed 0원 사전 필터 Dry-run (Gemini 호출 0건!)
    dryrun_days = calc_days_since(eff_start, 160)
    cutoff_dt = parse_flexible_date(eff_start) or datetime.date(2026, 7, 1)
    dryrun_rows: List[List[Any]] = []
    summary_parts: List[str] = []
    total_fetched = 0
    total_net_new = 0
    total_overlap = 0

    print(f"\n--- [Step 1] 신규 7개 저널 PubMed 0원 사전 필터 Dry-run ({eff_start} 이후, AI 호출 0건) ---")
    for code, full_name, query, policy_note in NEW_JOURNALS_V83:
        try:
            time.sleep(pubmed_delay)
            ids = fetch_pubmed_ids(query, dryrun_days, 400, ncbi_key)
            fetched_cnt = len(ids)
            total_fetched += fetched_cnt
            net_arts: List[Dict[str, Any]] = []
            overlap_cnt = 0
            if ids:
                time.sleep(pubmed_delay)
                for art in fetch_pubmed_articles(ids, ncbi_key):
                    pd_obj = parse_flexible_date(art.get("pubDate"))
                    if pd_obj and pd_obj < cutoff_dt:
                        continue
                    pmid_c = clean_pmid(art.get("pmid"))
                    doi_c = normalize_doi(art.get("doi"))
                    if (pmid_c and pmid_c in existing_pmids) or (doi_c and doi_c in existing_dois):
                        overlap_cnt += 1
                        continue
                    net_arts.append(art)
            total_overlap += overlap_cnt
            total_net_new += len(net_arts)
            summary_parts.append(f"{code}:{len(net_arts)}편")
            print(f"  [{code:9s}] {full_name:32s} -> 검색 {fetched_cnt:3d}건 | {eff_start}+ 순신규 {len(net_arts):3d}편 (기존중복 {overlap_cnt}건) | {policy_note}")
            for s_idx, a in enumerate(net_arts[:5], 1):
                print(f"      샘플{s_idx}) [{a.get('pubDate', '')}] ({a.get('pubTypes', '')[:35]}) {a.get('title', '')[:95]}")
            for a in net_arts:
                dryrun_rows.append(
                    [
                        code,
                        full_name,
                        a.get("pubDate", ""),
                        a.get("pmid", ""),
                        a.get("doi", ""),
                        a.get("pubTypes", ""),
                        policy_note,
                        "Y" if len(safe_str(a.get("abstractText"))) >= 40 else "N",
                        a.get("title", ""),
                        f"https://pubmed.ncbi.nlm.nih.gov/{a.get('pmid', '')}/",
                    ]
                )
        except Exception as e:
            summary_parts.append(f"{code}:오류")
            print(f"  [{code:9s}] 조회 오류: {e}")

    dryrun_rows.sort(key=lambda r: (r[2], r[0]), reverse=True)
    dryrun_headers = [
        "journal_code",
        "journal_name",
        "pub_date",
        "pmid",
        "doi",
        "pub_types",
        "cap_policy",
        "has_abstract",
        "title_orig",
        "source_url",
    ]
    dr_ws = get_or_create_worksheet(ss, SHEET_DRYRUN, dryrun_headers)
    dr_ws.clear()
    dr_ws.update(range_name="A1", values=[dryrun_headers] + dryrun_rows, value_input_option="RAW")
    print(f"\n[OK] Dry-run 후보 총 {total_net_new}편을 Google Sheets '{SHEET_DRYRUN}' 탭에 저장 완료 (본 DB items 미변경 · AI 비용 0원)")

    # 3) KDCA 4년치 API 간단 점검
    kdca_check = "미등록"
    if kdca_key:
        try:
            yr = str(now_kst().year)
            sample = call_kdca_eid(
                kdca_key,
                "PeriodBasic01",
                {"resType": "2", "searchPeriodType": "2", "searchStartYear": yr, "searchEndYear": yr, "pageNo": "1", "numOfRows": "5"},
            )
            kdca_check = f"정상({len(sample)}건 샘플)"
        except Exception as e:
            kdca_check = f"경고({str(e)[:60]})"
    print(f"\n--- [Step 2] KDCA API 점검: {kdca_check} ---")

    # 4) 공식 지침(CLSI 27종·EUCAST 5종·KDCA 지침 게시판) 라이브 자동 점검 + dashboard.json 최신 반영
    print("\n--- [Step 3] CLSI(M 20·MM 7) · EUCAST(5) · KDCA 지침 게시판 자동 점검 및 dashboard.json 반영 ---")
    rs_live, noti_live, kdca_rel_live, live_std_msg = build_reference_and_notifiable_live(run_live_probe=True)
    print(f"[공식 지침 점검 결과] {live_std_msg}")

    log_note = (
        f"v8.3 Dry-run(0원, {eff_start}+) 총 {total_net_new}편 -> {SHEET_DRYRUN} 탭 저장 "
        f"({', '.join(summary_parts)}) | {live_std_msg}"
    )
    append_run_log(ss, run_id, now_str, "DRYRUN_V83", "Success", total_fetched, total_net_new, total_overlap, log_note[:490])

    # Write dashboard.json with both the fresh DRYRUN_V83 log row AND the live standards/KDCA check results!
    cfg_map = load_config_map(ss)
    items_all = read_sheet_as_dicts(ss, SHEET_ITEMS)
    surv_all = read_sheet_as_dicts(ss, SHEET_KDCA_SURV)
    hai_all = read_sheet_as_dicts(ss, SHEET_HAI_STATS)
    logs_all = read_sheet_as_dicts(ss, SHEET_LOGS)
    write_dashboard_json(
        items_all, surv_all, hai_all, logs_all, cfg_map,
        run_live_standards_probe=False,
        precomputed_live=(rs_live, noti_live, kdca_rel_live, live_std_msg),
    )


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
    if is_enabled(cfg, "source_v83_new_journals_enabled"):
        for nj_code, _nj_name, nj_query, _nj_note in NEW_JOURNALS_V83:
            journals.append(("source_v83_new_journals_enabled", nj_code, nj_query))
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

    # 완료된 최신 분기의 AI 심층 브리핑이 아직 캐시에 없을 때만 분기당 1회 생성·영구 캐싱 (~₩95, 중복 과금 0원)
    try:
        cached_trends = _load_quarterly_trend_cache(cfg)
        all_items_now = read_sheet_as_dicts(ss, SHEET_ITEMS)
        active_papers_now = [
            x for x in all_items_now
            if safe_str(x.get("content_type")) == "paper" and safe_str(x.get("source")).upper() != "NRM"
        ]
        by_q_now: Dict[str, List[Dict[str, Any]]] = {}
        for p in active_papers_now:
            qk = _date_to_quarter_key(p.get("pub_date", ""))
            if qk and _is_quarter_completed(qk):
                by_q_now.setdefault(qk, []).append(p)
        completed_qs_now = sorted(q for q, plist in by_q_now.items() if len(plist) >= 5)
        if completed_qs_now:
            latest_done_q = completed_qs_now[-1]
            if latest_done_q not in cached_trends:
                trend_obj = generate_macro_trend_briefing_ai(
                    by_q_now[latest_done_q], latest_done_q, deep_model, fallback_model, gemini_key
                )
                cached_trends[latest_done_q] = trend_obj
                upsert_config_value(
                    ss,
                    TREND_CONFIG_KEY,
                    json.dumps(cached_trends, ensure_ascii=False, separators=(",", ":")),
                    "v8.4.3 완료 분기별 AI 논문 심층 브리핑 영구 캐시",
                )
                print(f"[OK] 완료 분기 AI 브리핑 영구 생성 완료 ({latest_done_q}, {trend_obj.get('model')})")
            else:
                print(f"[SKIP] 최신 완료 분기({latest_done_q}) AI 브리핑 이미 영구 캐시됨 (추가 과금 0원)")
    except Exception as tr_err:
        print(f"[경고] 완료 분기 AI 브리핑 생성 건너뜀(기존 캐시 유지): {tr_err}")

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


def mode_update_trend_only() -> None:
    """완료된 분기 중 미생성 AI 브리핑을 생성·영구 캐싱 + dashboard.json 재생성"""
    ss = connect_spreadsheet()
    cfg = load_config_map(ss)
    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not gemini_key:
        raise RuntimeError("GEMINI_API_KEY 가 설정되지 않았습니다.")
    _, deep_model, fallback_model, _, _ = get_gemini_models(cfg)
    all_items = read_sheet_as_dicts(ss, SHEET_ITEMS)
    active_papers = [
        x for x in all_items
        if safe_str(x.get("content_type")) == "paper" and safe_str(x.get("source")).upper() != "NRM"
    ]
    by_q: Dict[str, List[Dict[str, Any]]] = {}
    for p in active_papers:
        qk = _date_to_quarter_key(p.get("pub_date", ""))
        if qk and _is_quarter_completed(qk):
            by_q.setdefault(qk, []).append(p)
    completed_qs = sorted(q for q, plist in by_q.items() if len(plist) >= 5)
    cached_trends = _load_quarterly_trend_cache(cfg)
    generated = []
    for qk in completed_qs[-2:]:
        if qk not in cached_trends:
            cached_trends[qk] = generate_macro_trend_briefing_ai(by_q[qk], qk, deep_model, fallback_model, gemini_key)
            generated.append(qk)
    if generated:
        upsert_config_value(
            ss,
            TREND_CONFIG_KEY,
            json.dumps(cached_trends, ensure_ascii=False, separators=(",", ":")),
            "v8.4.3 완료 분기별 AI 논문 심층 브리핑 영구 캐시",
        )
    append_run_log(
        ss,
        "GH_TREND_" + now_kst().strftime("%Y%m%d_%H%M%S"),
        now_kst_str(),
        "TREND_AI",
        "Success",
        len(completed_qs),
        len(generated),
        0,
        f"{CODE_VERSION} 완료 분기 AI 브리핑 생성: {', '.join(generated) if generated else '모두 캐시됨(0원)'}",
    )
    sync_sheet_to_dashboard_json(ss)


def main():
    parser = argparse.ArgumentParser(description="임상미생물 주간 브리핑 파이프라인 v8.3.0")
    parser.add_argument(
        "--mode",
        choices=["sync_kdca_and_dashboard", "sync_sheet_only", "collect", "resummarize_all", "reclassify_only", "pilot_20", "test_apis", "update_trend_only", "build_from_csv"],
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
    elif args.mode == "update_trend_only":
        mode_update_trend_only()
    elif args.mode == "resummarize_all":
        mode_resummarize_all(force=args.force, allow_h1=args.allow_h1)
    else:
        mode_collect(allow_h1=args.allow_h1)


if __name__ == "__main__":
    main()
