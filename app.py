"""
CRC Manuscript Builder – Streamlit app (MVP)
Author: ChatGPT (for Jun)
Date: 2025-09-01 (KST)

What this app does
------------------
1) 사용자가 주제/연구계획서/결과를 입력하면
2) PubMed에서 최대 50편 검색(eUtils) + 사용자가 PDF 업로드(최대 50개)
3) DOI/메타데이터 정리, 저널 IF(또는 대체지표)로 정렬, 다운로드/요청 링크 제공
4) 섹션별(cover letter, title page, abstract, intro, methods, results, discussion, references)로 분리 생성
   - 인용 태그 형식: [CITE:DOI]  (예: [CITE:10.1056/NEJMxxxx])
   - 생성물은 업로드/선택한 문헌만 인용 가능(가드)
   - 섹션별로 작성 후, 마지막에 합치기 + 레퍼런스 넘버 지속(글로벌 관리)
5) 최종 병합 시, 최초 등장 순서대로 참조를 [1], [2], …로 재번호화(Vancouver 스타일)하고
   레퍼런스 리스트를 자동 생성
6) 파일 내보내기: Markdown(.md), Word(.docx)

중요한 설계 원칙
----------------
- 절대 허구/환각 금지: 허용된 DOI 집합 외 인용 시, 문구 거부/수정 요구
- 근거 중심: NCRN/ESMO/ASCRS 등 가이드라인 DOI를 “앵커”로 함께 넣을 것을 권장
- IF: Clarivate JIF는 공개 API가 없어 CSV 업로드를 권장(journal_if.csv). 미제공 시 OpenAlex 지표로 대체(선택 기능)
- 개인정보/PHI 포함 금지. 로컬 실행 추천.

필수 패키지(권장)
-----------------
streamlit, requests, pandas, lxml, PyMuPDF(fitz), python-docx, pydantic, tenacity

설치 예:
  pip install streamlit requests pandas lxml pymupdf python-docx pydantic tenacity

실행:
  streamlit run app.py

환경변수(선택):
  OPENAI_API_KEY   : OpenAI 키
  OPENAI_MODEL     : 기본 gpt-4o-mini (원하면 gpt-5 또는 적절한 모델명으로 변경)
  UNPAYWALL_EMAIL  : OA 링크 확인용 이메일(선택)
  ENTREZ_API_KEY   : NCBI eUtils 쿼터 완화 용(선택)

파일(선택):
  journal_if.csv   : columns = journal, if    (사용자 제공)

"""
from __future__ import annotations
import os
import io
import re
import json
import time
import math
import fitz  # PyMuPDF
import base64
import string
import random
import zipfile
import textwrap
import requests
import streamlit as st
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from lxml import etree
from tenacity import retry, stop_after_attempt, wait_exponential
from pydantic import BaseModel, Field
from docx import Document
from docx.shared import Pt

# =====================
# Config & constants
# =====================
APP_TITLE = "CRC Manuscript Builder (MVP)"
EUTILS_BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"
CROSSREF_BASE = "https://api.crossref.org/works"
OPENALEX_BASE = "https://api.openalex.org/sources"
UNPAYWALL_BASE = "https://api.unpaywall.org/v2/"

MAX_RESULTS = 50
MAX_UPLOADS = 50

OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ENTREZ_API_KEY = os.getenv("ENTREZ_API_KEY")
UNPAYWALL_EMAIL = os.getenv("UNPAYWALL_EMAIL")

# =====================
# Utilities
# =====================
DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+")


def norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def find_doi_in_text(txt: str) -> Optional[str]:
    m = DOI_RE.search(txt or "")
    return m.group(0) if m else None


def extract_pdf_text_and_doi(file_bytes: bytes) -> Tuple[str, Optional[str]]:
    try:
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            texts = []
            for page in doc:
                texts.append(page.get_text())
            full = "\n".join(texts)
            doi = find_doi_in_text(full)
            return full, doi
    except Exception:
        return "", None


# =====================
# Data models
# =====================
class RefMeta(BaseModel):
    doi: Optional[str] = None
    title: Optional[str] = None
    journal: Optional[str] = None
    year: Optional[str] = None
    authors: List[str] = Field(default_factory=list)
    pmid: Optional[str] = None
    url: Optional[str] = None


class ReferenceManager:
    """Global reference registry preserving numbering across sections.
    Internally keyed by DOI (fallback to PMID if DOI missing)."""

    def __init__(self):
        self.idx: Dict[str, int] = {}
        self.meta: Dict[str, RefMeta] = {}
        self.order: List[str] = []  # key order by first citation

    def _key(self, m: RefMeta) -> Optional[str]:
        if m.doi:
            return m.doi.lower()
        if m.pmid:
            return f"pmid:{m.pmid}"
        return None

    def register(self, meta: RefMeta) -> Optional[int]:
        key = self._key(meta)
        if not key:
            return None
        if key not in self.idx:
            self.order.append(key)
            self.idx[key] = len(self.order)
            self.meta[key] = meta
        else:
            # Update meta if we learn more
            old = self.meta.get(key)
            self.meta[key] = meta if len(json.dumps(meta.dict())) > len(json.dumps(old.dict())) else old
        return self.idx[key]

    def cite(self, doi_or_pmid: str) -> Optional[int]:
        key = doi_or_pmid.lower() if not doi_or_pmid.startswith("pmid:") else doi_or_pmid
        if key not in self.idx:
            # placeholder meta
            self.idx[key] = len(self.order) + 1
            self.order.append(key)
            self.meta.setdefault(key, RefMeta())
        return self.idx[key]

    def renumber_by_first_appearance(self, citation_sequence: List[str]):
        seen = []
        for key in citation_sequence:
            if key not in seen and key in self.idx:
                seen.append(key)
        # append any not seen
        for key in self.order:
            if key not in seen:
                seen.append(key)
        self.order = seen
        self.idx = {k: i + 1 for i, k in enumerate(self.order)}

    def vancouver(self, key: str) -> str:
        m = self.meta.get(key, RefMeta())
        # Authors: Lastname Initials, up to 6 then et al.
        def fmt_author(a: str) -> str:
            a = a.strip()
            # naive split
            parts = a.split()
            if not parts:
                return a
            last = parts[-1]
            initials = "".join(p[0] for p in parts[:-1] if p)
            return f"{last} {initials}".strip()

        authors = ", ".join(fmt_author(a) for a in (m.authors or [])[:6])
        if m.authors and len(m.authors) > 6:
            authors += ", et al."
        year = m.year or ""
        journal = m.journal or ""
        title = m.title or ""
        doi = f" https://doi.org/{m.doi}" if m.doi else (f" PMID:{m.pmid}" if m.pmid else "")
        return norm_text(f"{authors}. {title}. {journal}. {year}.{doi}")

    def render_reference_list(self) -> List[str]:
        return [self.vancouver(k) for k in self.order]


# =====================
# External services
# =====================
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=1, max=6))
def pubmed_search(term: str, retmax: int = MAX_RESULTS) -> List[str]:
    params = {
        "db": "pubmed",
        "term": term,
        "retmode": "json",
        "retmax": retmax,
        "sort": "relevance",
    }
    if ENTREZ_API_KEY:
        params["api_key"] = ENTREZ_API_KEY
    r = requests.get(f"{EUTILS_BASE}/esearch.fcgi", params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    return data.get("esearchresult", {}).get("idlist", [])


@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=1, max=6))
def pubmed_fetch_xml(pmids: List[str]) -> etree._Element:
    if not pmids:
        return etree.Element("empty")
    params = {
        "db": "pubmed",
        "id": ",".join(pmids),
        "retmode": "xml",
    }
    if ENTREZ_API_KEY:
        params["api_key"] = ENTREZ_API_KEY
    r = requests.get(f"{EUTILS_BASE}/efetch.fcgi", params=params, timeout=60)
    r.raise_for_status()
    return etree.fromstring(r.content)


def _xml_text(node, xpath: str) -> Optional[str]:
    el = node.find(xpath)
    return norm_text(el.text) if el is not None and el.text else None


def pubmed_parse_records(root: etree._Element) -> List[RefMeta]:
    out: List[RefMeta] = []
    for art in root.findall(".//PubmedArticle"):
        pmid = _xml_text(art, ".//MedlineCitation/PMID")
        title = _xml_text(art, ".//Article/ArticleTitle")
        year = _xml_text(art, ".//Article/Journal/JournalIssue/PubDate/Year")
        journal = _xml_text(art, ".//Article/Journal/Title")
        # Authors
        authors = []
        for a in art.findall(".//AuthorList/Author"):
            ln = _xml_text(a, "LastName") or ""
            fn = _xml_text(a, "ForeName") or ""
            full = norm_text(f"{fn} {ln}")
            if full:
                authors.append(full)
        # DOI
        doi = None
        for idn in art.findall(".//ArticleIdList/ArticleId"):
            if idn.get("IdType") == "doi" and idn.text:
                doi = idn.text.lower()
        url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else None
        out.append(RefMeta(doi=doi, title=title, journal=journal, year=year, authors=authors, pmid=pmid, url=url))
    return out


def crossref_get(doi: str) -> Optional[RefMeta]:
    try:
        r = requests.get(f"{CROSSREF_BASE}/{doi}", timeout=30)
        if r.status_code != 200:
            return None
        j = r.json().get("message", {})
        title = "; ".join(j.get("title", [])) or None
        journal = (j.get("container-title") or [None])[0]
        year = None
        if j.get("issued", {}).get("'date-parts'"):
            year = str(j["issued"]["'date-parts'"][0][0])
        elif j.get("issued", {}).get("date-parts"):
            year = str(j["issued"]["date-parts"][0][0])
        authors = []
        for a in j.get("author", []):
            nm = norm_text(f"{a.get('given','')} {a.get('family','')}")
            if nm:
                authors.append(nm)
        url = j.get("URL")
        return RefMeta(doi=doi.lower(), title=title, journal=journal, year=year, authors=authors, url=url)
    except Exception:
        return None


def openalex_metric(journal_title: str) -> Optional[float]:
    try:
        q = {"search": journal_title}
        r = requests.get(OPENALEX_BASE, params=q, timeout=20)
        if r.status_code != 200:
            return None
        data = r.json().get("results", [])
        if not data:
            return None
        # Use SJR2 or H-index-ish metric as proxy. Prefer higher.
        src = data[0]
        sjr2 = None
        try:
            sjr2 = src.get("summary_stats", {}).get("two_year_sjr")
        except Exception:
            sjr2 = None
        return float(sjr2) if sjr2 is not None else None
    except Exception:
        return None


def unpaywall_best_oa_link(doi: str) -> Optional[str]:
    if not UNPAYWALL_EMAIL:
        return None
    try:
        r = requests.get(f"{UNPAYWALL_BASE}{doi}", params={"email": UNPAYWALL_EMAIL}, timeout=20)
        if r.status_code != 200:
            return None
        j = r.json()
        rec = j.get("best_oa_location") or {}
        return rec.get("url")
    except Exception:
        return None


# =====================
# LLM wrapper (OpenAI)
# =====================
class LLM:
    def __init__(self, model: str, api_key: Optional[str]):
        self.model = model
        self.key = api_key
        self.enabled = bool(api_key)
        if not self.enabled:
            st.warning("OPENAI_API_KEY 가 설정되어 있지 않습니다. 생성 기능은 비활성화됩니다.")

    def generate_section(self, section: str, topic: str, protocol: str, results: str,
                         allowed_refs: Dict[str, RefMeta], style_note: str) -> str:
        """Generate a section text using only allowed_refs via [CITE:DOI] tags.
        If model outputs any citation not in allowed_refs, we will reject and show error.
        """
        if not self.enabled:
            return "(LLM 비활성화: OPENAI_API_KEY 설정 필요)"

        try:
            # Build a compact source list for grounding
            refs_serialized = []
            for k, m in allowed_refs.items():
                label = m.doi or m.pmid or k
                refs_serialized.append({
                    "key": k,
                    "label": label,
                    "title": m.title,
                    "journal": m.journal,
                    "year": m.year,
                })

            system = (
                "You are an evidence-based medical writing assistant for colorectal surgery/oncology. "
                "Strictly use only the provided source list and cite with [CITE:DOI] (or [CITE:pmid:ID]) tags after each claim that requires evidence. "
                "If evidence is insufficient, write 'No high-quality evidence available.' Do not invent citations. "
                "Prefer recent guidelines (NCCN/ESMO/ASCRS) when present."
            )

            user = {
                "task": f"Write the {section} for a colorectal manuscript.",
                "topic": topic,
                "study_protocol": protocol,
                "key_results": results,
                "journal_style": style_note,
                "citation_rule": "Use [CITE:DOI] tags strictly from allowed list only.",
                "allowed_sources": refs_serialized,
                "language": "Korean",
            }

            # Use minimal JSON input to reduce tokens
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
                ],
                "temperature": 0.2,
            }

            # Raw HTTP to avoid library version drift
            r = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.key}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=120,
            )
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            # Guard: verify citations
            bad = []
            for tag in re.findall(r"\[CITE:([^\]]+)\]", content or ""):
                k = tag.strip().lower()
                if k not in allowed_refs:
                    bad.append(k)
            if bad:
                return (
                    "(생성 거부) 허용되지 않은 인용 태그가 포함되어 있습니다: "
                    + ", ".join(sorted(set(bad)))
                    + "\n허용된 DOI/PMID만 사용하여 다시 생성하세요."
                )
            return content
        except Exception as e:
            return f"(LLM 오류) {e}"


# =====================
# Journal metrics loader
# =====================
@st.cache_data(show_spinner=False)
def load_journal_if_csv() -> Optional[pd.DataFrame]:
    try:
        if os.path.exists("journal_if.csv"):
            df = pd.read_csv("journal_if.csv")
            # normalize columns
            cols = {c.lower(): c for c in df.columns}
            jcol = cols.get("journal") or cols.get("title") or list(df.columns)[0]
            icol = cols.get("if") or cols.get("impact_factor") or list(df.columns)[1]
            df = df.rename(columns={jcol: "journal", icol: "if"})
            df["journal_norm"] = df["journal"].str.strip().str.lower()
            return df
    except Exception:
        pass
    return None


# =====================
# Streamlit UI
# =====================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

with st.expander("사용 지침(필독)", expanded=True):
    st.markdown(
        """
        - **허구 금지**: 인용은 오직 선택/업로드한 문헌으로 제한됩니다. 
        - **인용 태그**: 생성물에는 `[CITE:DOI]` 또는 `[CITE:pmid:ID]`가 삽입됩니다. 병합 시 자동으로 `[1]`, `[2]`로 변환됩니다.
        - **IF 정렬**: `journal_if.csv`를 제공하면 해당 IF로 정렬, 없으면 OpenAlex 지표(2-year SJR)를 선택적으로 사용합니다.
        - **업로드**: PDF는 최대 50개. 파일 내 DOI를 자동 추출 시도합니다.
        - **가이드라인**: NCCN/ESMO/ASCRS 등 DOI/PMID를 함께 추가하면 우선 인용됩니다.
        """
    )

# 1) Inputs
colA, colB = st.columns([2, 1])
with colA:
    topic = st.text_area("주제 (Topic)", height=80, placeholder="예: 국소 진행성 직장암에서 TME 전 신보조 방사선의 역할…")
    protocol = st.text_area("연구계획서 요약 (Study Protocol)", height=180)
    results_txt = st.text_area("핵심 결과 요약 (Key Results)", height=120)
with colB:
    target_journal = st.text_input("타깃 저널/스타일(선택)", placeholder="예: DCR (Diseases of the Colon & Rectum)")
    anchor_ids = st.text_area("앵커 가이드라인 DOI/PMID(콤마 구분)", placeholder="예: 10.1097/DCR.000000000000… , pmid:389…")
    use_openalex = st.checkbox("OpenAlex 지표 사용(대체 지표)", value=False)

st.divider()

# 2) PubMed search
st.subheader("2) PubMed 검색")
search_query = st.text_input(
    "검색식", 
    placeholder="예: (rectal cancer OR colorectal) AND (radiotherapy OR chemoradiation) AND guidelines[Title/Abstract]"
)
retmax = st.slider("검색 개수", 10, MAX_RESULTS, 50, step=5)
search_btn = st.button("PubMed 검색 실행")

if "search_results" not in st.session_state:
    st.session_state.search_results = []  # List[RefMeta]

if search_btn and search_query:
    with st.spinner("PubMed 검색 중…"):
        pmids = pubmed_search(search_query, retmax=retmax)
        root = pubmed_fetch_xml(pmids)
        recs = pubmed_parse_records(root)
        st.session_state.search_results = recs

# 2-1) Display + IF ranking
jif = load_journal_if_csv()

if st.session_state.search_results:
    df = pd.DataFrame([
        {
            "pmid": r.pmid,
            "doi": r.doi,
            "title": r.title,
            "journal": r.journal,
            "year": r.year,
            "url": r.url,
        }
        for r in st.session_state.search_results
    ])

    # Attach IF or OpenAlex metric
    df["IF"] = None
    if jif is not None:
        jmap = dict(zip(jif["journal_norm"], jif["if"]))
        df["IF"] = df["journal"].str.strip().str.lower().map(jmap)
    elif use_openalex:
        metrics = []
        for jn in df["journal"].fillna(""):
            metrics.append(openalex_metric(jn) if jn else None)
        df["IF"] = metrics

    st.markdown("**검색 결과**(선택하여 허용 목록에 추가)")
    sel = st.data_editor(
        df,
        hide_index=True,
        column_config={"url": st.column_config.LinkColumn("PubMed")},
        use_container_width=True,
    )

    add_sel = st.button("선택 항목을 허용 문헌으로 추가")
else:
    sel = None
    add_sel = False

# Allowed references bucket
if "allowed" not in st.session_state:
    st.session_state.allowed: Dict[str, RefMeta] = {}

if add_sel and sel is not None:
    for _, row in sel.iterrows():
        # Only add rows ticked in editor? data_editor doesn't include checkbox by default; add all visible rows.
        doi = (row.get("doi") or "").lower() or None
        pmid = str(row.get("pmid")) if row.get("pmid") else None
        key = doi or (f"pmid:{pmid}" if pmid else None)
        if not key:
            continue
        meta = RefMeta(doi=doi, pmid=pmid, title=row.get("title"), journal=row.get("journal"), year=str(row.get("year")), url=row.get("url"))
        st.session_state.allowed[key] = meta
    st.success(f"허용 문헌에 {len(st.session_state.allowed)}개 항목이 있습니다.")

st.divider()

# 3) PDF uploads
st.subheader("3) PDF 업로드(최대 50)")
pdfs = st.file_uploader("논문 PDF 업로드", type=["pdf"], accept_multiple_files=True)
add_pdfs = st.button("업로드 PDF에서 DOI 추출 후 허용 문헌에 추가")

if add_pdfs and pdfs:
    added = 0
    for up in pdfs[:MAX_UPLOADS]:
        content = up.read()
        text, doi = extract_pdf_text_and_doi(content)
        meta = None
        key = None
        if doi:
            meta = crossref_get(doi)
            key = doi.lower()
        # Fallback: try to guess title from first page (very naive)
        if not meta:
            meta = RefMeta(doi=doi)
            key = (doi.lower() if doi else None)
        if key:
            st.session_state.allowed[key] = meta
            added += 1
    st.success(f"PDF에서 {added}개 항목을 추가했습니다. (DOI 미탐지 파일은 생략됨)")

# 3-1) Anchor guideline IDs
if anchor_ids:
    added = 0
    for token in anchor_ids.split(","):
        token = token.strip()
        if not token:
            continue
        if token.lower().startswith("pmid:"):
            key = token.lower()
            st.session_state.allowed.setdefault(key, RefMeta(pmid=token.split(":",1)[1]))
            added += 1
        elif token.lower().startswith("10."):
            doi = token.lower()
            st.session_state.allowed.setdefault(doi, crossref_get(doi) or RefMeta(doi=doi))
            added += 1
    if added:
        st.info(f"앵커 문헌 {added}건 추가.")

# Allowed list view
st.markdown("### 허용 문헌 (인용 가능한 집합)")
if st.session_state.allowed:
    adf = pd.DataFrame([
        {
            "key": k,
            "doi": v.doi,
            "pmid": v.pmid,
            "title": v.title,
            "journal": v.journal,
            "year": v.year,
            "OA_link": unpaywall_best_oa_link(v.doi) if v.doi else None,
        }
        for k, v in st.session_state.allowed.items()
    ])
    st.dataframe(adf, use_container_width=True)
else:
    st.write("(아직 비어있습니다)")

st.divider()

# 4) Section-wise generation
st.subheader("4) 섹션별 생성 → 최종 병합")
llm = LLM(OPENAI_MODEL, OPENAI_API_KEY)
style_note = f"Target journal: {target_journal}" if target_journal else "Vancouver-style generic clinical paper"

SECTIONS = [
    "Cover Letter",
    "Title Page",
    "Abstract",
    "Introduction",
    "Methods",
    "Results",
    "Discussion",
]

if "sections" not in st.session_state:
    st.session_state.sections: Dict[str, str] = {}

cols = st.columns(2)
for i, sec in enumerate(SECTIONS):
    with cols[i % 2]:
        st.markdown(f"**{sec}**")
        gen_btn = st.button(f"{sec} 생성", key=f"gen_{sec}")
        if gen_btn:
            text = llm.generate_section(sec, topic, protocol, results_txt, st.session_state.allowed, style_note)
            st.session_state.sections[sec] = text
        st.text_area(f"{sec} 미리보기", value=st.session_state.sections.get(sec, ""), height=220, key=f"ta_{sec}")

st.markdown("**References** 섹션은 최종 병합 단계에서 자동 생성됩니다.")

merge_btn = st.button("최종 병합 및 번호 재정렬")

if merge_btn:
    # Build a single manuscript with reference numbering
    rm = ReferenceManager()

    # Register all allowed refs first to have meta
    for k, m in st.session_state.allowed.items():
        # ensure key shape
        if k.startswith("pmid:"):
            m.pmid = k.split(":",1)[1]
        else:
            m.doi = m.doi or k
        rm.register(m)

    def replace_citations(text: str) -> Tuple[str, List[str]]:
        # returns text with temporary placeholders and sequence of keys
        seq = []
        def _rep(m):
            tag = m.group(1).strip().lower()
            # Accept pmid:ID or DOI
            key = tag if tag.startswith("pmid:") else tag
            if key not in st.session_state.allowed:
                # leave as-is but mark
                return f"[CITE-INVALID:{tag}]"
            seq.append(key)
            n = rm.cite(key)
            return f"[{n}]"
        new = re.sub(r"\[CITE:([^\]]+)\]", _rep, text or "")
        return new, seq

    merged_parts = []
    citation_seq = []
    for sec in SECTIONS:
        txt = st.session_state.sections.get(sec, "")
        if not txt:
            continue
        rep, seq = replace_citations(txt)
        merged_parts.append(f"## {sec}\n\n" + rep.strip())
        citation_seq.extend(seq)

    # Renumber by first appearance
    rm.renumber_by_first_appearance(citation_seq)

    # After renumbering, the [n] numbers we placed already reflect cite() order;
    # We could re-run replace to ensure consistency, but current approach holds.

    # Build reference list
    refs = rm.render_reference_list()

    final_md = "\n\n".join(merged_parts) + "\n\n## References\n\n" + "\n".join(
        [f"[{i+1}] {line}" for i, line in enumerate(refs)]
    )

    st.session_state.final_md = final_md
    st.success("병합 완료 – 아래에서 미리보기/내보내기 하세요.")

# Preview + Export
if "final_md" in st.session_state:
    st.subheader("미리보기 (Markdown)")
    st.text_area("Final Markdown", value=st.session_state.final_md, height=420)

    # Export MD
    md_bytes = st.session_state.final_md.encode("utf-8")
    st.download_button("Download .md", data=md_bytes, file_name="manuscript.md", mime="text/markdown")

    # Export DOCX
    def md_to_docx(md_text: str) -> bytes:
        # Minimal: write as raw paragraphs (for MVP). Could integrate markdown->docx later.
        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = "Calibri"
        style.font.size = Pt(11)
        for block in md_text.split("\n\n"):
            doc.add_paragraph(block)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio.read()

    docx_bytes = md_to_docx(st.session_state.final_md)
    st.download_button("Download .docx", data=docx_bytes, file_name="manuscript.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.divider()

st.caption("© 2025 CRC Manuscript Builder (MVP). Evidence-locked generation. No PHI.")
