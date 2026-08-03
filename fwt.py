from __future__ import annotations

import json
import logging
import re
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib import error, parse, request

import pandas as pd
import streamlit as st


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s - %(message)s")


BASE_DIR = Path(__file__).resolve().parent
USAGE_XLSX = BASE_DIR / "Foreign word transcription_Usage List.xlsx"
RAG_MD = BASE_DIR / "Foreign word transcription.rag.full.md"
OLLAMA_MODEL = "gpt-oss:20b"
OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_SEC = 180
OLLAMA_HEALTH_TIMEOUT_SEC = 3
OLLAMA_MAX_RETRIES = 2


LANGUAGE_ALIASES: Dict[str, List[str]] = {
    "영어": ["영어", "english", "en", "미국", "영국", "캐나다", "오스트레일리아", "호주"],
    "독일어": ["독일어", "german", "de", "독일", "deutsch", "독일어권"],
    "프랑스어": ["프랑스어", "french", "fr", "프랑스", "프랑스어권"],
    "에스파냐어": ["에스파냐어", "스페인어", "spanish", "es", "스페인", "에스파냐"],
    "이탈리아어": ["이탈리아어", "italian", "it", "이탈리아"],
    "일본어": ["일본어", "japanese", "ja", "jp", "일본"],
    "중국어": ["중국어", "chinese", "zh", "cn", "중국"],
    "폴란드어": ["폴란드어", "polish", "pl", "폴란드"],
    "체코어": ["체코어", "czech", "cz", "체코"],
    "세르보크로아트어": ["세르보크로아트어", "serbocroatian", "세르비아어", "크로아티아어"],
    "루마니아어": ["루마니아어", "romanian", "ro", "루마니아"],
    "헝가리어": ["헝가리어", "hungarian", "hu", "헝가리"],
    "스웨덴어": ["스웨덴어", "swedish", "sv", "스웨덴"],
    "노르웨이어": ["노르웨이어", "norwegian", "no", "노르웨이"],
    "덴마크어": ["덴마크어", "danish", "da", "덴마크"],
    "말레이인도네시아어": [
        "말레이인도네시아어",
        "말레이시아어",
        "인도네시아어",
        "말레이어",
        "malay",
        "indonesian",
        "ms",
        "id",
    ],
    "타이어": ["타이어", "thai", "th", "태국", "타이"],
    "베트남어": ["베트남어", "vietnamese", "vi", "베트남"],
    "포르투갈어": ["포르투갈어", "portuguese", "pt", "포르투갈", "브라질", "브라질포르투갈어"],
    "네덜란드어": ["네덜란드어", "dutch", "nl", "네덜란드"],
    "러시아어": ["러시아어", "russian", "ru", "러시아"],
}

LANGUAGE_SECTION_TITLES: Dict[str, str] = {
    "영어": "제1절 영어의 표기",
    "독일어": "제2절 독일어의 표기",
    "프랑스어": "제3절 프랑스어의 표기",
    "에스파냐어": "제4절 에스파냐어의 표기",
    "이탈리아어": "제5절 이탈리아어의 표기",
    "일본어": "제6절 일본어의 표기",
    "중국어": "제7절 중국어의 표기",
    "폴란드어": "제8절 폴란드어의 표기",
    "체코어": "제9절 체코어의 표기",
    "세르보크로아트어": "제10절 세르보크로아트어의 표기",
    "루마니아어": "제11절 루마니아어의 표기",
    "헝가리어": "제12절 헝가리어의 표기",
    "스웨덴어": "제13절 스웨덴어의 표기",
    "노르웨이어": "제14절 노르웨이어의 표기",
    "덴마크어": "제15절 덴마크어의 표기",
    "말레이인도네시아어": "제16절 말레이인도네시아어의 표기",
    "타이어": "제17절 타이어의 표기",
    "베트남어": "제18절 베트남어의 표기",
    "포르투갈어": "제19절 포르투갈어의 표기",
    "네덜란드어": "제20절 네덜란드어의 표기",
    "러시아어": "제21절 러시아어의 표기",
}

LANGUAGE_RULE_CONFIG: Dict[str, Dict[str, List[str]]] = {
    "영어": {"tables": ["T001"], "sections": ["제1절 영어의 표기"], "inherits": []},
    "독일어": {"tables": ["T001"], "sections": ["제2절 독일어의 표기"], "inherits": ["제1절 영어의 표기"]},
    "프랑스어": {"tables": ["T001"], "sections": ["제3절 프랑스어의 표기"], "inherits": ["제1절 영어의 표기"]},
    "에스파냐어": {"tables": ["T002"], "sections": ["제4절 에스파냐어의 표기"], "inherits": []},
    "이탈리아어": {"tables": ["T003"], "sections": ["제5절 이탈리아어의 표기"], "inherits": []},
    "일본어": {"tables": ["T004"], "sections": ["제6절 일본어의 표기"], "inherits": []},
    "중국어": {"tables": ["T005"], "sections": ["제7절 중국어의 표기"], "inherits": []},
    "폴란드어": {"tables": ["T006"], "sections": ["제8절 폴란드어의 표기"], "inherits": []},
    "체코어": {"tables": ["T007"], "sections": ["제9절 체코어의 표기"], "inherits": []},
    "세르보크로아트어": {"tables": ["T008"], "sections": ["제10절 세르보크로아트어의 표기"], "inherits": []},
    "루마니아어": {"tables": ["T009"], "sections": ["제11절 루마니아어의 표기"], "inherits": []},
    "헝가리어": {"tables": ["T010"], "sections": ["제12절 헝가리어의 표기"], "inherits": []},
    "스웨덴어": {"tables": ["T011"], "sections": ["제13절 스웨덴어의 표기"], "inherits": []},
    "노르웨이어": {"tables": ["T012"], "sections": ["제14절 노르웨이어의 표기"], "inherits": []},
    "덴마크어": {"tables": ["T013"], "sections": ["제15절 덴마크어의 표기"], "inherits": []},
    "말레이인도네시아어": {"tables": ["T014"], "sections": ["제16절 말레이인도네시아어의 표기"], "inherits": []},
    "타이어": {"tables": ["T015"], "sections": ["제17절 타이어의 표기"], "inherits": []},
    "베트남어": {"tables": ["T016"], "sections": ["제18절 베트남어의 표기"], "inherits": []},
    "포르투갈어": {"tables": ["T017"], "sections": ["제19절 포르투갈어의 표기"], "inherits": []},
    "네덜란드어": {"tables": ["T018"], "sections": ["제20절 네덜란드어의 표기"], "inherits": []},
    "러시아어": {"tables": ["T019"], "sections": ["제21절 러시아어의 표기"], "inherits": []},
}

BLOCK_PATTERN = re.compile(
    r"(?ms)^### (B\d{5}) (Paragraph|Table) \((P\d{5}|T\d{3})\)\n\n```text\n(.*?)\n```"
)
ANCHOR_PATTERN = re.compile(r"(?m)^- `(chapter|section|article|table-title)` \| `(B\d{5})` \| `(P\d{5})` \| (.+)$")
HANGUL_TRANSCRIPTION_PATTERN = re.compile(r"[가-힣·\- ]{1,40}")


@dataclass
class Recommendation:
    source: str
    language_key: Optional[str]
    recommendation: str
    reason: str
    references: List[str]
    candidates: pd.DataFrame | None = None
    llm_note: str = ""
    ipa: str = ""
    confidence: float | None = None
    needs_pronunciation_confirmation: bool = False
    error_message: str = ""


@dataclass
class RagBlock:
    block_id: str
    block_type: str
    content_id: str
    title: str
    text: str
    language: Optional[str] = None
    section_title: Optional[str] = None


@dataclass
class RagAnchor:
    kind: str
    block_id: str
    content_id: str
    title: str


@dataclass
class TableCell:
    table_id: str
    row_id: str
    col_id: str
    rowspan: int
    colspan: int
    text: str


@dataclass
class PronunciationInfo:
    language: str
    original: str
    ipa: str
    syllables: List[str]
    confidence: float
    needs_pronunciation_confirmation: bool
    source: str
    reason: str = ""


@dataclass
class RagDocument:
    anchors: List[RagAnchor]
    blocks: List[RagBlock]
    blocks_by_id: Dict[str, RagBlock]
    blocks_by_content_id: Dict[str, RagBlock]
    section_anchors: List[RagAnchor]
    section_to_anchor_index: Dict[str, int]
    table_title_by_table_id: Dict[str, str]
    table_rows_by_table_id: Dict[str, List[str]]


def validate_required_files() -> None:
    required_files = [USAGE_XLSX, RAG_MD]
    missing = [path.name for path in required_files if not path.is_file()]
    if missing:
        raise FileNotFoundError("필수 파일을 찾을 수 없습니다: " + ", ".join(missing))


def normalize_text(value: str) -> str:
    value = str(value or "").strip().lower()
    return re.sub(r"[^0-9a-zA-Z가-힣]+", "", value)


def normalize_for_match(value: str) -> str:
    value = str(value or "").strip().lower()
    return re.sub(r"\s+", "", value)


def block_num(block_id: str) -> int:
    return int(block_id[1:])


def content_num(content_id: str) -> int:
    return int(re.sub(r"^[A-Z]", "", content_id))


def resolve_language_key(user_scope: str) -> Optional[str]:
    key = normalize_text(user_scope)
    if not key:
        return None
    for canonical, aliases in LANGUAGE_ALIASES.items():
        alias_norm = [normalize_text(alias) for alias in aliases]
        if key in alias_norm:
            return canonical
        if any(alias and alias in key for alias in alias_norm):
            return canonical
    return None


def language_from_section_title(section_title: Optional[str]) -> Optional[str]:
    if not section_title:
        return None
    for language_key, title in LANGUAGE_SECTION_TITLES.items():
        if title == section_title:
            return language_key
    return None


@st.cache_data(show_spinner=False)
def load_usage_data(path: str, modified_time: float) -> pd.DataFrame:
    del modified_time
    df = pd.read_excel(path)
    df = df[df["공개 여부"].fillna("Y") == "Y"].copy()
    df["원어_정규화"] = df["원어 표기"].fillna("").astype(str).map(normalize_for_match)
    df["국명_정규화"] = df["국명"].fillna("").astype(str).map(normalize_text)
    df["언어명_정규화"] = df["언어명"].fillna("").astype(str).map(normalize_text)
    return df


@st.cache_data(show_spinner=False)
def load_rag_document(path: str, modified_time: float) -> str:
    del modified_time
    return Path(path).read_text(encoding="utf-8")


def parse_table_cells(table_block: RagBlock) -> List[TableCell]:
    cells: List[TableCell] = []
    for line in table_block.text.splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) < 6:
            continue
        tid, row_id, col_id, row_span, col_span = parts[:5]
        text = "\t".join(parts[5:]).strip()
        try:
            cells.append(
                TableCell(
                    table_id=tid.strip(),
                    row_id=row_id.strip(),
                    col_id=col_id.strip(),
                    rowspan=int(row_span),
                    colspan=int(col_span),
                    text=text,
                )
            )
        except ValueError:
            continue
    return cells


def build_table_row_meanings(cells: List[TableCell], table_title: str) -> Dict[str, List[str]]:
    by_table_row: Dict[Tuple[str, str], List[TableCell]] = {}
    for cell in cells:
        by_table_row.setdefault((cell.table_id, cell.row_id), []).append(cell)

    table_rows: Dict[str, List[str]] = {}
    for (table_id, row_id), row_cells in by_table_row.items():
        row_cells.sort(key=lambda c: int(c.col_id[1:]))
        joined = " | ".join([f"{c.col_id}:{c.text}" for c in row_cells if c.text != ""])
        if len(row_cells) >= 3 and row_cells[0].text:
            source_symbol = row_cells[0].text
            before_vowel = row_cells[1].text if len(row_cells) > 1 else ""
            before_consonant = row_cells[2].text if len(row_cells) > 2 else ""
            sentence = (
                f"[{table_id}-{row_id}] {table_title} | 기호 {source_symbol}: "
                f"모음 앞={before_vowel or '(없음)'}, 자음 앞/어말={before_consonant or '(없음)'}"
            )
        else:
            sentence = f"[{table_id}-{row_id}] {table_title} | {joined}"
        table_rows.setdefault(table_id, []).append(sentence)
    for table_id in table_rows:
        table_rows[table_id].sort(key=lambda line: int(re.search(r"R(\d+)", line).group(1)) if re.search(r"R(\d+)", line) else 0)
    return table_rows


@st.cache_data(show_spinner=False)
def load_rag_parsed(path: str, modified_time: float) -> RagDocument:
    text = load_rag_document(path, modified_time)

    anchors: List[RagAnchor] = []
    for kind, bid, pid, title in ANCHOR_PATTERN.findall(text):
        anchors.append(RagAnchor(kind=kind, block_id=bid, content_id=pid, title=title.strip()))
    anchors.sort(key=lambda a: block_num(a.block_id))

    blocks: List[RagBlock] = []
    for bid, btype, cid, body in BLOCK_PATTERN.findall(text):
        first_line = body.splitlines()[0] if body.splitlines() else ""
        payload_text = ""
        if btype == "Paragraph":
            parts = first_line.split("\t", 1)
            payload_text = parts[1].strip() if len(parts) == 2 else first_line.strip()
            title = payload_text
        else:
            payload_text = body
            title = f"{cid} table"
        blocks.append(
            RagBlock(
                block_id=bid,
                block_type=btype,
                content_id=cid,
                title=title,
                text=payload_text,
            )
        )
    blocks.sort(key=lambda b: block_num(b.block_id))

    blocks_by_id = {b.block_id: b for b in blocks}
    blocks_by_content_id = {b.content_id: b for b in blocks}

    section_anchors = [a for a in anchors if a.kind == "section"]
    section_to_anchor_index = {a.title: idx for idx, a in enumerate(section_anchors)}

    # block에 section/language 정보 부여
    for block in blocks:
        current_section = None
        for section_anchor in section_anchors:
            if block_num(section_anchor.block_id) <= block_num(block.block_id):
                current_section = section_anchor.title
            else:
                break
        block.section_title = current_section
        block.language = language_from_section_title(current_section)

    # table-title anchor와 다음 table block 연결
    table_blocks = [b for b in blocks if b.block_type == "Table"]
    table_title_by_table_id: Dict[str, str] = {}
    for anchor in [a for a in anchors if a.kind == "table-title"]:
        next_table = next((tb for tb in table_blocks if block_num(tb.block_id) > block_num(anchor.block_id)), None)
        if next_table:
            table_title_by_table_id[next_table.content_id] = anchor.title

    # 전체 표 행 의미 구조 생성
    all_table_rows: Dict[str, List[str]] = {}
    for tb in table_blocks:
        title = table_title_by_table_id.get(tb.content_id, tb.title)
        meanings = build_table_row_meanings(parse_table_cells(tb), title)
        for tid, lines in meanings.items():
            all_table_rows.setdefault(tid, []).extend(lines)

    return RagDocument(
        anchors=anchors,
        blocks=blocks,
        blocks_by_id=blocks_by_id,
        blocks_by_content_id=blocks_by_content_id,
        section_anchors=section_anchors,
        section_to_anchor_index=section_to_anchor_index,
        table_title_by_table_id=table_title_by_table_id,
        table_rows_by_table_id=all_table_rows,
    )


def section_bundle_block_ids(doc: RagDocument, section_title: str) -> List[str]:
    idx = doc.section_to_anchor_index.get(section_title)
    if idx is None:
        return []
    section_anchor = doc.section_anchors[idx]
    start = block_num(section_anchor.block_id)
    end = block_num(doc.section_anchors[idx + 1].block_id) if idx + 1 < len(doc.section_anchors) else 99999

    selected: List[str] = []
    for anchor in doc.anchors:
        n = block_num(anchor.block_id)
        if start <= n < end and anchor.kind in {"section", "table-title", "article"}:
            selected.append(anchor.block_id)
    return selected


def lookup_anchor_line(doc: RagDocument, block_id: str) -> str:
    block = doc.blocks_by_id.get(block_id)
    if not block:
        return block_id
    return f"{block.block_id}/{block.content_id} {block.title}"


def row_matches_language(row: pd.Series, language_key: str) -> bool:
    aliases = [normalize_text(a) for a in LANGUAGE_ALIASES.get(language_key, [language_key])]
    country = row.get("국명_정규화", "")
    lang = row.get("언어명_정규화", "")
    return any(a and (a in country or a in lang) for a in aliases)


def common_prefix_length(a: str, b: str) -> int:
    length = 0
    for c1, c2 in zip(a, b):
        if c1 != c2:
            break
        length += 1
    return length


def common_suffix_length(a: str, b: str) -> int:
    return common_prefix_length(a[::-1], b[::-1])


def estimate_syllable_count(word: str) -> int:
    token = normalize_for_match(word)
    if not token:
        return 0
    groups = re.findall(r"[aeiouy]+", token)
    return max(1, len(groups))


def search_similar_examples(df: pd.DataFrame, word: str, language_key: str, top_k: int = 5) -> pd.DataFrame:
    subset = df[df.apply(lambda row: row_matches_language(row, language_key), axis=1)].copy()
    if subset.empty:
        return subset

    query = normalize_for_match(word)
    q_len = max(len(query), 1)
    query_syllables = estimate_syllable_count(query)

    subset["prefix_score"] = subset["원어_정규화"].map(lambda x: common_prefix_length(query, str(x)) / q_len)
    subset["suffix_score"] = subset["원어_정규화"].map(lambda x: common_suffix_length(query, str(x)) / q_len)
    subset["spell_score"] = subset["원어_정규화"].map(lambda x: _seq_ratio(query, str(x)))
    subset["syllable_score"] = subset["원어_정규화"].map(
        lambda x: 1.0 - abs(estimate_syllable_count(str(x)) - query_syllables) / max(query_syllables, 1)
    )
    subset["score"] = (
        subset["prefix_score"] * 0.25
        + subset["suffix_score"] * 0.25
        + subset["syllable_score"] * 0.20
        + subset["spell_score"] * 0.30
    )
    subset = subset.sort_values("score", ascending=False)
    return subset.head(top_k)


def _seq_ratio(a: str, b: str) -> float:
    from difflib import SequenceMatcher

    return SequenceMatcher(None, a, b).ratio()


def extract_json_object(text: str) -> Optional[Dict[str, Any]]:
    candidate = text.strip()
    if not candidate:
        return None

    # 1) direct parse
    try:
        parsed = json.loads(candidate)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    # 2) strip markdown code fences
    fenced = re.sub(r"^```(?:json)?\s*|\s*```$", "", candidate, flags=re.IGNORECASE | re.DOTALL).strip()
    if fenced != candidate:
        try:
            parsed = json.loads(fenced)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass

    # 3) extract first balanced top-level object
    start = candidate.find("{")
    if start == -1:
        return None
    depth = 0
    in_str = False
    esc = False
    end = -1
    for i in range(start, len(candidate)):
        ch = candidate[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    end = i + 1
                    break
    if end == -1:
        return None
    fragment = candidate[start:end]
    try:
        parsed = json.loads(fragment)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        return None
    return None


def build_rag_context(
    doc: RagDocument,
    word: str,
    language_key: str,
    pronunciation: PronunciationInfo,
    similar_examples: pd.DataFrame,
) -> Tuple[List[str], Dict[str, Any]]:
    config = LANGUAGE_RULE_CONFIG.get(language_key)
    if not config:
        return [], {"table_ids": [], "sections": [], "block_count": 0, "table_row_count": 0}

    context_lines: List[str] = []
    reference_blocks: List[str] = []

    # 공통 기본 원칙 (제1장)
    for anchor in doc.anchors:
        if anchor.kind in {"chapter", "article"} and ("제1장 표기의 기본 원칙" in anchor.title or anchor.title.startswith("제1항 외래어는")):
            context_lines.append(f"{anchor.block_id}/{anchor.content_id} {anchor.title}")
            reference_blocks.append(anchor.block_id)
        if len(context_lines) >= 6:
            break

    section_titles = list(config["sections"]) + list(config.get("inherits", []))
    for section_title in section_titles:
        bundle_ids = section_bundle_block_ids(doc, section_title)
        if not bundle_ids:
            continue
        for bid in bundle_ids:
            context_lines.append(lookup_anchor_line(doc, bid))
            reference_blocks.append(bid)

    table_row_count = 0
    for table_id in config["tables"]:
        title = doc.table_title_by_table_id.get(table_id, table_id)
        context_lines.append(f"{table_id} {title}")
        rows = doc.table_rows_by_table_id.get(table_id, [])
        # 너무 길어지는 것을 막기 위해 대표 행 위주로 제한
        selected_rows = rows[:60]
        table_row_count += len(selected_rows)
        context_lines.extend(selected_rows)

    if not similar_examples.empty:
        context_lines.append("관련 기존 용례")
        for _, row in similar_examples.iterrows():
            context_lines.append(
                f"- 원어:{row['원어 표기']} / 표기:{row['한글 표기']} / 국명:{row['국명']} / 언어:{row['언어명']} / 구분:{row['구분']}"
            )

    context_lines.append(
        f"발음정보 ipa={pronunciation.ipa or '(없음)'} syllables={','.join(pronunciation.syllables) or '(없음)'} confidence={pronunciation.confidence:.2f}"
    )

    # 중복 제거
    dedup: List[str] = []
    seen = set()
    for line in context_lines:
        if line not in seen and line.strip():
            dedup.append(line)
            seen.add(line)

    return dedup, {
        "table_ids": config["tables"],
        "sections": section_titles,
        "block_count": len(set(reference_blocks)),
        "table_row_count": table_row_count,
    }


def query_ollama_json(prompt: str) -> Tuple[Optional[Dict[str, Any]], str]:
    last_error = ""
    for attempt in range(1, OLLAMA_MAX_RETRIES + 1):
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "raw": True,
            "format": "json",
            "options": {"temperature": 0.0, "num_predict": 450},
        }
        req = request.Request(
            OLLAMA_URL,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with request.urlopen(req, timeout=OLLAMA_TIMEOUT_SEC) as response:
                raw = response.read().decode("utf-8")
        except error.HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8")
            except Exception:
                detail = str(exc)
            if "not found" in detail.lower() or "model" in detail.lower():
                msg = f"모델 미설치 또는 이름 오류: {OLLAMA_MODEL}"
            else:
                msg = f"HTTP 오류: {exc.code}"
            logger.error("Ollama HTTP 오류(attempt=%s): %s", attempt, detail)
            return None, msg
        except error.URLError as exc:
            logger.error("Ollama 연결 실패(attempt=%s): %s", attempt, getattr(exc, "reason", exc))
            return None, f"Ollama 연결 실패: {getattr(exc, 'reason', exc)}"
        except TimeoutError:
            last_error = "Ollama 요청 시간이 초과됐습니다."
            logger.error("Ollama 요청 시간 초과(attempt=%s)", attempt)
            continue
        except Exception as exc:  # noqa: BLE001
            logger.error("예상하지 못한 Ollama 오류(attempt=%s): %s: %s", attempt, type(exc).__name__, exc)
            return None, f"{type(exc).__name__}: {exc}"

        try:
            envelope = json.loads(raw)
        except json.JSONDecodeError:
            logger.error("Ollama envelope JSON 디코딩 오류(attempt=%s)", attempt)
            last_error = "Ollama 응답 JSON 디코딩에 실패했습니다."
            continue

        response_text = envelope.get("response")
        if response_text is None:
            logger.error("Ollama 응답 필드 누락(attempt=%s): response", attempt)
            last_error = "Ollama 응답 필드(response)가 누락됐습니다."
            continue
        if not str(response_text).strip():
            logger.error("Ollama 빈 응답(attempt=%s)", attempt)
            last_error = "Ollama가 빈 응답을 반환했습니다."
            continue

        parsed = extract_json_object(str(response_text))
        if parsed is None:
            preview = str(response_text)[:280].replace("\n", " ")
            logger.error("Ollama 출력 형식(JSON) 오류(attempt=%s): %s", attempt, preview)
            last_error = "모델 출력 형식(JSON)이 올바르지 않습니다."
            continue
        return parsed, ""

    return None, last_error or "Ollama 호출 재시도 후에도 실패했습니다."


def check_ollama_server() -> Tuple[bool, str]:
    try:
        parsed_url = parse.urlparse(OLLAMA_URL)
        tags_url = f"{parsed_url.scheme}://{parsed_url.netloc}/api/tags"
        req = request.Request(tags_url, method="GET")
        with request.urlopen(req, timeout=OLLAMA_HEALTH_TIMEOUT_SEC) as response:
            if response.status != 200:
                return False, f"Ollama 상태 확인 실패(HTTP {response.status})"
            body = json.loads(response.read().decode("utf-8"))
        models = body.get("models", [])
        if not any(str(model.get("name", "")).startswith(OLLAMA_MODEL) for model in models):
            return False, f"Ollama는 실행 중이지만 모델 `{OLLAMA_MODEL}`을 찾지 못했습니다."
        return True, ""
    except error.URLError as exc:
        return False, f"Ollama 서버 연결 실패: {getattr(exc, 'reason', exc)}"
    except Exception as exc:  # noqa: BLE001
        return False, f"Ollama 상태 확인 오류: {type(exc).__name__}: {exc}"


def ensure_ollama_server_running() -> Tuple[bool, str]:
    ok, msg = check_ollama_server()
    if ok:
        return True, ""

    try:
        subprocess.Popen(
            ["ollama", "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
    except Exception as exc:  # noqa: BLE001
        logger.error("Ollama 자동 시작 실패: %s: %s", type(exc).__name__, exc)
        return False, f"Ollama 자동 시작 실패: {type(exc).__name__}: {exc}"

    for _ in range(8):
        time.sleep(0.8)
        ok, msg = check_ollama_server()
        if ok:
            logger.info("Ollama 서버 자동 시작 성공")
            return True, ""

    logger.error("Ollama 자동 시작 후에도 미응답: %s", msg)
    return False, f"Ollama 자동 시작 후에도 연결 실패: {msg}"


def call_pronunciation_llm(word: str, language_key: str, context_lines: List[str]) -> Tuple[Optional[PronunciationInfo], str]:
    prompt = f"""
역할: 외래어 발음 보조기.
반드시 JSON 객체 하나만 출력한다. 코드블록 금지.
규칙 정보는 제공된 RAG 문맥만 사용한다.
발음을 확신하기 어려우면 needs_pronunciation_confirmation=true로 둔다.

입력:
- language: {language_key}
- original: {word}
- rag_context:
{chr(10).join(context_lines[:120])}

출력 JSON 스키마:
{{
  "language": "언어명",
  "original": "원어",
  "ipa": "IPA 문자열",
  "syllables": ["음절1","음절2"],
  "confidence": 0.0,
  "needs_pronunciation_confirmation": false
}}
""".strip()

    compact_prompt = prompt if len(prompt) <= 12000 else prompt[:12000]
    parsed, err = query_ollama_json(compact_prompt)
    if err:
        return None, err
    if parsed is None:
        return None, "발음 JSON 응답이 없습니다."

    try:
        info = PronunciationInfo(
            language=str(parsed.get("language", language_key)),
            original=str(parsed.get("original", word)),
            ipa=str(parsed.get("ipa", "")).strip(),
            syllables=[str(s) for s in parsed.get("syllables", []) if str(s).strip()],
            confidence=float(parsed.get("confidence", 0.0)),
            needs_pronunciation_confirmation=bool(parsed.get("needs_pronunciation_confirmation", False)),
            source="llm",
        )
    except (TypeError, ValueError):
        logger.exception("발음 추정 JSON 파싱 실패")
        return None, "발음 추정 JSON 파싱에 실패했습니다."
    return info, ""


def call_transcription_llm(
    word: str,
    scope: str,
    language_key: str,
    pronunciation: PronunciationInfo,
    context_lines: List[str],
) -> Tuple[Optional[Dict[str, Any]], str]:
    prompt = f"""
당신은 외래어 한글 표기 추천기다.
다음 단계로 작업한다.
1. 입력 언어를 확인한다.
2. 원어의 발음 또는 IPA를 확인한다.
3. 기존 용례에 정확히 일치하는지 확인한다.
4. 관련 대조표의 음소 대응을 확인한다.
5. 해당 언어의 표기 세칙을 확인한다.
6. 어두·어중·어말 환경을 적용한다.
7. 관용 표기가 있는지 확인한다.
8. 최종 한글 표기 하나를 결정한다.
9. 실제 제공된 블록 ID만 인용한다.

제약:
- 반드시 JSON 객체 하나만 출력한다.
- 마크다운 코드 블록을 사용하지 않는다.
- RAG 문맥에 없는 조항 ID를 생성하지 않는다.
- 규칙 정보는 제공된 RAG 문맥만 사용한다.
- 발음 정보가 제공되지 않은 경우 발음을 추정할 수 있다.
- 추정 발음은 추정이라고 reason에 명시한다.
- 발음에 확신이 낮더라도 가능한 경우 best-effort transcription을 제시하고
  needs_pronunciation_confirmation=true로 둔다.
- 규칙 근거가 전혀 부족할 때만 transcription을 빈 문자열로 둔다.

입력:
- scope: {scope}
- language: {language_key}
- original: {word}
- ipa: {pronunciation.ipa}
- syllables: {', '.join(pronunciation.syllables)}
- pronunciation_confidence: {pronunciation.confidence:.2f}

RAG 문맥:
{chr(10).join(context_lines[:180])}

출력 JSON 스키마:
{{
  "transcription": "한글 표기",
  "reason": "적용 근거",
  "citations": [
    {{
      "block_id": "B00000",
      "content_id": "P00000",
      "description": "설명"
    }}
  ],
  "ipa": "",
  "confidence": 0.0,
  "needs_pronunciation_confirmation": false
}}
""".strip()

    # 프롬프트가 길수록 JSON 준수율이 떨어져서 너무 긴 문맥은 제한한다.
    compact_prompt = prompt
    if len(compact_prompt) > 16000:
        compact_prompt = compact_prompt[:16000]
    parsed, err = query_ollama_json(compact_prompt)
    if err:
        return None, err
    if parsed is None:
        return None, "표기 JSON 응답이 없습니다."
    return parsed, ""


def validate_transcription_payload(payload: Dict[str, Any], doc: RagDocument) -> Tuple[bool, str, List[str]]:
    transcription = str(payload.get("transcription", "")).strip()
    if transcription and not HANGUL_TRANSCRIPTION_PATTERN.fullmatch(transcription):
        return False, "transcription 형식이 유효하지 않습니다.", []

    references: List[str] = []
    citations = payload.get("citations", [])
    if not isinstance(citations, list):
        return False, "citations 형식이 올바르지 않습니다.", []

    for citation in citations:
        if not isinstance(citation, dict):
            continue
        bid = str(citation.get("block_id", "")).strip()
        cid = str(citation.get("content_id", "")).strip()
        if not re.fullmatch(r"B\d{5}", bid):
            continue
        block = doc.blocks_by_id.get(bid)
        if not block:
            continue
        if cid and cid != block.content_id:
            continue
        desc = str(citation.get("description", "")).strip()
        references.append(f"{bid}/{block.content_id} {block.title}" + (f" - {desc}" if desc else ""))

    return True, "", references


def infer_pronunciation(
    word: str,
    language_key: str,
    pronunciation_input: str,
    doc: RagDocument,
) -> Tuple[PronunciationInfo, str]:
    if pronunciation_input.strip():
        return (
            PronunciationInfo(
                language=language_key,
                original=word,
                ipa=pronunciation_input.strip(),
                syllables=[],
                confidence=1.0,
                needs_pronunciation_confirmation=False,
                source="user",
                reason="사용자 발음 입력 사용",
            ),
            "",
        )

    base_context = [f"{a.block_id}/{a.content_id} {a.title}" for a in doc.anchors[:30]]
    pron, err = call_pronunciation_llm(word, language_key, base_context)
    if err or pron is None:
        return (
            PronunciationInfo(
                language=language_key,
                original=word,
                ipa="",
                syllables=[],
                confidence=0.0,
                needs_pronunciation_confirmation=True,
                source="llm",
                reason="발음 추정 실패",
            ),
            err or "발음 추정 실패",
        )
    return pron, ""


def recommend(word: str, scope: str, pronunciation_input: str = "") -> Recommendation:
    query = normalize_for_match(word)
    if not query:
        return Recommendation(
            source="입력 오류",
            language_key=None,
            recommendation="",
            reason="외래어 문자열을 입력해 주세요.",
            references=[],
        )

    language_key = resolve_language_key(scope)
    if language_key is None:
        return Recommendation(
            source="언어권 확인 필요",
            language_key=None,
            recommendation="",
            reason="입력 언어권을 판정할 수 없습니다. 언어명 또는 국명을 더 구체적으로 입력해 주세요.",
            references=[],
        )

    usage_df = load_usage_data(str(USAGE_XLSX), USAGE_XLSX.stat().st_mtime)
    rag_doc = load_rag_parsed(str(RAG_MD), RAG_MD.stat().st_mtime)

    # 1) 용례집 정확 일치
    exact = usage_df[(usage_df["원어_정규화"] == query) & usage_df.apply(lambda r: row_matches_language(r, language_key), axis=1)].copy()
    if not exact.empty:
        row = exact.iloc[0]
        refs = [lookup_anchor_line(rag_doc, bid) for bid in section_bundle_block_ids(rag_doc, LANGUAGE_SECTION_TITLES[language_key])[:6]]
        return Recommendation(
            source="용례집 정확 일치",
            language_key=language_key,
            recommendation=str(row["한글 표기"]),
            reason="Foreign word transcription_Usage List.xlsx의 같은 언어권 공개 용례와 정확히 일치합니다.",
            references=refs,
            candidates=exact.head(5)[["원어 표기", "한글 표기", "국명", "언어명", "의미"]],
        )

    # 2) 유사 용례 검색 (해당 언어만)
    similar = search_similar_examples(usage_df, word, language_key, top_k=5)

    ollama_ok, ollama_status = check_ollama_server()
    if not ollama_ok:
        auto_ok, auto_msg = ensure_ollama_server_running()
        if auto_ok:
            ollama_ok, ollama_status = check_ollama_server()
        else:
            ollama_status = auto_msg or ollama_status

    if not ollama_ok:
        return Recommendation(
            source="Ollama 호출 실패",
            language_key=language_key,
            recommendation="",
            reason="Ollama 서버 또는 모델 상태 문제로 표기를 생성할 수 없습니다.",
            references=[],
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            llm_note=f"모델: {OLLAMA_MODEL}",
            error_message=ollama_status,
        )

    # 3) 발음/IPA 확인 (선택값: 없으면 추정 기반으로 계속 진행)
    pronunciation, pron_err = infer_pronunciation(word, language_key, pronunciation_input, rag_doc)
    pronunciation_uncertain = pronunciation.needs_pronunciation_confirmation or pronunciation.confidence < 0.55

    # 4) RAG 문맥 구축
    rag_context, rag_meta = build_rag_context(rag_doc, word, language_key, pronunciation, similar)
    logger.info(
        "RAG 검색 진단 | 원어=%s | 입력언어=%s | 판정언어=%s | 대상표=%s | 대상절=%s | 블록수=%s | 표행수=%s | 문맥길이=%s",
        word,
        scope,
        language_key,
        ",".join(rag_meta["table_ids"]),
        ",".join(rag_meta["sections"]),
        rag_meta["block_count"],
        rag_meta["table_row_count"],
        len(rag_context),
    )
    if not rag_context:
        return Recommendation(
            source="RAG 검색 실패",
            language_key=language_key,
            recommendation="",
            reason="입력 언어에 해당하는 외래어 표기 규칙을 RAG 문서에서 찾지 못했습니다.",
            references=[],
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            ipa=pronunciation.ipa,
            confidence=pronunciation.confidence,
        )

    # 5) LLM 호출(JSON)
    payload, llm_err = call_transcription_llm(word, scope, language_key, pronunciation, rag_context)
    if llm_err or payload is None:
        return Recommendation(
            source="Ollama 호출 실패",
            language_key=language_key,
            recommendation="",
            reason="모델 호출에 실패해 표기를 확정하지 않았습니다.",
            references=[],
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            llm_note=f"모델: {OLLAMA_MODEL}",
            ipa=pronunciation.ipa,
            confidence=pronunciation.confidence,
            error_message=llm_err or pron_err,
        )

    valid, payload_err, refs = validate_transcription_payload(payload, rag_doc)
    if not valid:
        return Recommendation(
            source="표기 확정 실패",
            language_key=language_key,
            recommendation="",
            reason="모델 응답 형식이 유효하지 않아 표기를 확정하지 않았습니다.",
            references=[],
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            llm_note=f"모델: {OLLAMA_MODEL}",
            ipa=str(payload.get("ipa", pronunciation.ipa)),
            confidence=float(payload.get("confidence", pronunciation.confidence or 0.0)),
            error_message=payload_err,
        )

    needs_pron = bool(payload.get("needs_pronunciation_confirmation", False))
    transcription = str(payload.get("transcription", "")).strip()
    llm_confidence = float(payload.get("confidence", pronunciation.confidence or 0.0))
    reason = str(payload.get("reason", "")).strip()
    ipa = str(payload.get("ipa", pronunciation.ipa)).strip()

    if needs_pron and not transcription:
        return Recommendation(
            source="발음 확인 필요",
            language_key=language_key,
            recommendation="",
            reason=reason or "발음 불확실로 표기를 확정하지 않았습니다.",
            references=refs,
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            llm_note=f"모델: {OLLAMA_MODEL}",
            ipa=ipa,
            confidence=llm_confidence,
            needs_pronunciation_confirmation=True,
        )

    if not transcription:
        return Recommendation(
            source="표기 생성 실패",
            language_key=language_key,
            recommendation="",
            reason="원어 발음 또는 관련 규칙을 충분히 확인할 수 없어 표기를 확정하지 않았습니다.",
            references=refs,
            candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
            llm_note=f"모델: {OLLAMA_MODEL}",
            ipa=ipa,
            confidence=llm_confidence,
        )

    source_name = "RAG 기반 신규 표기(발음 추정)" if (needs_pron or pronunciation_uncertain) else "RAG 기반 신규 표기"
    return Recommendation(
        source=source_name,
        language_key=language_key,
        recommendation=transcription,
        reason=reason or "RAG 규칙과 입력/추정 발음을 바탕으로 표기를 생성했습니다.",
        references=refs or rag_context[:10],
        candidates=similar[["원어 표기", "한글 표기", "국명", "언어명", "score"]] if not similar.empty else None,
        llm_note=f"모델: {OLLAMA_MODEL}",
        ipa=ipa,
        confidence=llm_confidence,
        needs_pronunciation_confirmation=bool(needs_pron or pronunciation_uncertain),
        error_message=pron_err if pronunciation_uncertain else "",
    )


def render_ui() -> None:
    st.set_page_config(page_title="외래어 표기 추천 시스템", layout="wide")
    st.title("외래어 한글 표기 추천 시스템")

    try:
        validate_required_files()
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()

    st.caption("용례집 정확 일치 우선. 신규 단어는 RAG 규칙 + 로컬 gpt-oss:20b를 이용합니다.")
    ollama_ok, ollama_status = check_ollama_server()
    if not ollama_ok:
        auto_ok, auto_msg = ensure_ollama_server_running()
        if auto_ok:
            ollama_ok, ollama_status = check_ollama_server()
        else:
            ollama_status = auto_msg or ollama_status

    if not ollama_ok:
        st.warning(
            "Ollama 연결 상태가 비정상입니다. 신규 단어 생성은 실패할 수 있습니다. "
            "터미널에서 `ollama serve` 실행 여부를 확인해 주세요."
        )
        st.caption(f"진단: {ollama_status}")

    with st.form("recommend_form"):
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            foreign_word = st.text_input("외래어 문자열", placeholder="예: champagne")
        with c2:
            scope = st.text_input("언어권(국명/언어명)", placeholder="예: 프랑스어, 프랑스 파리")
        with c3:
            pronunciation = st.text_input("발음/IPA (선택)", placeholder="예: ʃɑ̃paɲ")
        submit = st.form_submit_button("표기 추천", use_container_width=True)

    if not submit:
        st.info("외래어 문자열과 언어권을 입력해 주세요.")
        return

    with st.spinner("추천 파이프라인 실행 중..."):
        rec = recommend(foreign_word, scope, pronunciation)

    left, right = st.columns([2, 1])
    with left:
        st.subheader("추천 결과")
        if rec.recommendation:
            st.success(f"추천 표기: **{rec.recommendation}**")
        elif rec.source in {"Ollama 호출 실패", "RAG 검색 실패"}:
            st.error(rec.reason)
        else:
            st.warning(rec.reason)
        st.write(f"- 상태: `{rec.source}`")
        st.write(f"- 판정 언어: `{rec.language_key}`")
        if rec.ipa:
            st.write(f"- IPA/발음: `{rec.ipa}`")
        if rec.confidence is not None:
            st.write(f"- 신뢰도: `{rec.confidence:.2f}`")
        if rec.needs_pronunciation_confirmation:
            st.write("- 발음 확인 필요: `True`")
        if rec.llm_note:
            st.write(f"- 모델 정보: {rec.llm_note}")
        if rec.error_message:
            st.error(f"진단 정보: {rec.error_message}")
        if rec.candidates is not None and not rec.candidates.empty:
            st.markdown("#### 참고 용례")
            st.dataframe(rec.candidates, use_container_width=True, hide_index=True)

    with right:
        st.subheader("참조 조항")
        if rec.references:
            for ref in rec.references[:30]:
                st.markdown(f"- {ref}")
        else:
            st.caption("표시할 참조 조항이 없습니다.")

    st.markdown("---")
    st.caption("실행: `streamlit run fwt.py`")


if __name__ == "__main__":
    render_ui()
