from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

import fwt


def _rag_doc() -> fwt.RagDocument:
    return fwt.load_rag_parsed(str(fwt.RAG_MD), fwt.RAG_MD.stat().st_mtime)


def _usage_df() -> pd.DataFrame:
    return fwt.load_usage_data(str(fwt.USAGE_XLSX), fwt.USAGE_XLSX.stat().st_mtime)


def test_validate_required_files_ok() -> None:
    fwt.validate_required_files()


def test_validate_required_files_missing_rag(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(fwt, "USAGE_XLSX", tmp_path / "ok.xlsx")
    monkeypatch.setattr(fwt, "RAG_MD", tmp_path / "missing.rag.full.md")
    (tmp_path / "ok.xlsx").write_bytes(b"dummy")
    with pytest.raises(FileNotFoundError) as exc:
        fwt.validate_required_files()
    assert "missing.rag.full.md" in str(exc.value)


def test_validate_required_files_missing_xlsx(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(fwt, "USAGE_XLSX", tmp_path / "missing.xlsx")
    monkeypatch.setattr(fwt, "RAG_MD", tmp_path / "ok.rag.full.md")
    (tmp_path / "ok.rag.full.md").write_text("ok", encoding="utf-8")
    with pytest.raises(FileNotFoundError) as exc:
        fwt.validate_required_files()
    assert "missing.xlsx" in str(exc.value)


def test_rag_parse_core_ids() -> None:
    doc = _rag_doc()
    assert "B03028" in doc.blocks_by_id
    assert "P03009" in doc.blocks_by_content_id
    assert "T001" in doc.table_rows_by_table_id
    english_bundle = fwt.section_bundle_block_ids(doc, "제1절 영어의 표기")
    assert "B03030" in english_bundle


def test_language_resolution_cases() -> None:
    assert fwt.resolve_language_key("프랑스어") == "프랑스어"
    assert fwt.resolve_language_key("프랑스 파리") == "프랑스어"
    assert fwt.resolve_language_key("독일어권") == "독일어"
    assert fwt.resolve_language_key("브라질 포르투갈어") == "포르투갈어"
    assert fwt.resolve_language_key("말레이시아어") == "말레이인도네시아어"
    assert fwt.resolve_language_key("인도네시아어") == "말레이인도네시아어"
    assert fwt.resolve_language_key("알 수 없는 언어") is None


def test_rag_context_search_rules() -> None:
    doc = _rag_doc()
    examples = _usage_df().head(3).copy()
    pron = fwt.PronunciationInfo(
        language="영어",
        original="test",
        ipa="tɛst",
        syllables=["test"],
        confidence=1.0,
        needs_pronunciation_confirmation=False,
        source="user",
    )
    english_context, _ = fwt.build_rag_context(doc, "test", "영어", pron, examples)
    german_context, _ = fwt.build_rag_context(doc, "test", "독일어", pron, examples)
    french_context, _ = fwt.build_rag_context(doc, "test", "프랑스어", pron, examples)
    spanish_context, _ = fwt.build_rag_context(doc, "test", "에스파냐어", pron, examples)

    assert any("T001" in line for line in english_context)
    assert any("제1절 영어의 표기" in line for line in english_context)
    assert any("T001" in line for line in german_context)
    assert any("제1절 영어의 표기" in line for line in german_context)
    assert any("제1절 영어의 표기" in line for line in french_context)
    assert any("T002" in line for line in spanish_context)


def test_rag_empty_skips_model(monkeypatch: pytest.MonkeyPatch) -> None:
    called = {"llm": False}

    def fake_build(*args, **kwargs):  # type: ignore[no-untyped-def]
        return [], {"table_ids": [], "sections": [], "block_count": 0, "table_row_count": 0}

    def fake_llm(*args, **kwargs):  # type: ignore[no-untyped-def]
        called["llm"] = True
        return {"transcription": "실패"}, ""

    monkeypatch.setattr(fwt, "build_rag_context", fake_build)
    monkeypatch.setattr(fwt, "check_ollama_server", lambda: (True, ""))
    monkeypatch.setattr(fwt, "call_transcription_llm", fake_llm)
    rec = fwt.recommend("newwordthatdoesnotexist", "영어", pronunciation_input="njuːwɜːd")
    assert rec.source == "RAG 검색 실패"
    assert called["llm"] is False


def test_ollama_failure_no_fake_korean(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_llm(*args, **kwargs):  # type: ignore[no-untyped-def]
        return None, "Ollama 연결 실패"

    monkeypatch.setattr(fwt, "check_ollama_server", lambda: (True, ""))
    monkeypatch.setattr(fwt, "call_transcription_llm", fake_llm)
    rec = fwt.recommend("brandnewword", "영어", pronunciation_input="brænd.njuː.wɜːd")
    assert rec.source == "Ollama 호출 실패"
    assert rec.recommendation == ""


def test_invalid_json_payload_no_fallback_extraction(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_llm(*args, **kwargs):  # type: ignore[no-untyped-def]
        return {"reason": "설명만 있고 표기 없음", "citations": []}, ""

    monkeypatch.setattr(fwt, "check_ollama_server", lambda: (True, ""))
    monkeypatch.setattr(fwt, "call_transcription_llm", fake_llm)
    rec = fwt.recommend("brandnewword2", "영어", pronunciation_input="brænd.njuː.wɜːd")
    assert rec.recommendation == ""
    assert rec.source in {"표기 생성 실패", "표기 확정 실패"}


def test_no_copy_from_similar_example(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_llm(*args, **kwargs):  # type: ignore[no-untyped-def]
        return None, "Ollama 실패"

    monkeypatch.setattr(fwt, "check_ollama_server", lambda: (True, ""))
    monkeypatch.setattr(fwt, "call_transcription_llm", fake_llm)
    rec = fwt.recommend("hoabanx", "베트남어", pronunciation_input="ho.a.ban")
    assert rec.recommendation == ""
    assert rec.source == "Ollama 호출 실패"


def test_extract_json_object_recovers_from_wrapped_text() -> None:
    text = """설명 문장\n```json\n{\"transcription\":\"음바페\",\"reason\":\"규칙\",\"citations\":[]}\n```\n추가문장"""
    parsed = fwt.extract_json_object(text)
    assert parsed is not None
    assert parsed["transcription"] == "음바페"
