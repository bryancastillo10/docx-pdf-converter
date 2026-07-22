from pathlib import Path

import src.converter as converter


def test_convert_documents_continues_after_failure(monkeypatch, tmp_path: Path) -> None:
    sources = [tmp_path / "good.docx", tmp_path / "bad.docx"]

    def fake_convert(source: Path, output_dir: Path) -> Path:
        if source.name == "bad.docx":
            raise RuntimeError("conversion failed")
        return output_dir / "good.pdf"

    monkeypatch.setattr(converter, "convert_document", fake_convert)
    results = converter.convert_documents(sources, tmp_path)

    assert results[0].succeeded
    assert not results[1].succeeded
    assert results[1].error == "conversion failed"
