from pathlib import Path
from zipfile import ZipFile

from src.archive import create_pdf_archive


def test_create_pdf_archive(tmp_path: Path) -> None:
    first = tmp_path / "first.pdf"
    second = tmp_path / "second.pdf"
    first.write_bytes(b"first")
    second.write_bytes(b"second")

    result = create_pdf_archive([first, second], tmp_path / "documents.zip")

    assert result.is_file()
    with ZipFile(result) as archive:
        assert archive.namelist() == ["first.pdf", "second.pdf"]
