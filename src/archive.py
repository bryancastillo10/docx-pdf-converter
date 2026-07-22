from pathlib import Path
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile


def create_pdf_archive(pdf_files: Iterable[Path], archive_path: Path) -> Path:
    files = list(pdf_files)
    if not files:
        raise ValueError("There are no PDF files to archive.")

    archive_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(archive_path, "w", compression=ZIP_DEFLATED) as archive:
        for pdf_file in files:
            if not pdf_file.is_file():
                raise FileNotFoundError(pdf_file)
            archive.write(pdf_file, arcname=pdf_file.name)
    return archive_path
