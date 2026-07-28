import os
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Callable, Iterable

from src.models import ConversionResult

SUPPORTED_EXTENSIONS = {".doc", ".docx"}


def _find_libreoffice() -> str | None:
    for command in ("libreoffice", "soffice"):
        executable = shutil.which(command)
        if executable:
            return executable

    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        base = os.environ.get(variable)
        if base:
            candidate = Path(base) / "LibreOffice" / "program" / "soffice.exe"
            if candidate.is_file():
                return str(candidate)
    return None


def _convert_with_libreoffice(source: Path, output_dir: Path, executable: str) -> None:
    process = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(output_dir),
            str(source),
        ],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    if process.returncode != 0:
        details = (
            process.stderr.strip()
            or process.stdout.strip()
            or "Unknown LibreOffice error"
        )
        raise RuntimeError(f"LibreOffice conversion failed: {details}")


def _convert_with_word(source: Path, output: Path) -> None:
    if platform.system() == "Darwin":
        from docx2pdf import convert

        convert(str(source), str(output))
        return

    # docx2pdf wraps this same Word COM operation in a tqdm progress bar.
    import pythoncom
    import win32com.client

    word = None
    document = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(source.resolve()), ReadOnly=True)
        document.SaveAs(str(output.resolve()), FileFormat=17)  # wdFormatPDF
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def convert_document(source: Path, output_dir: Path) -> Path:
    """Convert a Word document and preserve its filename stem."""
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {source.suffix}")
    if not source.is_file():
        raise FileNotFoundError(source)

    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{source.stem}.pdf"
    errors: list[str] = []

    if platform.system() in {"Windows", "Darwin"}:
        try:
            _convert_with_word(source, output)
        except Exception as exc:
            errors.append(f"Microsoft Word: {str(exc) or type(exc).__name__}")
        else:
            if output.is_file():
                return output
            errors.append("Microsoft Word finished without creating a PDF")

    libreoffice = _find_libreoffice()
    if libreoffice:
        try:
            _convert_with_libreoffice(source, output_dir, libreoffice)
        except Exception as exc:
            errors.append(str(exc) or type(exc).__name__)
        else:
            if output.is_file():
                return output
            errors.append("LibreOffice finished without creating a PDF")
    else:
        errors.append("LibreOffice was not found")

    raise RuntimeError(
        "No conversion backend succeeded. "
        + " | ".join(errors)
        + ". Install Microsoft Word on Windows/macOS or LibreOffice on any supported platform."
    )


def convert_documents(
    sources: Iterable[Path],
    output_dir: Path,
    on_result: Callable[[ConversionResult, int, int], None] | None = None,
) -> list[ConversionResult]:
    source_list = list(sources)
    results: list[ConversionResult] = []

    for index, source in enumerate(source_list, start=1):
        try:
            result = ConversionResult(
                source=source, output=convert_document(source, output_dir)
            )
        except Exception as exc:
            result = ConversionResult(
                source=source, error=str(exc) or type(exc).__name__
            )
        results.append(result)
        if on_result:
            on_result(result, index, len(source_list))

    return results
