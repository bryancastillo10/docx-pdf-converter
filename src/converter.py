from pathlib import Path
from typing import Callable, Iterable

from src.models import ConversionResult

SUPPORTED_EXTENSIONS = {".doc", ".docx"}


def convert_document(source: Path, output_dir: Path) -> Path:
    """Convert one Word document with the locally installed Microsoft Word."""
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {source.suffix}")
    if not source.is_file():
        raise FileNotFoundError(source)

    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{source.stem}.pdf"

    # Import lazily so the GUI can still show a useful conversion error.
    from docx2pdf import convert

    convert(str(source), str(output))
    if not output.is_file():
        raise RuntimeError("The conversion finished without creating a PDF.")
    return output


def convert_documents(
    sources: Iterable[Path],
    output_dir: Path,
    on_result: Callable[[ConversionResult, int, int], None] | None = None,
) -> list[ConversionResult]:
    source_list = list(sources)
    results: list[ConversionResult] = []

    for index, source in enumerate(source_list, start=1):
        try:
            result = ConversionResult(source=source, output=convert_document(source, output_dir))
        except Exception as exc:  # Keep the rest of the batch running.
            result = ConversionResult(source=source, error=str(exc) or type(exc).__name__)
        results.append(result)
        if on_result:
            on_result(result, index, len(source_list))

    return results
