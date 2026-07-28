from pathlib import Path
import sys
from types import ModuleType, SimpleNamespace

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


def test_word_conversion_uses_com_without_console_streams(monkeypatch, tmp_path: Path) -> None:
    calls: list[object] = []

    class FakeDocument:
        def SaveAs(self, path: str, FileFormat: int) -> None:
            calls.append(("save", path, FileFormat))

        def Close(self, save_changes: bool) -> None:
            calls.append(("close", save_changes))

    class FakeWord:
        def __init__(self) -> None:
            self.Documents = SimpleNamespace(
                Open=lambda path, ReadOnly: calls.append(("open", path, ReadOnly))
                or FakeDocument()
            )

        def Quit(self) -> None:
            calls.append("quit")

    pythoncom = ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: calls.append("initialize")
    pythoncom.CoUninitialize = lambda: calls.append("uninitialize")
    win32com = ModuleType("win32com")
    win32com.client = SimpleNamespace(
        DispatchEx=lambda name: calls.append(("dispatch", name)) or FakeWord()
    )

    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", win32com.client)

    source = tmp_path / "input.docx"
    output = tmp_path / "output.pdf"
    converter._convert_with_word(source, output)

    assert calls == [
        "initialize",
        ("dispatch", "Word.Application"),
        ("open", str(source.resolve()), True),
        ("save", str(output.resolve()), 17),
        ("close", False),
        "quit",
        "uninitialize",
    ]
