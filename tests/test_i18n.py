from src.i18n import translate


def test_translate_english_with_values() -> None:
    assert translate("en", "files_selected", count=2) == "2 file(s) selected."


def test_translate_traditional_chinese_with_values() -> None:
    assert translate("zh-TW", "files_selected", count=2) == "已選取 2 個檔案。"


def test_unknown_language_falls_back_to_english() -> None:
    assert translate("unknown", "convert") == "Convert"
