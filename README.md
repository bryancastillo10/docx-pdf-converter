# Office to PDF Converter

A simple desktop application for converting Microsoft Word documents
(`.doc` and `.docx`) and PowerPoint presentations (`.ppt` and `.pptx`) into
PDF files.

## Features

- Convert multiple Word documents and PowerPoint presentations in one operation.
- Use the matching Microsoft Office app or LibreOffice as the conversion backend.
- Switch the interface between English and Traditional Chinese (`ZH-TW`).
- Select a custom output folder.
- Display conversion progress and per-file results.
- Optionally package converted PDFs into a ZIP archive.

## Requirements

- Python 3.11 or newer
- Microsoft Word on Windows/macOS for Word files
- Microsoft PowerPoint on Windows for PowerPoint files
- Or LibreOffice on any supported platform for both file types
- Dependencies listed in `requirements.txt`

## Project Structure

```text
docx-pdf-converter/
├── assets/
│   └── app.ico           # Application icon
├── src/
│   ├── archive.py        # ZIP archive creation
│   ├── converter.py      # Word/PowerPoint-to-PDF conversion
│   ├── gui.py            # Tkinter user interface
│   ├── i18n.py           # EN and ZH-TW translations
│   └── models.py         # Conversion result model
├── tests/
│   ├── test_archive.py
│   ├── test_converter.py
│   └── test_i18n.py
├── LICENSE.md
├── main.py               # Application entry point
├── main.spec             # Generated PyInstaller configuration
└── requirements.txt
```

## Development Setup

Create a virtual environment and install the dependencies:

```powershell
python -m venv docpdfenv
.\docpdfenv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

## Build with PyInstaller

Run the following command from the project root:

```powershell
pyinstaller main.spec
```

The packaged application will be created under
`dist/office-to-pdf-converter/`.

## License

This project is available under the MIT License and is credited to
Bryan Castillo. See [LICENSE.md](LICENSE.md).
