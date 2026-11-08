# PDF to Word - offline, Bengali / Hindi / English

Converts PDFs (including 100 MB+ files) to Word (.docx) with exact text, real tables, equations
and graphs placed where they are on the page. Fully offline.

## Option 1 - Windows app (no Python needed)
Download `PDFtoWord-windows.zip` from the **Releases** page (or the latest run under **Actions**),
unzip, and run `PDFtoWord.exe`. Tesseract and the Bengali/Hindi/English OCR data are inside.

## Option 2 - Run from source
```
pip install -r requirements.txt
python download_tessdata.py        # one-time, ~38 MB OCR data for ben/hin/eng
python app.py                      # window
python app.py book.pdf -o out.docx # command line
```
Scanned PDFs also need the **Tesseract program** (Windows: UB-Mannheim installer, Linux:
`apt install tesseract-ocr`, macOS: `brew install tesseract`). Digital PDFs do not.

## How it works
- Digital PDFs: exact text from the PDF's text layer (bold/italic/size/colour kept).
- Scanned pages: Tesseract OCR with Bengali + Hindi + English, plus the original page image under
  the text so graphs, equations and tables are never lost. Lines that come out half Bengali / half
  Hindi are detected and re-read with the right script.
- Tables become real Word tables; equations are exact pictures with copyable text; graphs are
  pictures at their position.
- Big files: one page at a time, output split into `name_part001.docx`, ... (default 200 pages),
  parallel OCR, a Stop button, and a failed page becomes an image instead of stopping the run.

## Limits
- Maths symbols inside Bengali/Hindi OCR lines (`=`, `+`, look-alike digits) can be misread: check
  against the page image shown under each scanned page.
- Old Bengali/Hindi PDFs using legacy fonts (SutonnyMJ, Kruti Dev) copy out as gibberish: tick
  "Force OCR".
- Digital two-column pages can interleave their columns.

## Building the Windows app yourself
Push a tag (`git tag v1.0 && git push --tags`) or run **Actions -> Build Windows app**. The workflow
installs Tesseract, downloads the OCR data and packages everything with PyInstaller.

## Licence note
PyMuPDF is AGPL-3.0 (or commercial). If you publish the app or its binaries, publish this
repository under a compatible licence (AGPL-3.0 is the simple choice) or buy a PyMuPDF commercial licence.
