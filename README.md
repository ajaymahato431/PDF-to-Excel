# PDF to Excel (Nepali-friendly)

Convert table-centric PDFs into multi-sheet Excel workbooks with a flexible pipeline that supports both direct text extraction and OCR tuned for Nepali (`nep`) content.

## Key Features
- **Hybrid extraction workflow** – try `pdfplumber`, fall back to OCR with Tesseract, or force either mode explicitly.
- **Nepali OCR ready** – defaults to Tesseract’s `nep` language pack with Poppler-based rasterisation at configurable DPI.
- **Table reconstruction utilities** – multiple heuristics (vector edges, word clustering) reduce the manual tweaking often needed for semi-structured PDFs.
- **Excel output** – each detected table is written to its own sheet (`Table01`, `Table02`, …) with stripped, normalised cell values.

## Prerequisites
1. **Python** 3.11+ (matches the repository’s `venv`).
2. **Tesseract OCR** installed locally (Windows default path: `C:\Program Files\Tesseract-OCR`). Install the `nep.traineddata` language file in the `tessdata` directory.
3. **Poppler** binaries (for `pdf2image`). Suggested path: `C:\poppler-25.07.0\Library\bin`.
4. Optional but recommended: set an environment variable so Tesseract always finds the language data:

   ```text
   Variable name: TESSDATA_PREFIX
   Variable value: C:\Program Files\Tesseract-OCR\tessdata
   ```

   > The value must point directly to the `tessdata` folder, not the parent directory.

## Installation
```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

## Usage
Basic invocation:
```powershell
python pdf_to_excel.py "input.pdf"
```
You will be prompted to choose the extraction mode unless you pass `--mode` explicitly.

### Common Options
| Option | Description |
| --- | --- |
| `--mode {prompt,auto,pdfplumber,ocr}` | Runtime strategy. `prompt` (default) asks interactively. `auto` favours pdfplumber, falling back to OCR. `pdfplumber` disables OCR. `ocr` skips straight to Tesseract. |
| `-o`, `--output PATH` | Target Excel file (default `extracted_table.xlsx`). |
| `--pages "1,3-5"` | Limit processing to specific pages (1-based). |
| `--column-breaks <floats>` | Manually supply column separators for reconstruction (measured in PDF space). |
| `--poppler-path PATH` | Override Poppler binary directory if it isn’t on PATH. |
| `--tesseract-cmd PATH` | Explicit path to `tesseract.exe`. |
| `--tessdata-dir PATH` | Directory that contains the `tessdata` folder; used if `TESSDATA_PREFIX` isn’t already set. |
| `--ocr-lang CODE` | Tesseract language code (default `nep`). |
| `--dpi INT` | Rendering DPI for OCR fallback (default `300`). Higher values improve accuracy but increase runtime. |
| `--force-ocr` | Skip all pdfplumber-based steps and run OCR directly (equivalent to `--mode ocr`). |
| `--verbose` | Enable detailed logging for troubleshooting. |

Run `python pdf_to_excel.py --help` for the full list, including advanced table-reconstruction knobs (`--edge-*`, `--min-gap`, `--y-tolerance`, etc.).

### Example Workflows
1. **Direct OCR in Nepali (explicit binaries)**  
   ```powershell
   python pdf_to_excel.py ".\Bharatpur Municipality.pdf" --mode ocr ^
     --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe" ^
     --tessdata-dir "C:\Program Files\Tesseract-OCR\tessdata" ^
     --poppler-path "C:\poppler-25.07.0\Library\bin" ^
     -o ".\output\bharatpur.xlsx"
   ```

2. **Auto mode with manual column hints**  
   ```powershell
   python pdf_to_excel.py ".\input.pdf" --mode auto --column-breaks 120 360 710
   ```

3. **Restrict to pages 2–4 and increase OCR DPI**  
   ```powershell
   python pdf_to_excel.py ".\survey.pdf" --mode ocr --pages 2-4 --dpi 400
   ```

## Customising OCR Column Detection
OCR-based extraction benefits from accurate column boundaries. If results look jumbled:
1. Increase `--dpi` (e.g., 400).
2. Supply `--column-breaks` from the PDF coordinate space, or capture approximate x-positions from the rendered images (300 DPI by default).
3. Experiment with `--y-tolerance` and `--min-gap` to stabilise row grouping.

## Troubleshooting Tips
- **Tesseract can’t load `nep.traineddata`:** ensure the language file exists inside the folder pointed to by `TESSDATA_PREFIX` or pass `--tessdata-dir` directly.
- **No tables detected:** try `--mode ocr`, increase `--max-columns`, or provide manual `--column-breaks`.
- **Garbage OCR output:** clean the source (deskew/denoise) or adjust `--ocr-psm` (e.g., `--ocr-psm 4` for column text). Double-check the Poppler path sends high-quality images to Tesseract.

## License
This repository does not currently specify a license; treat it as proprietary unless the author states otherwise.

## Contributing
1. Create a branch for your changes.
2. Run lint/tests (if available) and `python -m compileall pdf_to_excel.py` to catch syntax errors.
3. Submit a pull request summarising the modifications.

---
Need help adapting the OCR boundaries for a new layout? Open an issue or start a discussion with example pages so the heuristics can be fine-tuned.
