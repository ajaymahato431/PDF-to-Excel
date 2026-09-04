from __future__ import annotations

import argparse
import logging
import math
import os
import shutil
import sys
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import numpy as np
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

try:
    import cv2  # type: ignore
except ImportError as exc:  # pragma: no cover - cv2 should be available via requirements
    raise SystemExit(
        "OpenCV (cv2) is required but not installed. "
        "Install it with `pip install opencv-python` before running this script."
    ) from exc


LOGGER = logging.getLogger("pdf_to_excel")


@dataclass
class WordBBox:
    text: str
    left: float
    top: float
    right: float
    bottom: float

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2


def cluster_positions(values: Iterable[float], tolerance: float) -> List[float]:
    """
    Group nearby coordinate values into averaged clusters.
    """
    sorted_values = sorted(float(v) for v in values)
    if not sorted_values:
        return []

    clusters: List[List[float]] = [[sorted_values[0]]]
    for value in sorted_values[1:]:
        if value - clusters[-1][-1] <= tolerance:
            clusters[-1].append(value)
        else:
            clusters.append([value])
    return [sum(cluster) / len(cluster) for cluster in clusters]


def table_rows_from_edges(
    page: pdfplumber.page.Page,
    cluster_tolerance: float,
    min_vertical_height: float,
    min_horizontal_width: float,
) -> List[List[str]]:
    """
    Reconstruct table rows by intersecting vector edges (lines) with character positions.
    """
    vertical_edges = [
        edge
        for edge in page.edges
        if edge.get("orientation") == "v" and edge.get("height", 0) >= min_vertical_height
    ]
    horizontal_edges = [
        edge
        for edge in page.edges
        if edge.get("orientation") == "h" and edge.get("width", 0) >= min_horizontal_width
    ]

    if len(vertical_edges) < 2 or len(horizontal_edges) < 2:
        return []

    x_boundaries = [-math.inf] + cluster_positions(
        (edge["x0"] for edge in vertical_edges),
        tolerance=cluster_tolerance,
    ) + [math.inf]

    y_boundaries = [-math.inf] + cluster_positions(
        (edge["bottom"] for edge in horizontal_edges),
        tolerance=cluster_tolerance,
    ) + [math.inf]

    grid: List[List[List[Tuple[float, str]]]] = [
        [[] for _ in range(len(x_boundaries) - 1)] for _ in range(len(y_boundaries) - 1)
    ]

    for char in page.chars:
        x_center = (char["x0"] + char["x1"]) / 2
        y_center = (char["top"] + char["bottom"]) / 2

        row_idx = next(
            (i for i in range(len(y_boundaries) - 1) if y_boundaries[i] <= y_center < y_boundaries[i + 1]),
            None,
        )
        col_idx = next(
            (j for j in range(len(x_boundaries) - 1) if x_boundaries[j] <= x_center < x_boundaries[j + 1]),
            None,
        )
        if row_idx is None or col_idx is None:
            continue
        grid[row_idx][col_idx].append((char["x0"], char["text"]))

    reconstructed_rows: List[List[str]] = []
    for row_cells in grid:
        row_text: List[str] = []
        for cell in row_cells:
            if not cell:
                row_text.append("")
                continue
            cell.sort(key=lambda item: item[0])
            merged = "".join(piece for _, piece in cell)
            row_text.append(" ".join(merged.split()))
        if any(cell for cell in row_text):
            reconstructed_rows.append(row_text)

    return reconstructed_rows


def normalize_rows(
    rows: Sequence[Sequence[str]],
    min_column_fill: int,
    drop_leading_sparse: bool = True,
) -> List[List[str]]:
    """
    Filter empty columns, trim whitespace, and optionally drop a sparse title row.
    """
    if not rows:
        return []

    max_cols = max(len(row) for row in rows)
    column_counts = [
        sum(1 for row in rows if idx < len(row) and row[idx].strip())
        for idx in range(max_cols)
    ]

    keep_indices = [idx for idx, count in enumerate(column_counts) if count >= min_column_fill]
    if not keep_indices:
        return []

    filtered_rows: List[List[str]] = []
    for row in rows:
        filtered_rows.append(
            [row[idx].strip() if idx < len(row) else "" for idx in keep_indices]
        )

    filtered_rows = [row for row in filtered_rows if any(row)]
    if drop_leading_sparse and len(filtered_rows) >= 2:
        first_non_empty = sum(bool(cell) for cell in filtered_rows[0])
        second_non_empty = sum(bool(cell) for cell in filtered_rows[1])
        if first_non_empty < second_non_empty:
            filtered_rows = filtered_rows[1:]

    return filtered_rows


def _env_bool(key: str, default: bool = False) -> bool:
    val = os.getenv(key)
    if val is None:
        return default
    return val.strip().lower() in ("1", "true", "yes", "on")


def _env_float(key: str, default: float) -> float:
    val = os.getenv(key)
    if val is None:
        return default
    try:
        return float(val)
    except ValueError:
        return default


def _env_int(key: str, default: int) -> int:
    val = os.getenv(key)
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        return default


ENV_POPPLER_PATH = os.getenv("POPPLER_PATH")
ENV_TESSERACT_CMD = os.getenv("TESSERACT_CMD")
ENV_TESSDATA_DIR = os.getenv("TESSDATA_PREFIX") or os.getenv("TESSDATA_DIR")
ENV_OCR_LANG = os.getenv("OCR_LANG", "nep")
ENV_OCR_PSM = os.getenv("OCR_PSM", "6")
ENV_OCR_DPI = _env_int("OCR_DPI", 300)
ENV_OCR_CONFIDENCE = _env_int("OCR_CONFIDENCE", 70)
ENV_EXTRACTION_MODE = os.getenv("EXTRACTION_MODE", "prompt")
ENV_DEFAULT_OUTPUT = os.getenv("DEFAULT_OUTPUT", "extracted_table.xlsx")
ENV_MAX_COLUMNS = _env_int("MAX_COLUMNS", 12)
ENV_EDGE_CLUSTER_TOLERANCE = _env_float("EDGE_CLUSTER_TOLERANCE", 2.5)
ENV_EDGE_MIN_VERTICAL = _env_float("EDGE_MIN_VERTICAL", 18.0)
ENV_EDGE_MIN_HORIZONTAL = _env_float("EDGE_MIN_HORIZONTAL", 30.0)
ENV_EDGE_COLUMN_MIN_FILL = _env_int("EDGE_COLUMN_MIN_FILL", 1)
ENV_Y_TOLERANCE = _env_float("Y_TOLERANCE", 4.0)
ENV_MIN_GAP = _env_float("MIN_GAP", 18.0)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert tabular Nepali PDFs into Excel with pdfplumber + OCR fallback."
    )
    parser.add_argument("pdf_path", type=Path, help="Path to the input PDF file.")
    parser.add_argument(
        "--mode",
        choices=("prompt", "auto", "pdfplumber", "ocr"),
        default=ENV_EXTRACTION_MODE,
        help=(
            f"Extraction mode to use (default: {ENV_EXTRACTION_MODE}). 'prompt' asks at runtime, "
            "'auto' falls back to OCR only when needed, 'pdfplumber' disables OCR, and 'ocr' runs Tesseract directly."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(ENV_DEFAULT_OUTPUT),
        help=f"Output Excel path (default: {ENV_DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--pages",
        type=str,
        default=None,
        help="Pages to process (e.g. '1,2,5' or '1-3'). Default: all pages.",
    )
    parser.add_argument(
        "--column-breaks",
        type=float,
        nargs="+",
        default=None,
        help=(
            "Optional x-coordinate separators (PDF space) between columns. "
            "Example: --column-breaks 120 260 430"
        ),
    )
    parser.add_argument(
        "--y-tolerance",
        type=float,
        default=ENV_Y_TOLERANCE,
        help=f"Tolerance in points for grouping words into the same row (default: {ENV_Y_TOLERANCE}).",
    )
    parser.add_argument(
        "--min-gap",
        type=float,
        default=ENV_MIN_GAP,
        help=f"Minimum x-gap to consider when inferring column breaks automatically (default: {ENV_MIN_GAP}).",
    )
    parser.add_argument(
        "--poppler-path",
        type=Path,
        default=Path(ENV_POPPLER_PATH) if ENV_POPPLER_PATH else None,
        help=(
            "Directory containing Poppler binaries for pdf2image (default: auto-detected or env POPPLER_PATH)."
        ),
    )
    parser.add_argument(
        "--tesseract-cmd",
        type=Path,
        default=Path(ENV_TESSERACT_CMD) if ENV_TESSERACT_CMD else None,
        help="Override tesseract binary location (default: auto-detected or env TESSERACT_CMD).",
    )
    parser.add_argument(
        "--tessdata-dir",
        type=Path,
        default=Path(ENV_TESSDATA_DIR) if ENV_TESSDATA_DIR else None,
        help=(
            "Directory containing Tesseract language data (default: auto-detected or env TESSDATA_PREFIX)."
        ),
    )
    parser.add_argument(
        "--ocr-lang",
        type=str,
        default=ENV_OCR_LANG,
        help=f"Language code for Tesseract OCR (default: {ENV_OCR_LANG}).",
    )
    parser.add_argument(
        "--ocr-psm",
        type=str,
        default=ENV_OCR_PSM,
        help=f"Tesseract page segmentation mode (PSM). Default: {ENV_OCR_PSM}.",
    )
    parser.add_argument(
        "--ocr-confidence",
        type=int,
        default=ENV_OCR_CONFIDENCE,
        help=f"Minimum OCR confidence (0-100) for words to be kept (default: {ENV_OCR_CONFIDENCE}).",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=ENV_OCR_DPI,
        help=f"Rendering DPI for OCR fallback (default: {ENV_OCR_DPI}).",
    )
    parser.add_argument(
        "--force-ocr",
        action="store_true",
        help="Skip pdfplumber extraction and go straight to OCR.",
    )
    parser.add_argument(
        "--max-columns",
        type=int,
        default=ENV_MAX_COLUMNS,
        help=f"Maximum allowed columns before a detected table is discarded as noise (default: {ENV_MAX_COLUMNS}).",
    )
    parser.add_argument(
        "--edge-cluster-tolerance",
        type=float,
        default=ENV_EDGE_CLUSTER_TOLERANCE,
        help=f"Tolerance (points) when clustering table edge positions (default: {ENV_EDGE_CLUSTER_TOLERANCE}).",
    )
    parser.add_argument(
        "--edge-min-vertical",
        type=float,
        default=ENV_EDGE_MIN_VERTICAL,
        help=f"Minimum vertical edge height to consider for column detection (default: {ENV_EDGE_MIN_VERTICAL}).",
    )
    parser.add_argument(
        "--edge-min-horizontal",
        type=float,
        default=ENV_EDGE_MIN_HORIZONTAL,
        help=f"Minimum horizontal edge width to consider for row detection (default: {ENV_EDGE_MIN_HORIZONTAL}).",
    )
    parser.add_argument(
        "--edge-column-min-fill",
        type=int,
        default=ENV_EDGE_COLUMN_MIN_FILL,
        help=f"Discard edge-derived columns that have <= this many non-empty cells (default: {ENV_EDGE_COLUMN_MIN_FILL}).",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug logging.",
    )
    return parser.parse_args(argv)


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s | %(message)s",
    )


def configure_poppler(poppler_path: Optional[Path] = None) -> Optional[Path]:
    if poppler_path and poppler_path.exists():
        LOGGER.debug("Using Poppler binaries at %s", poppler_path)
        return poppler_path
    elif poppler_path:
        LOGGER.warning("Specified Poppler path %s does not exist.", poppler_path)

    if shutil.which("pdftoppm"):
        return None

    if sys.platform.startswith("win"):
        candidates = [
            Path(r"C:\Program Files\poppler\Library\bin"),
            Path(r"C:\Program Files\poppler\bin"),
            Path(r"C:\poppler\Library\bin"),
            Path(r"C:\poppler\bin"),
            Path(r"C:\poppler-25.07.0\Library\bin"),
        ]
        for candidate in candidates:
            if candidate.exists():
                LOGGER.debug("Auto-detected Poppler binaries at %s", candidate)
                return candidate

    return None


def configure_tesseract(binary: Optional[Path] = None, tessdata_dir: Optional[Path] = None) -> Optional[Path]:
    resolved_binary: Optional[Path] = None
    if binary and binary.exists():
        resolved_binary = binary
    elif binary:
        LOGGER.warning("Specified tesseract binary not found at %s", binary)

    if not resolved_binary:
        which_path = shutil.which("tesseract")
        if which_path:
            resolved_binary = Path(which_path)

    if not resolved_binary and sys.platform.startswith("win"):
        candidates = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
        ]
        for candidate in candidates:
            if candidate.exists():
                resolved_binary = candidate
                break

    if not resolved_binary:
        LOGGER.warning(
            "Tesseract binary not found on PATH or standard locations. "
            "OCR will fail unless Tesseract is installed or --tesseract-cmd is specified."
        )
        return None

    pytesseract.pytesseract.tesseract_cmd = str(resolved_binary)
    LOGGER.debug("Using Tesseract binary at %s", resolved_binary)

    candidate: Optional[Path] = None
    if tessdata_dir:
        candidate = tessdata_dir
    elif os.environ.get("TESSDATA_PREFIX"):
        candidate = Path(os.environ["TESSDATA_PREFIX"])
    elif resolved_binary:
        default_tessdata = resolved_binary.parent / "tessdata"
        if default_tessdata.exists():
            candidate = default_tessdata

    if candidate:
        if candidate.is_file():
            candidate = candidate.parent
        if candidate.is_dir() and candidate.name != "tessdata" and (candidate / "tessdata").is_dir():
            candidate = candidate / "tessdata"
        if candidate.exists():
            resolved_candidate = candidate.resolve()
            os.environ["TESSDATA_PREFIX"] = str(resolved_candidate)
            LOGGER.debug("Configured TESSDATA_PREFIX=%s", os.environ["TESSDATA_PREFIX"])
            return resolved_candidate
        else:
            LOGGER.warning("Candidate tessdata directory %s does not exist.", candidate)

    return None


def parse_pages(pages_argument: Optional[str], total_pages: int) -> List[int]:
    if not pages_argument:
        return list(range(total_pages))

    page_ids: List[int] = []
    for chunk in pages_argument.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start_str, end_str = chunk.split("-", 1)
            start = int(start_str) - 1
            end = int(end_str) - 1
            page_ids.extend(range(start, end + 1))
        else:
            page_ids.append(int(chunk) - 1)

    unique_sorted = sorted({idx for idx in page_ids if 0 <= idx < total_pages})
    if not unique_sorted:
        raise ValueError("No valid pages computed from --pages argument.")
    return unique_sorted


def resolve_extraction_mode(mode: str) -> str:
    if mode != "prompt":
        return mode
    if not sys.stdin.isatty():
        LOGGER.warning("Prompt mode requested but no interactive stdin is available; defaulting to auto.")
        return "auto"

    prompt = (
        "\nChoose an extraction method:\n"
        "  [1] Direct extraction with pdfplumber\n"
        "  [2] OCR with Tesseract (recommended for Nepali text)\n"
        "  [3] Auto (try pdfplumber first, fallback to OCR)\n"
        "Enter your choice [1-3, default 3]: "
    )

    while True:
        try:
            choice = input(prompt).strip()
        except EOFError:
            LOGGER.warning("No input detected; defaulting to auto mode.")
            return "auto"
        if choice in ("", "3"):
            return "auto"
        if choice == "1":
            return "pdfplumber"
        if choice == "2":
            return "ocr"
        print("Invalid choice. Please enter 1, 2, or 3.")


def extract_tables_with_pdfplumber(
    pdf_path: Path,
    pages: List[int],
    max_columns: int,
) -> Tuple[List[pd.DataFrame], List[Tuple[int, int]]]:
    dataframes: List[pd.DataFrame] = []
    provenance: List[Tuple[int, int]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx in pages:
            page = pdf.pages[page_idx]
            LOGGER.debug("Attempting pdfplumber.extract_tables on page %s", page_idx + 1)
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                header = [col.strip() if col else f"Column_{idx}" for idx, col in enumerate(table[0])]
                rows = [row for row in table[1:] if any(cell and cell.strip() for cell in row)]
                if not rows:
                    continue
                df = pd.DataFrame(rows, columns=header)
                if len(df.columns) <= max_columns:
                    dataframes.append(df)
                    provenance.append((page_idx + 1, table_idx + 1))
                else:
                    LOGGER.debug(
                        "Discarding table on page %s (table %s) due to %s columns exceeding limit %s",
                        page_idx + 1,
                        table_idx + 1,
                        len(df.columns),
                        max_columns,
                    )
    return dataframes, provenance


def extract_words_from_page(page: pdfplumber.page.Page) -> List[WordBBox]:
    raw_words = page.extract_words(
        x_tolerance=2,
        y_tolerance=2,
        keep_blank_chars=False,
        use_text_flow=True,
    )
    words: List[WordBBox] = []
    for raw in raw_words:
        text = raw.get("text", "").strip()
        if not text:
            continue
        words.append(
            WordBBox(
                text=text,
                left=float(raw.get("x0")),
                top=float(raw.get("top")),
                right=float(raw.get("x1")),
                bottom=float(raw.get("bottom")),
            )
        )
    return words


def group_words_into_rows(words: Iterable[WordBBox], y_tolerance: float) -> List[List[WordBBox]]:
    rows: List[List[WordBBox]] = []
    current_row: List[WordBBox] = []
    current_center = math.inf

    for word in sorted(words, key=lambda w: w.top):
        if current_row and abs(word.center_y - current_center) <= y_tolerance:
            current_row.append(word)
            current_center = (current_center * (len(current_row) - 1) + word.center_y) / len(current_row)
        else:
            if current_row:
                rows.append(current_row)
            current_row = [word]
            current_center = word.center_y
    if current_row:
        rows.append(current_row)
    return rows


def infer_column_edges(rows: Sequence[Sequence[WordBBox]], min_gap: float) -> List[float]:
    left_positions = sorted({round(word.left, 2) for row in rows for word in row})
    if not left_positions:
        return []

    tentative_edges: List[float] = [left_positions[0] - min_gap]
    for idx in range(len(left_positions) - 1):
        gap = left_positions[idx + 1] - left_positions[idx]
        if gap >= min_gap:
            tentative_edges.append(left_positions[idx] + gap / 2)
    tentative_edges.append(left_positions[-1] + min_gap)
    return sorted(tentative_edges)


def build_column_edges(
    column_breaks: Optional[Sequence[float]],
    rows: Sequence[Sequence[WordBBox]],
    min_gap: float,
) -> List[float]:
    if column_breaks:
        edges = [-math.inf] + sorted(column_breaks) + [math.inf]
        LOGGER.debug("Using user-specified column breaks: %s", column_breaks)
        return edges

    inferred_breaks = infer_column_edges(rows, min_gap=min_gap)
    if inferred_breaks:
        edges = [-math.inf] + inferred_breaks + [math.inf]
        LOGGER.debug("Inferred column edges: %s", inferred_breaks)
        return edges

    return [-math.inf, math.inf]


def rows_to_cells(rows: Sequence[Sequence[WordBBox]], column_edges: Sequence[float]) -> List[List[str]]:
    from bisect import bisect_right

    num_columns = len(column_edges) - 1
    table_rows: List[List[str]] = []

    for row in rows:
        sorted_row = sorted(row, key=lambda w: w.left)
        cells = [""] * num_columns
        for word in sorted_row:
            col_idx = bisect_right(column_edges, word.left) - 1
            col_idx = min(max(col_idx, 0), num_columns - 1)
            if cells[col_idx]:
                cells[col_idx] = f"{cells[col_idx]} {word.text}"
            else:
                cells[col_idx] = word.text
        if any(cell for cell in cells):
            table_rows.append([cell.strip() for cell in cells])
    return table_rows


def clean_structured_rows(rows: List[List[str]]) -> List[List[str]]:
    if not rows:
        return rows

    header = rows[0]
    cleaned: List[List[str]] = [header]
    for row in rows[1:]:
        if all(not cell for cell in row):
            continue
        if row == header:
            continue
        cleaned.append(row)
    return cleaned


def dataframes_from_words(
    pdf_path: Path,
    pages: List[int],
    column_breaks: Optional[Sequence[float]],
    y_tolerance: float,
    min_gap: float,
) -> List[pd.DataFrame]:
    dataframes: List[pd.DataFrame] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx in pages:
            page = pdf.pages[page_idx]
            words = extract_words_from_page(page)
            if not words:
                continue
            rows = group_words_into_rows(words, y_tolerance=y_tolerance)
            if not rows:
                continue
            column_edges = build_column_edges(column_breaks, rows, min_gap=min_gap)
            structured_rows = clean_structured_rows(rows_to_cells(rows, column_edges))
            if len(structured_rows) < 2:
                continue
            header, *data_rows = structured_rows
            df = pd.DataFrame(data_rows, columns=header)
            dataframes.append(df)

    return dataframes


def dataframes_from_edges(
    pdf_path: Path,
    pages: List[int],
    cluster_tolerance: float,
    min_vertical_height: float,
    min_horizontal_width: float,
    min_column_fill: int,
) -> List[pd.DataFrame]:
    rows_accumulated: List[List[str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx in pages:
            page = pdf.pages[page_idx]
            rows = table_rows_from_edges(
                page,
                cluster_tolerance=cluster_tolerance,
                min_vertical_height=min_vertical_height,
                min_horizontal_width=min_horizontal_width,
            )
            if rows:
                rows_accumulated.extend(rows)

    normalized_rows = normalize_rows(rows_accumulated, min_column_fill=min_column_fill)
    if len(normalized_rows) < 2:
        return []

    header, *data_rows = normalized_rows
    data_rows = [row for row in data_rows if row != header]
    if not data_rows:
        return []

    dataframe = pd.DataFrame(data_rows, columns=header)
    return [dataframe]


def preprocess_image_for_ocr(image: Image.Image) -> np.ndarray:
    cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return thresh


def words_from_ocr(
    image: np.ndarray,
    lang: str,
    psm: str,
    min_confidence: int,
    tessdata_dir: Optional[Path] = None,
) -> List[WordBBox]:
    config_parts = [f"--psm {psm}"]
    if tessdata_dir:
        resolved_dir = tessdata_dir.resolve()
        config_parts.append(f'--tessdata-dir "{resolved_dir.as_posix()}"')
    config = " ".join(config_parts)
    ocr_result = pytesseract.image_to_data(
        image,
        lang=lang,
        output_type=pytesseract.Output.DICT,
        config=config,
    )
    words: List[WordBBox] = []
    for idx, text in enumerate(ocr_result["text"]):
        if not text or not text.strip():
            continue
        try:
            conf = int(ocr_result["conf"][idx])
        except ValueError:
            conf = -1
        if conf < min_confidence:
            continue
        left = float(ocr_result["left"][idx])
        top = float(ocr_result["top"][idx])
        width = float(ocr_result["width"][idx])
        height = float(ocr_result["height"][idx])
        words.append(
            WordBBox(
                text=text.strip(),
                left=left,
                top=top,
                right=left + width,
                bottom=top + height,
            )
        )
    return words


def dataframes_from_ocr(
    pdf_path: Path,
    pages: List[int],
    column_breaks: Optional[Sequence[float]],
    y_tolerance: float,
    min_gap: float,
    dpi: int,
    lang: str,
    psm: str,
    min_confidence: int,
    poppler_path: Optional[Path] = None,
    tessdata_dir: Optional[Path] = None,
) -> List[pd.DataFrame]:
    dataframes: List[pd.DataFrame] = []
    for page_idx in pages:
        poppler_kwargs = {"poppler_path": str(poppler_path)} if poppler_path else {}
        images = convert_from_path(
            pdf_path,
            first_page=page_idx + 1,
            last_page=page_idx + 1,
            dpi=dpi,
            **poppler_kwargs,
        )
        if not images:
            continue
        processed = preprocess_image_for_ocr(images[0])
        tessdata_config = None if os.environ.get("TESSDATA_PREFIX") else tessdata_dir
        words = words_from_ocr(
            processed,
            lang=lang,
            psm=psm,
            min_confidence=min_confidence,
            tessdata_dir=tessdata_config,
        )
        if not words:
            continue
        rows = group_words_into_rows(words, y_tolerance=y_tolerance)
        if not rows:
            continue
        column_edges = build_column_edges(column_breaks, rows, min_gap=min_gap)
        structured_rows = clean_structured_rows(rows_to_cells(rows, column_edges))
        if len(structured_rows) < 2:
            continue
        header, *data_rows = structured_rows
        df = pd.DataFrame(data_rows, columns=header)
        dataframes.append(df)
    return dataframes


def group_tables_by_header(dfs: Sequence[pd.DataFrame]) -> "OrderedDict[str, pd.DataFrame]":
    if not dfs:
        raise ValueError("No dataframes provided.")

    grouped: "dict[Tuple[str, ...], List[pd.DataFrame]]" = {}
    for df in dfs:
        normalized_columns = tuple(col.strip() for col in df.columns)
        grouped.setdefault(normalized_columns, []).append(df)

    ordered_tables: "OrderedDict[str, pd.DataFrame]" = OrderedDict()
    for idx, (columns, frames) in enumerate(grouped.items(), start=1):
        merged = pd.concat(frames, ignore_index=True)
        for col_idx in range(len(merged.columns)):
            series = merged.iloc[:, col_idx]
            merged.iloc[:, col_idx] = series.astype(str).str.strip()

        sheet_name = f"Table{idx:02d}"
        ordered_tables[sheet_name] = merged

    return ordered_tables


def export_to_excel(tables: "OrderedDict[str, pd.DataFrame]", output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, dataframe in tables.items():
            dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    total_rows = sum(len(df) for df in tables.values())
    LOGGER.info(
        "Exported %s sheet(s) / %s row(s) to %s",
        len(tables),
        total_rows,
        output_path,
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    configure_logging(args.verbose)

    if not args.pdf_path.exists():
        LOGGER.error("PDF not found at %s", args.pdf_path)
        return 1

    mode = "ocr" if args.force_ocr else resolve_extraction_mode(args.mode)
    if args.force_ocr and mode != "ocr":
        LOGGER.info("Overriding prompted mode to 'ocr' because --force-ocr was supplied.")
        mode = "ocr"
    args.force_ocr = mode == "ocr"
    allow_ocr_fallback = mode != "pdfplumber"
    LOGGER.info("Extraction mode selected: %s", mode)

    resolved_tessdata: Optional[Path] = None
    poppler_path: Optional[Path] = None
    if allow_ocr_fallback:
        resolved_tessdata = configure_tesseract(args.tesseract_cmd, args.tessdata_dir)
        poppler_path = configure_poppler(args.poppler_path)

    with pdfplumber.open(args.pdf_path) as pdf:
        selected_pages = parse_pages(args.pages, len(pdf.pages))

    dataframes: List[pd.DataFrame] = []
    provenance: List[Tuple[int, int]] = []

    if not args.force_ocr:
        LOGGER.info("Step 1: pdfplumber table extraction")
        dfs, provenance = extract_tables_with_pdfplumber(
            args.pdf_path,
            selected_pages,
            max_columns=args.max_columns,
        )
        if dfs:
            dataframes.extend(dfs)
            LOGGER.info(
                "pdfplumber extracted %s table(s) from pages: %s",
                len(dfs),
                ", ".join(f"{page}" for page, _ in provenance),
            )

    if not dataframes and not args.force_ocr:
        LOGGER.info("Step 2: line-aware vector reconstruction")
        dfs = dataframes_from_edges(
            args.pdf_path,
            selected_pages,
            cluster_tolerance=args.edge_cluster_tolerance,
            min_vertical_height=args.edge_min_vertical,
            min_horizontal_width=args.edge_min_horizontal,
            min_column_fill=args.edge_column_min_fill,
        )
        if dfs:
            dataframes.extend(dfs)
            LOGGER.info(
                "Line-based reconstruction produced %s row(s).",
                sum(len(df) for df in dfs),
            )

    if not dataframes and not args.force_ocr:
        LOGGER.info("Step 3: pdfplumber word reconstruction")
        dfs = dataframes_from_words(
            args.pdf_path,
            selected_pages,
            column_breaks=args.column_breaks,
            y_tolerance=args.y_tolerance,
            min_gap=args.min_gap,
        )
        if dfs:
            dataframes.extend(dfs)
            LOGGER.info("Reconstructed %s table(s) from pdf words.", len(dfs))

    if not dataframes and allow_ocr_fallback:
        LOGGER.info("Step 4: OCR fallback via Tesseract (lang=%s)", args.ocr_lang)
        dfs = dataframes_from_ocr(
            args.pdf_path,
            selected_pages,
            column_breaks=args.column_breaks,
            y_tolerance=args.y_tolerance,
            min_gap=args.min_gap,
            dpi=args.dpi,
            lang=args.ocr_lang,
            psm=args.ocr_psm,
            min_confidence=args.ocr_confidence,
            poppler_path=poppler_path,
            tessdata_dir=resolved_tessdata,
        )
        if dfs:
            dataframes.extend(dfs)
            LOGGER.info("OCR recovered %s table(s).", len(dfs))
        else:
            LOGGER.error("OCR fallback failed to recover any tables.")
            return 2
    elif not dataframes:
        LOGGER.error("No tables were detected using pdfplumber-only mode.")
        return 3

    tables = group_tables_by_header(dataframes)
    export_to_excel(tables, args.output)

    return 0


if __name__ == "__main__":
    sys.exit(main())
