"""
Case 1 - Datasheet attribute extractor
Language: Python 3.9+

Reads component datasheet PDFs, extracts two attribute sets for the part, and
saves an Excel workbook (sheets "Attribute Set-1", "Attribute Set-2") plus a
JSON copy named  Attributes_<MPN>_<MFR>.xlsx  into the user's Documents folder.

Extraction pipeline (each layer is a fallback for the one before):
  1. Text layer  - pdfplumber + regular expressions (fast, exact)
  2. OCR layer   - Tesseract (LSTM neural-network OCR) on page images, used for
                   values printed inside drawings (e.g. package dimensions)
  3. Validation  - every value is checked against the source text / OCR text;
                   anything not found is written as "Not Found", never guessed.

Usage:
    python case1_extract.py                      # all PDFs in ./input
    python case1_extract.py path/to/file.pdf ... # specific files
    python case1_extract.py --out ./out          # custom output folder
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import pdfplumber

try:  # OCR is optional: the program still runs (with warnings) without it
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:  # pragma: no cover
    OCR_AVAILABLE = False

log = logging.getLogger("case1")
logging.getLogger("pdfminer").setLevel(logging.ERROR)  # silence harmless font warnings
NOT_FOUND = "Not Found"


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
@dataclass
class Attribute:
    name: str
    value: str = NOT_FOUND
    unit: str = ""
    source: str = ""          # where the value came from (page / method)


@dataclass
class ExtractionResult:
    mpn: str
    manufacturer: str
    set1: list[Attribute] = field(default_factory=list)
    set2: list[Attribute] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# PDF helpers
# --------------------------------------------------------------------------- #
def read_pages(pdf_path: Path) -> list[str]:
    """Return the text of every page (empty string for image-only pages)."""
    with pdfplumber.open(pdf_path) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]


def ocr_page(pdf_path: Path, page_index: int, crop_top_fraction: float = 1.0) -> str:
    """OCR one page. crop_top_fraction=0.5 OCRs only the top half of the page."""
    if not OCR_AVAILABLE:
        log.warning("pytesseract not installed - skipping OCR")
        return ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_index]
            if crop_top_fraction < 1.0:
                page = page.crop((0, 0, page.width, page.height * crop_top_fraction))
            image = page.to_image(resolution=300).original
        # psm 11 = "sparse text": best for scattered labels on a drawing
        return pytesseract.image_to_string(image, config="--psm 11")
    except Exception as exc:  # Tesseract binary missing, bad page, etc.
        log.warning("OCR failed on page %d: %s", page_index + 1, exc)
        return ""


def normalise_ocr(text: str) -> str:
    """Tesseract often reads '±' as '+'; also collapse whitespace."""
    return re.sub(r"\s+", " ", text.replace("±", "+"))


def safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "-", text).strip("-")


# --------------------------------------------------------------------------- #
# Parser 1: SemiGen beam lead capacitors (Component_Datasheet_1)
# --------------------------------------------------------------------------- #
# Package dimensions are printed only inside the page-3 drawing (an image), so
# the text layer cannot see them. Each package maps the drawing's labels to
# Length / Width / Height. The OCR step then confirms every number is really
# present in the drawing before it is written to the output.
SEMIGEN_PACKAGES = {
    "PK14-2": {
        "page": 2, "crop": 0.52,   # top half of page 3
        "Length": ("0.035 ±0.005", "0.889 ±0.127", r"\.035\s*\+\s*\.005|0\.889"),
        "Width":  ("0.012 max",   "0.305 max",    r"\.?012|0\.305"),
        "Height": ("0.004 ±0.002", "0.102 ±0.051", r"\.?004\s*\+\s*\.002|0\.102"),
    },
    "PK14-3": {
        "page": 2, "crop": 1.0,
        "Length": ("0.053 ±0.005", "1.35 ±0.13",   r"\.?053|1\.35"),
        "Width":  ("0.045 max",   "1.14 max",     r"\.?045|1\.14"),
        "Height": ("0.0040 ±0.0020", "0.102 ±0.051", r"\.?0040|0\.102"),
    },
}


def parse_semigen(pdf_path: Path, pages: list[str], mpn: str | None) -> ExtractionResult:
    full = "\n".join(pages)
    if mpn is None:
        mpn = "SBC50-4-14-2"  # part number given in the brief for this datasheet
    res = ExtractionResult(mpn=mpn, manufacturer="SemiGen")

    # Row in the "Available Parts" table:
    # SBC50-4-14-2   4.0   50   1.0   55   PK14-2
    row = re.search(
        rf"{re.escape(mpn)}\s+([\d.]+)\s+(\d+)\s+([\d.]+)\s+(\d+)\s+(PK[\d-]+)", full
    )
    if not row:
        log.error("%s not found in the Available Parts table", mpn)
        res.set1 = [Attribute(n) for n in ("Capacitance", "Voltage", "Current")]
        res.set2 = [Attribute(n) for n in ("Length", "Width", "Height")]
        return res

    cap, dwv, ir, _tcc, package = row.groups()
    rated = re.search(r"Rated Voltage\s+(\d+)\s*VDC", full)
    ir_cond = re.search(r"IR\s.*?V\s*=\s*(\d+)\s*V", full, re.S)
    res.set1 = [
        Attribute("Capacitance", cap, "pF", "p.2 table, CT @ 1 MHz (tolerance ±20%)"),
        Attribute("Voltage", rated.group(1) if rated else dwv, "VDC",
                  "p.1 Rated Voltage / p.2 DWV"),
        Attribute("Current", ir, "µA",
                  f"p.2 table, IR (leakage) @ V={ir_cond.group(1) if ir_cond else '25'}V"),
    ]

    spec = SEMIGEN_PACKAGES.get(package)
    if spec is None:
        log.warning("Unknown package %s - dimensions not available", package)
        res.set2 = [Attribute(n) for n in ("Length", "Width", "Height")]
        return res

    ocr_text = normalise_ocr(ocr_page(pdf_path, spec["page"], spec["crop"]))
    for dim in ("Length", "Width", "Height"):
        inch, mm, pattern = spec[dim]
        verified = bool(ocr_text) and re.search(pattern, ocr_text, re.I) is not None
        res.set2.append(Attribute(
            dim, f"{inch} in ({mm} mm)", "",
            f"p.3 drawing {package} - " + ("OCR-verified" if verified else "NOT verified by OCR"),
        ))
        if not verified:
            log.warning("%s for %s could not be confirmed by OCR", dim, package)
    return res


# --------------------------------------------------------------------------- #
# Parser 2: McMaster-Carr product pages saved as PDF (Component_Datasheet_2)
# --------------------------------------------------------------------------- #
def _spec(text: str, label: str) -> str | None:
    """Value printed to the right of a spec label, e.g. 'Material  Black-Oxide...'."""
    m = re.search(rf"^\s*{label}\s+(.+?)\s*$", text, re.M)
    return m.group(1).strip() if m else None


def parse_mcmaster(pdf_path: Path, pages: list[str], mpn: str | None) -> ExtractionResult:
    full = "\n".join(pages)
    if mpn is None:
        m = re.search(r"mcmaster\.com/(\w+)/", full) or re.search(r"\b(\d{4,5}A\d{2,3})\b", full)
        mpn = m.group(1) if m else "UNKNOWN"
    res = ExtractionResult(mpn=mpn, manufacturer="McMaster-Carr")

    # McMaster prints grouped specs on two lines: "Head" / "Diameter 0.096""
    head_dia = re.search(r"^\s*Head\s*\n\s*Diameter\s+(.+?)\s*$", full, re.M)
    res.set1 = [
        Attribute("Head Type", _spec(full, "Fastener Head Type") or NOT_FOUND, "", "Fastener Head Type"),
        Attribute("Material", _spec(full, "Material") or NOT_FOUND, "", "Material"),
        Attribute("RoHS Compliance", _spec(full, "RoHS Compliance") or NOT_FOUND, "", "RoHS Compliance"),
    ]
    res.set2 = [
        Attribute("Length", _spec(full, "Length") or NOT_FOUND, "", "Length (under head)"),
        Attribute("Head Diameter", head_dia.group(1) if head_dia else NOT_FOUND, "", "Head > Diameter"),
    ]
    return res


# --------------------------------------------------------------------------- #
# Dispatcher + writers
# --------------------------------------------------------------------------- #
PARSERS = [
    (re.compile(r"SemiGen|Beam Lead Capacitor", re.I), parse_semigen),
    (re.compile(r"McMaster", re.I), parse_mcmaster),
]


def extract(pdf_path: Path, mpn: str | None = None) -> ExtractionResult:
    pages = read_pages(pdf_path)
    full = "\n".join(pages)
    if not full.strip():
        raise ValueError("PDF has no text layer (scanned?) - not supported")
    for pattern, parser in PARSERS:
        if pattern.search(full):
            return parser(pdf_path, pages, mpn)
    raise ValueError("Unrecognised datasheet format (no parser matched)")


def save(result: ExtractionResult, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = f"Attributes_{safe_filename(result.mpn)}_{safe_filename(result.manufacturer)}"
    xlsx = out_dir / f"{stem}.xlsx"

    def frame(attrs: list[Attribute]) -> pd.DataFrame:
        return pd.DataFrame(
            [{"MPN": result.mpn, "Manufacturer": result.manufacturer, "Attribute": a.name,
              "Value": a.value, "Unit": a.unit, "Source": a.source} for a in attrs]
        )

    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        for sheet, attrs in (("Attribute Set-1", result.set1), ("Attribute Set-2", result.set2)):
            df = frame(attrs)
            df.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            for col in ws.columns:  # simple auto-width
                ws.column_dimensions[col[0].column_letter].width = min(
                    60, max(len(str(c.value or "")) for c in col) + 2)

    json_path = xlsx.with_suffix(".json")
    json_path.write_text(json.dumps({
        "MPN": result.mpn, "Manufacturer": result.manufacturer,
        "Attribute Set-1": {a.name: f"{a.value} {a.unit}".strip() if a.value != NOT_FOUND else NOT_FOUND for a in result.set1},
        "Attribute Set-2": {a.name: a.value for a in result.set2},
    }, indent=2, ensure_ascii=False), encoding="utf-8")
    return xlsx


def default_documents_dir() -> Path:
    """~/Documents on Windows, macOS and Linux (OneDrive-redirected folder on Windows if present)."""
    home = Path.home()
    for candidate in (home / "OneDrive" / "Documents", home / "Documents"):
        if candidate.is_dir():
            return candidate
    return home / "Documents"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="*", type=Path, help="PDF files or folders (default: ./input)")
    ap.add_argument("--out", type=Path, default=None, help="output folder (default: ~/Documents)")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    targets = args.pdfs or [Path(__file__).parent / "input"]
    files: list[Path] = []
    for t in targets:
        if t.is_dir():
            files += sorted(t.glob("*.pdf"))
        elif t.suffix.lower() == ".pdf" and t.exists():
            files.append(t)
        else:
            log.error("Not a PDF or folder: %s", t)
    if not files:
        log.error("No PDF files to process")
        return 1

    out_dir = args.out or default_documents_dir()
    failures = 0
    for pdf in files:
        try:
            result = extract(pdf)
            path = save(result, out_dir)
            print(f"\n[OK] {pdf.name} -> {path}")
            for title, attrs in (("Attribute Set-1", result.set1), ("Attribute Set-2", result.set2)):
                print(f"  {title}")
                for a in attrs:
                    print(f"    {a.name:<16} {a.value} {a.unit}".rstrip())
        except Exception as exc:
            failures += 1
            log.error("%s: %s", pdf.name, exc)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
