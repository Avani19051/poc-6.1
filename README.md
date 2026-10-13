# PoC – Datasheet Extraction & McMaster-Carr Lookup

**Language:** Python 3.9+
**Libraries:** pdfplumber, pandas, openpyxl, pytesseract (Case 1) · Playwright, openpyxl (Case 2)

```
poc/
├── case1_extract.py        Case 1 – PDF datasheet → Excel + JSON
├── case2_mcmaster.py       Case 2 – MPN → McMaster-Carr → TableA (main program)
├── tests_mcmaster_lookup.py  Unit tests for Case 2 (run with unittest)
├── TableA.xlsx             Input table for Case 2 (auto-created if missing)
├── input/                  Datasheet PDFs for Case 1
├── sample_output/          Results produced by the programs
└── archive/                Superseded drafts (e.g. the original case2 program)
```

## Setup

```bash
pip install pdfplumber pandas openpyxl pytesseract playwright
playwright install chromium
```

Install the **Tesseract OCR** program too. It is used to read the dimension drawing:
- Windows: installer from https://github.com/UB-Mannheim/tesseract/wiki (add it to PATH)
- macOS: `brew install tesseract`
- Linux: `sudo apt install tesseract-ocr`

## Case 1 – Datasheet attribute extraction

```bash
python case1_extract.py                    # processes every PDF in ./input
python case1_extract.py some/datasheet.pdf # a specific file
python case1_extract.py --out ./results    # different output folder
```

Output goes to **Documents**: `Attributes_<MPN>_<Manufacturer>.xlsx`, with the sheets `Attribute Set-1` and `Attribute Set-2`, plus a `.json` copy.

**How it works (layered pipeline)**
1. **Text layer:** pdfplumber reads the text and regular expressions pick out each value. This is fast and exact.
2. **OCR layer (AI/ML):** the SemiGen package dimensions appear only inside a drawing (an image), which has no text layer. The program renders that page and runs Tesseract, an LSTM neural-network OCR engine, to confirm each dimension is present before writing it. The *Source* column says `OCR-verified` or `NOT verified by OCR`.
3. **Validation:** a value that can't be found is written as `Not Found`. The program never guesses.

The program detects which kind of datasheet it has (SemiGen or McMaster-Carr), so each format goes to the right parser. Supporting a new format means adding one parser function.

| SBC50-4-14-2 (SemiGen) | Value |
|---|---|
| Capacitance | 4.0 pF |
| Voltage | 50 VDC |
| Current | 1.0 µA (insulation/leakage current at 25 V) |
| Length | 0.035 ±0.005 in (0.889 ±0.127 mm) |
| Width | 0.012 in max (0.305 mm) |
| Height | 0.004 ±0.002 in (0.102 ±0.051 mm) |

| 91864A002 (McMaster-Carr) | Value |
|---|---|
| Head Type | Socket |
| Material | Black-Oxide Alloy Steel |
| RoHS Compliance | RoHS 3 (2015/863/EU) Compliant |
| Length | 5/16" |
| Head Diameter | 0.096" |

## Case 2 – McMaster-Carr lookup

```bash
python case2_mcmaster.py              # prompts for an MPN
python case2_mcmaster.py 91864A028    # MPN on the command line
python case2_mcmaster.py --all        # every MPN in TableA
python case2_mcmaster.py --headed     # watch the browser (helps if the site blocks headless)
python case2_mcmaster.py --login      # save a signed-in browser session once, then reuse it
python case2_mcmaster.py --signup --email you@example.com --password '...'   # create an account, then continue
```

Tests: `python -m unittest tests_mcmaster_lookup` (14 tests: MPN validation, parsing, Excel round-trip, rerun-resume).

Flow: validate the MPN against TableA → open mcmaster.com → search the MPN → click **Product Detail** if the page shows it → wait for the spec table → read **Material** and **RoHS Compliance** → update that row → save the complete table as `Documents/TableA_Updated.xlsx`.

Results (all four verified on 8 Oct 2026 from mcmaster.com product data; final workbook also saved to `~/Documents/TableA_Updated.xlsx`):

| MPN | Material | RoHS |
|---|---|---|
| 91864A028 | Black-Oxide Alloy Steel | RoHS 3 (2015/863/EU) Compliant |
| 90044A249 | Black-Oxide Alloy Steel | RoHS 3 (2015/863/EU) Compliant |
| 91251A107 | Black-Oxide Alloy Steel | RoHS 3 (2015/863/EU) Compliant |
| 91251A105 | Black-Oxide Alloy Steel | RoHS 3 (2015/863/EU) Compliant |

**Error handling**
- Input that isn't a valid MPN, or an MPN not in TableA → the user is asked again and shown the valid list.
- Page timeouts and the site's rate limiting → 3 retries with increasing back-off, plus a random 4–8 s gap between parts.
- McMaster's bot protection (Akamai WAF / "Please log in" wall / `ProdDatProtection`) → recognised and reported as `[BLOCKED]`; nothing is written for that row.
- A failed lookup writes `Lookup failed` **only** into cells that still hold a placeholder (`?`/blank), so a rerun never erases a previously verified value.
- The part isn't on the page, or the spec table is missing → the row is marked `Lookup failed` and the remaining parts still run.
- The workbook is saved after **every** MPN, so a crash loses nothing. Each rerun continues from the previous output.
- Playwright or Chromium not installed → the program prints the install command.

## Assumptions
- "Documents" means the user's `~/Documents` folder (OneDrive Documents on Windows if it exists).
- **Current** for the capacitor is the insulation-resistance (leakage) current, IR = 1.0 µA at 25 V. It is the only current value in the datasheet.
- Capacitor dimensions use package **PK14-2**, taken from the part's row in the "Available Parts" table. Length is the long axis of the beam-lead outline, width is the 0.012" max lead width, and height is the 0.004" chip thickness.
- For Datasheet 2 (a McMaster product page saved as PDF), the manufacturer/supplier is McMaster-Carr. "Head Type" maps to McMaster's "Fastener Head Type".
- McMaster often opens the detail view directly for an exact part number. In that case the "Product Detail" click is skipped and the log says so.

## Notes and possible improvements
- McMaster rate-limits and bot-blocks fast automated requests: pages stopped loading after a few quick lookups or from datacenter/AI-crawler IPs, and product data endpoints return 403 Akamai WAF errors. This is why the script has back-off, delays, login support, and recognises the block instead of failing. From a normal home/office network, a fresh `python case2_mcmaster.py --all` fetches live data; McMaster also offers an official Product Information API (approval required), which would be the production choice.
- An LLM or vision-model fallback (for example Gemini, or a local model through Ollama) could extract attributes from datasheets in unseen layouts. Its output should be checked against the source text, just as the OCR layer is.
