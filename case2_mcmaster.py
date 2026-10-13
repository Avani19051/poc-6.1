#!/usr/bin/env python3
"""
McMaster-Carr Material / RoHS lookup  ->  TableA.xlsx
=====================================================

Language : Python 3.9+   (browser automation with Playwright + Chromium)

What it does
------------
1. Asks the user for an MPN.  The MPN must be one of the part numbers already
   listed in TableA.xlsx (a McMaster-style number such as 91864A028); otherwise
   the user is asked again and shown the valid list.
2. Opens https://www.mcmaster.com/products/<mpn>/, then clicks the
   "Product Detail" hyperlink (the red box in the brief) to open the final page.
3. Reads the "Material" and "RoHS Compliance" values from the product detail
   page's specification table.
4. Writes both values into that MPN's row in TableA and saves the complete table
   as  TableA_Updated.xlsx  in the user's Documents folder.

Usage
-----
    python mcmaster_lookup.py                 # interactive prompt for one MPN
    python mcmaster_lookup.py 91864A028       # one or more MPNs on the command line
    python mcmaster_lookup.py --all           # every MPN in TableA
    python mcmaster_lookup.py --login         # open a browser to sign in once, then exit
    python mcmaster_lookup.py --all --headed  # watch the browser while it works

    # Non-interactive sign-in with a real McMaster account (free to create):
    export MCMASTER_EMAIL="you@example.com"
    export MCMASTER_PASSWORD="your-password"
    python mcmaster_lookup.py --all

    # ...or pass them directly:
    python mcmaster_lookup.py --all --email you@example.com --password 'your-password'

Install once
------------
    pip install playwright openpyxl
    python -m playwright install chromium

Notes
-----
McMaster-Carr protects its product data against anonymous automation.  When the
program detects a login / data-protection wall it signs in (if credentials were
supplied) or explains what to do instead of silently returning junk.  It uses a
genuine account you created yourself; it does not attempt to bypass
authentication.  A successful sign-in is cached in a local browser profile, so
later runs usually stay signed in.
"""

from __future__ import annotations

import argparse
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover - dependency guard
    sys.exit("Missing dependency: openpyxl.  Install it with:  pip install openpyxl")

LOG = logging.getLogger("mcmaster")

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
BASE_URL = "https://www.mcmaster.com/"
# Direct product URL.  It shows a "Product Detail" hyperlink that leads to the
# final specification page (Material / RoHS Compliance).
PRODUCT_URL = "https://www.mcmaster.com/products/{mpn}/"
HERE = Path(__file__).resolve().parent
TABLE_A = HERE / "TableA.xlsx"
OUTPUT_NAME = "TableA_Updated.xlsx"
PROFILE_DIR = HERE / ".playwright-profile"

HEADERS = ["MPN", "MANUFACTURER NAME", "Material", "RoHS"]

# TableA exactly as given in the brief (used to create the workbook if missing).
DEFAULT_ROWS = [
    ["91864A028", "MCMASTER CARR SUPPLY", "?", "?"],
    ["90044A249", "MCMASTER CARR SUPPLY", "?", "?"],
    ["91251A107", "MCMASTER CARR SUPPLY", "?", "?"],
    ["91251A105", "MCMASTER CARR SUPPLY", "?", "?"],
]

# A McMaster part number looks like 91864A028 / 90044A249.
MPN_RE = re.compile(r"^\d{4,5}[A-Z]\d{1,4}$")

PAGE_TIMEOUT_MS = 45_000
RETRIES = 3
MIN_DELAY_BETWEEN_PARTS_S = 4.0
MAX_DELAY_BETWEEN_PARTS_S = 8.0

# Specification labels we care about (matched case-insensitively).
MATERIAL_LABELS = ("material", "base material", "body material", "fastener material",
                   "head material", "shaft material", "housing material")
ROHS_LABELS = ("rohs compliance", "rohs", "rohs status", "rohs 3", "rohs compliant")

# Credentials for a real McMaster account.  Set these environment variables
# (or pass --email/--password) and the program signs in before looking parts up.
# Use a genuine account you created at https://www.mcmaster.com/ ; the program
# does not attempt to bypass authentication.
ENV_EMAIL = "MCMASTER_EMAIL"
ENV_PASSWORD = "MCMASTER_PASSWORD"

# Strings that indicate the product-data / login protection kicked in.
BLOCK_MARKERS = (
    "to continue browsing, please log in",
    "please log in",
    "access denied",
    "unusual traffic",
    "verify you are human",
    "are you a robot",
)

NOT_FOUND = "Not Found"


# --------------------------------------------------------------------------- #
# Result container
# --------------------------------------------------------------------------- #
@dataclass
class Specs:
    """The two attributes we extract, plus the raw label/value pairs we saw."""

    material: str = NOT_FOUND
    rohs: str = NOT_FOUND
    raw: dict[str, str] = field(default_factory=dict)
    blocked: bool = False

    def is_complete(self) -> bool:
        return self.material != NOT_FOUND and self.rohs != NOT_FOUND


# --------------------------------------------------------------------------- #
# TableA (Excel) helpers
# --------------------------------------------------------------------------- #
def documents_dir() -> Path:
    """Return the user's Documents folder (OneDrive Documents on Windows if used)."""
    home = Path.home()
    for candidate in (home / "OneDrive" / "Documents", home / "Documents"):
        if candidate.is_dir():
            return candidate
    fallback = home / "Documents"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def ensure_table_a(path: Path = TABLE_A) -> Path:
    """Create TableA.xlsx from the brief's rows if it does not already exist."""
    if path.exists():
        return path
    wb = Workbook()
    ws = wb.active
    ws.title = "TableA"
    ws.append(HEADERS)
    for row in DEFAULT_ROWS:
        ws.append(row)
    style_sheet(ws)
    wb.save(path)
    LOG.info("Created %s", path)
    return path


def style_sheet(ws) -> None:
    """Bold header row + sensible column widths (safe to call repeatedly)."""
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        width = 14
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row_idx, col_idx).value
            width = max(width, len(str(value)) + 2 if value is not None else 0)
        ws.column_dimensions[letter].width = width


@dataclass
class Table:
    """A loaded workbook plus quick lookup helpers."""

    workbook: object
    sheet: object
    header: dict[str, int]           # UPPER-cased header name -> 1-based column
    rows_by_mpn: dict[str, int]      # UPPER-cased MPN -> 1-based row

    @property
    def mpns(self) -> list[str]:
        return list(self.rows_by_mpn)

    def set_cell(self, mpn: str, column: str, value: str) -> None:
        r = self.rows_by_mpn[mpn.upper()]
        c = self.header[column.upper()]
        self.sheet.cell(r, c, value)

    def get_cell(self, mpn: str, column: str):
        r = self.rows_by_mpn[mpn.upper()]
        c = self.header[column.upper()]
        return self.sheet.cell(r, c).value

    def set_cell_if_unverified(self, mpn: str, column: str, value: str) -> None:
        """Only overwrite a cell that still holds a placeholder (nothing real to lose)."""
        current = str(self.get_cell(mpn, column) or "").strip()
        if current in {"", "?", "Not Found", "Lookup failed"}:
            self.set_cell(mpn, column, value)

    def save(self, path: Path) -> None:
        try:
            self.workbook.save(path)
        except PermissionError as exc:  # file is open in Excel/LibreOffice
            raise RuntimeError(
                f"Cannot write {path} - is it open in Excel? Close it and retry."
            ) from exc


def load_table(input_path: Path, output_path: Path) -> Table:
    """
    Load TableA, preferring the previous run's output so earlier results survive
    a rerun (reruns continue instead of starting over).
    """
    source = output_path if output_path.exists() else input_path
    wb = load_workbook(source)
    ws = wb.active

    header = {}
    for cell in ws[1]:
        if cell.value is not None:
            header[str(cell.value).strip().upper()] = cell.column

    required = {h.upper() for h in HEADERS}
    missing = required - set(header)
    if missing:
        raise RuntimeError(
            f"{source} is missing column(s): {', '.join(sorted(missing))}. "
            f"Expected headers: {', '.join(HEADERS)}"
        )

    rows: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        mpn = ws.cell(r, header["MPN"]).value
        if mpn is None or not str(mpn).strip():
            continue
        key = normalise_mpn(str(mpn))
        if key in rows:
            LOG.warning("TableA contains duplicate MPN %s (row %d); using the first", key, r)
            continue
        rows[key] = r
    return Table(wb, ws, header, rows)


def normalise_mpn(value: str) -> str:
    """Upper-case and strip an MPN so 91864a028 and ' 91864A028 ' match."""
    return re.sub(r"\s+", "", value).upper()


# --------------------------------------------------------------------------- #
# Parsing --------------------------------------------------------------------
# --------------------------------------------------------------------------- #
def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _match_label(label: str, candidates: Iterable[str]) -> bool:
    """
    Match a spec label against known spellings.

    Trailing punctuation and a trailing qualifier such as "Material (Body)" are
    dropped first, then an exact match is required.  Exact matching is deliberate:
    a loose prefix match would wrongly treat the catalog category
    "Material Handling" as the spec label "Material".
    """
    label = _clean(label).lower().rstrip(":").strip()
    label = re.sub(r"\s*\(.*?\)\s*$", "", label).strip()
    return label in candidates


def parse_specs_from_rows(rows: list[list[str]]) -> dict[str, str]:
    """
    Parse spec label/value pairs from structured rows.

    `rows` is a list of [label, value, ...] lists (from table rows / dl pairs).
    Only the first value for a label is kept.
    """
    found: dict[str, str] = {}
    for cells in rows:
        if len(cells) < 2:
            continue
        label, value = _clean(cells[0]), _clean(cells[1])
        if not value:
            continue
        if "Material" not in found and _match_label(label, MATERIAL_LABELS):
            found["Material"] = value
        if "RoHS Compliance" not in found and _match_label(label, ROHS_LABELS):
            found["RoHS Compliance"] = value
    return found


def parse_specs_from_text(text: str) -> dict[str, str]:
    """
    Fallback parser for McMaster's rendered text, where the spec table appears as
    alternating label / value lines:

        Material
        Black-Oxide Alloy Steel
        RoHS Compliance
        RoHS 3 (2015/863/EU) Compliant
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    found: dict[str, str] = {}
    for i, line in enumerate(lines[:-1]):
        if "Material" not in found and _match_label(line, MATERIAL_LABELS):
            found["Material"] = lines[i + 1]
        elif "RoHS Compliance" not in found and _match_label(line, ROHS_LABELS):
            found["RoHS Compliance"] = lines[i + 1]
    return found


def parse_specs(text: str, rows: Optional[list[list[str]]] = None) -> Specs:
    """Combine the structured and text parsers, preferring structured data."""
    merged: dict[str, str] = {}
    if rows:
        merged.update(parse_specs_from_rows(rows))
    for key, value in parse_specs_from_text(text).items():
        merged.setdefault(key, value)

    blocked = any(marker in text.lower() for marker in BLOCK_MARKERS)
    material = merged.get("Material", NOT_FOUND) or NOT_FOUND
    rohs = merged.get("RoHS Compliance", NOT_FOUND) or NOT_FOUND
    return Specs(material=material, rohs=rohs, raw=merged, blocked=blocked)


# --------------------------------------------------------------------------- #
# Browser automation
# --------------------------------------------------------------------------- #
# Collect every spec-like label/value pair that exists in the DOM.  We try the
# common markup patterns (tables, definition lists, and McMaster's row divs).
_SPEC_ROWS_JS = r"""
() => {
  const out = [];
  const push = (a, b) => {
    if (a && b) out.push([a.trim(), b.trim()]);
  };
  document.querySelectorAll('table').forEach(t => {
    t.querySelectorAll('tr').forEach(tr => {
      const cells = [...tr.querySelectorAll('th,td')].map(c => c.innerText);
      if (cells.length >= 2) push(cells[0], cells[1]);
    });
  });
  document.querySelectorAll('dl').forEach(dl => {
    dl.querySelectorAll('dt').forEach(dt => {
      const dd = dt.nextElementSibling;
      if (dd && dd.tagName === 'DD') push(dt.innerText, dd.innerText);
    });
  });
  document.querySelectorAll('[class*="spec" i] [class*="row" i]').forEach(row => {
    const cells = [...row.children].map(c => c.innerText);
    if (cells.length >= 2) push(cells[0], cells[1]);
  });
  return out;
}
"""


class McMasterBlocked(RuntimeError):
    """Raised when the login / product-data-protection wall stops us."""


class Scraper:
    """Thin Playwright wrapper around the McMaster search + product detail flow."""

    def __init__(self, *, headless: bool = True, profile_dir: Path = PROFILE_DIR,
                 timeout_ms: int = PAGE_TIMEOUT_MS,
                 credentials: Optional[tuple[str, str]] = None):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:  # pragma: no cover - dependency guard
            raise RuntimeError(
                "Playwright is not installed.  Run:\n"
                "    pip install playwright\n"
                "    python -m playwright install chromium"
            ) from exc

        self.credentials = credentials          # (email, password) or None
        self._login_attempted = False
        self._pw = sync_playwright().start()
        self.profile_dir = Path(profile_dir)
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        self.context = self._pw.chromium.launch_persistent_context(
            user_data_dir=str(self.profile_dir),
            headless=headless,
            viewport={"width": 1440, "height": 900},
            locale="en-US",
            args=["--disable-blink-features=AutomationControlled"],
        )
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.set_default_timeout(timeout_ms)

    # -- lifecycle ---------------------------------------------------------- #
    def close(self) -> None:
        try:
            self.context.close()
        finally:
            self._pw.stop()

    # -- flow --------------------------------------------------------------- #
    def _quiet_wait(self) -> None:
        """Let the client-side app swap in content without hard-waiting on network idle."""
        try:
            self.page.wait_for_load_state("networkidle", timeout=15_000)
        except Exception:
            pass
        self.page.wait_for_timeout(1_500)

    def _open_product_page(self, mpn: str) -> None:
        """Go straight to /products/<mpn>/, which lists the 'Product Detail' link."""
        self.page.goto(PRODUCT_URL.format(mpn=mpn.lower()), wait_until="domcontentloaded")
        self._quiet_wait()

    def _search_box(self, mpn: str) -> None:
        """Fallback: use the masthead search box."""
        page = self.page
        page.goto(BASE_URL, wait_until="domcontentloaded")
        search = page.get_by_placeholder(re.compile("search", re.I)).first
        search.wait_for(state="visible")
        search.click()
        search.fill(mpn)
        search.press("Enter")
        self._quiet_wait()

    def _click_product_detail(self, mpn: str) -> bool:
        """Click the 'Product Detail' hyperlink (the red box in the brief) if shown."""
        page = self.page
        # Prefer a real anchor (a hyperlink) whose text or href says "Product Detail".
        detail = page.locator(
            "a:has-text('Product Detail'), a[href*='productdetail' i], "
            "[role='link']:has-text('Product Detail')"
        ).first
        try:
            detail.wait_for(state="visible", timeout=8_000)
            detail.scroll_into_view_if_needed()
            detail.click()
            LOG.debug("%s: clicked 'Product Detail'", mpn)
            self._quiet_wait()
            return True
        except Exception:
            # Some exact part numbers open the detail view without the extra click.
            LOG.debug("%s: no 'Product Detail' link - assuming the detail view is open", mpn)
            return False

    def _read_page(self) -> Specs:
        text = self.page.locator("body").inner_text()
        rows = self.page.evaluate(_SPEC_ROWS_JS)
        return parse_specs(text, rows)

    def lookup(self, mpn: str) -> Specs:
        """
        Open https://www.mcmaster.com/products/<mpn>/, click its 'Product Detail'
        hyperlink, and return Material + RoHS Compliance from the detail page.
        """
        last_error: Optional[Exception] = None
        for attempt in range(1, RETRIES + 1):
            try:
                # Primary route: the direct /products/<mpn>/ page.
                self._open_product_page(mpn)
                # If that produced nothing usable, fall back to the search box.
                if not self._read_page().raw:
                    LOG.debug("%s: /products/ page empty, trying the search box", mpn)
                    self._search_box(mpn)
                self._click_product_detail(mpn)
                specs = self._read_page()

                text = self.page.locator("body").inner_text()
                if specs.blocked and not specs.is_complete():
                    # If we have credentials we haven't tried yet, sign in and retry.
                    if self.credentials and not self._login_attempted:
                        self._login_attempted = True
                        LOG.info("Blocked - signing in with %s...", self.credentials[0])
                        self.login(*self.credentials)
                        continue
                    raise McMasterBlocked(
                        "McMaster blocked automated access (login required). "
                        "Provide MCMASTER_EMAIL / MCMASTER_PASSWORD or run --login once."
                    )
                if normalise_mpn(mpn) not in normalise_mpn(text):
                    raise RuntimeError("page did not show this part number")
                if not specs.is_complete():
                    # One more moment in case the spec table is still rendering.
                    self.page.wait_for_timeout(2_000)
                    specs = self._read_page()
                if not specs.is_complete() and not specs.raw:
                    raise RuntimeError("specification table not found on the page")
                return specs
            except McMasterBlocked:
                raise
            except Exception as exc:  # noqa: BLE001 - retry any transient failure
                last_error = exc
                wait = 4 * attempt + random.uniform(0, 3)
                LOG.warning("%s: attempt %d/%d failed (%s); retrying in %.0fs",
                            mpn, attempt, RETRIES, str(exc).splitlines()[0], wait)
                time.sleep(wait)
        raise RuntimeError(f"could not read {mpn} after {RETRIES} attempts: {last_error}")

    def is_logged_in(self) -> bool:
        """Heuristic: not showing the login wall and the login form is hidden."""
        if "mcmaster.com" not in (self.page.url or ""):
            return False
        body = self.page.locator("body").inner_text().lower()
        if "to continue browsing, please log in" in body:
            return False
        email = self.page.locator("#Email")
        return email.count() == 0 or not email.first.is_visible()

    def login(self, email: str, password: str) -> None:
        """
        Sign in with a real McMaster account (email + password).

        The form fields are #Email and #Password with a "Log in" submit button.
        This uses *your* credentials; it does not attempt to bypass auth.
        """
        page = self.page
        page.goto(BASE_URL, wait_until="domcontentloaded")
        self._quiet_wait()

        if self.is_logged_in():
            LOG.info("Already signed in (session restored from %s)", self.profile_dir)
            return

        # Open the masthead login panel if the fields are not on screen yet.
        email_field = page.locator("#Email")
        if email_field.count() == 0 or not email_field.first.is_visible():
            try:
                page.get_by_role("link", name=re.compile(r"^\s*Log in\s*$", re.I)).first.click()
                page.wait_for_timeout(1_500)
            except Exception as exc:
                raise RuntimeError(f"Could not open the login panel: {exc}") from exc

        page.locator("#Email").first.fill(email)
        page.locator("#Password").first.fill(password)

        submit = page.locator("input[type=submit][value='Log in' i], button:has-text('Log in')").first
        try:
            submit.click()
        except Exception:
            page.locator("#Password").first.press("Enter")
        self._quiet_wait()
        page.wait_for_timeout(2_500)

        if not self.is_logged_in():
            raise RuntimeError(
                "Login did not succeed - check MCMASTER_EMAIL / MCMASTER_PASSWORD."
            )
        LOG.info("Signed in as %s (session saved to %s)", email, self.profile_dir)

    def register(self, name: str, email: str, password: str, *,
                 accept_terms: bool = True) -> None:
        """
        Create a McMaster account and sign in.

        Fills the "Create login" form (Name / Email / Password, min 8 chars),
        ticks "I agree to the Terms and Conditions" and submits.  The caller is
        expected to have obtained the user's consent to these terms.
        """
        page = self.page
        page.goto(BASE_URL, wait_until="domcontentloaded")
        self._quiet_wait()
        if self.is_logged_in():
            LOG.info("Already signed in; skipping account creation")
            return

        # Open the login panel and switch to the "Create login" form if needed.
        if page.locator("#Name").count() == 0 or not page.locator("#Name").first.is_visible():
            try:
                page.get_by_role("link", name=re.compile(r"^\s*Log in\s*$", re.I)).first.click()
            except Exception:
                pass
            page.wait_for_timeout(1_500)
            try:
                page.get_by_text(re.compile(r"^\s*Create login\s*$", re.I)).last.click()
            except Exception:
                try:
                    page.get_by_role("button", name=re.compile("create login", re.I)).first.click()
                except Exception as exc:
                    raise RuntimeError(f"Could not open the 'Create login' form: {exc}") from exc
            page.wait_for_timeout(2_500)

        # The login form and the create-login form share the #Email/#Password ids,
        # so scope to the container that also holds #Name (unique to this form).
        box = page.locator("div:has(#Name):has(#Email)").last
        box.locator("#Name").fill(name)
        box.locator("#Email").fill(email)
        box.locator("#Password").fill(password)
        if accept_terms:
            try:
                box.get_by_text(re.compile(r"I agree to the Terms", re.I)).first.click()
            except Exception:
                LOG.debug("Terms checkbox not found via label; leaving as-is")

        submit = box.locator(
            "input[type=submit][value*='Create' i], button:has-text('Create Login')"
        ).first
        try:
            submit.click()
        except Exception:
            box.locator("#Password").press("Enter")
        self._quiet_wait()
        page.wait_for_timeout(3_500)

        if not self.is_logged_in():
            raise RuntimeError(
                "Signup did not complete - the username may already exist, or a "
                "CAPTCHA / verification step is required."
            )
        LOG.info("Account created and signed in: %s", email)

    def interactive_login(self) -> None:
        """Open the site so the user can sign in; the session is stored in the profile."""
        self.page.goto(BASE_URL, wait_until="domcontentloaded")
        self.page.bring_to_front()
        print("\nA browser window is open.  Sign in to mcmaster.com, then come back")
        print("here and press Enter.  The session is saved in:")
        print(f"    {self.profile_dir}\n")
        try:
            input("Press Enter when you are logged in... ")
        except EOFError:
            pass
        LOG.info("Session saved to %s", self.profile_dir)


# --------------------------------------------------------------------------- #
# Interactive input
# --------------------------------------------------------------------------- #
def prompt_for_mpns(valid: list[str]) -> list[str]:
    """Ask for an MPN until the user gives one that is in TableA."""
    print("\nMPNs available in TableA:")
    for mpn in valid:
        print(f"    {mpn}")
    while True:
        try:
            raw = input("\nEnter an MPN (or 'all' / 'q' to quit): ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return []
        if not raw or raw.lower() in {"q", "quit", "exit"}:
            return []
        if raw.lower() == "all":
            return list(valid)
        key = normalise_mpn(raw)
        if not MPN_RE.match(key):
            print(f"  '{raw}' does not look like a McMaster part number (e.g. 91864A028).")
        elif key not in valid:
            print(f"  '{key}' is not in TableA.  Choose one of: {', '.join(valid)}")
        else:
            return [key]


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        description="Look up Material and RoHS Compliance on mcmaster.com and update TableA.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("Usage", 1)[1] if "Usage" in __doc__ else None,
    )
    ap.add_argument("mpns", nargs="*", help="MPN(s) to look up; if omitted you are prompted")
    ap.add_argument("--all", action="store_true", help="look up every MPN in TableA")
    ap.add_argument("--headed", action="store_true", help="show the browser window")
    ap.add_argument("--login", action="store_true",
                    help="open a browser to sign in once (session is saved), then exit")
    ap.add_argument("--signup", action="store_true",
                    help="create a McMaster account (needs --email/--password) then continue")
    ap.add_argument("--name", default=None,
                    help="account holder name for --signup (default: email local part)")
    ap.add_argument("--email", default=None,
                    help=f"McMaster account email (or set ${ENV_EMAIL})")
    ap.add_argument("--password", default=None,
                    help=f"McMaster account password (or set ${ENV_PASSWORD})")
    ap.add_argument("--table", type=Path, default=TABLE_A, help="input TableA.xlsx")
    ap.add_argument("--out", type=Path, default=None,
                    help="output folder (default: ~/Documents)")
    ap.add_argument("--profile", type=Path, default=PROFILE_DIR,
                    help="browser profile folder (keeps the login session)")
    ap.add_argument("-v", "--verbose", action="store_true", help="debug logging")
    return ap


def main(argv: Optional[list[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    ensure_table_a(args.table)
    out_path = (args.out or documents_dir()) / OUTPUT_NAME
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Credentials: CLI flags take priority, then environment variables.
    email = args.email or os.environ.get(ENV_EMAIL)
    password = args.password or os.environ.get(ENV_PASSWORD)
    credentials = (email, password) if email and password else None

    try:
        table = load_table(args.table, out_path)
    except RuntimeError as exc:
        print(f"Error: {exc}")
        return 2

    # --login is a one-off helper that must run in a visible browser.
    if args.login:
        try:
            scraper = Scraper(headless=False, profile_dir=args.profile)
        except RuntimeError as exc:
            print(f"Error: {exc}")
            return 1
        try:
            if credentials:
                scraper.login(*credentials)
                print("Login OK.")
            else:
                scraper.interactive_login()
        except RuntimeError as exc:
            print(f"Login failed: {exc}")
            return 1
        finally:
            scraper.close()
        return 0

    valid = table.mpns
    if not valid:
        print("TableA has no MPNs to look up.")
        return 0

    if args.all:
        targets = list(valid)
    elif args.mpns:
        targets = [normalise_mpn(m) for m in args.mpns]
        bad = [m for m in targets if m not in table.rows_by_mpn]
        if bad:
            print(f"Not in TableA: {', '.join(bad)}")
            print(f"Valid MPNs: {', '.join(valid)}")
            targets = [m for m in targets if m in table.rows_by_mpn]
    else:
        targets = prompt_for_mpns(valid)

    if not targets:
        print("Nothing to do.")
        return 0

    try:
        scraper = Scraper(headless=not args.headed, profile_dir=args.profile,
                          credentials=credentials)
        if credentials:
            if args.signup and not scraper.is_logged_in():
                print(f"Creating McMaster account for {credentials[0]} ...")
                try:
                    scraper.register(args.name or credentials[0].split("@")[0], *credentials)
                    print("Signup OK.")
                except RuntimeError as exc:
                    print(f"Signup: {exc}")
            try:
                scraper.login(*credentials)
                print(f"Signed in as {credentials[0]}.")
            except RuntimeError as exc:
                print(f"Login failed: {exc}")
                return 1
    except RuntimeError as exc:
        print(f"Error: {exc}")
        return 1

    failures = 0
    try:
        for i, mpn in enumerate(targets):
            if i:
                time.sleep(random.uniform(MIN_DELAY_BETWEEN_PARTS_S, MAX_DELAY_BETWEEN_PARTS_S))
            try:
                specs = scraper.lookup(mpn)
                table.set_cell(mpn, "Material", specs.material)
                table.set_cell(mpn, "RoHS", specs.rohs)
                print(f"[OK]      {mpn}: Material = {specs.material} | RoHS = {specs.rohs}")
            except McMasterBlocked as exc:
                failures += 1
                print(f"[BLOCKED] {mpn}: {exc}")
            except Exception as exc:  # noqa: BLE001 - keep going through the list
                failures += 1
                LOG.error("%s: %s", mpn, exc)
                table.set_cell_if_unverified(mpn, "Material", "Lookup failed")
                table.set_cell_if_unverified(mpn, "RoHS", "Lookup failed")
            # Persist after every MPN so a crash never loses completed rows.
            style_sheet(table.sheet)
            table.save(out_path)
    finally:
        scraper.close()

    print(f"\nSaved: {out_path}")
    if failures:
        print(f"{failures} of {len(targets)} lookup(s) did not complete.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
