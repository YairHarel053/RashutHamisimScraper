# -*- coding: utf-8 -*-
"""
nadlan_scraper.py
=================
Automation for extracting real-estate transactions from the Israeli Tax
Authority information system (svinfonadlan2010 - "executed transactions")
and producing an Excel file.

The scraper walks every row of the results grid (including pagination),
opens the detail page of each transaction (perutOfDira.aspx), extracts all
fields by their stable element id, and produces an Excel file with two
sheets:
  1) Main  - the 11 primary columns.
  2) Extended - a grouped 32-column table (category groups, calculation
     coefficients and adjusted-value formulas), matching the target workbook.

Principles:
  * Drives a real, managed browser (Microsoft Edge) using a dedicated
    persistent profile, so the existing corporate session/policies apply and
    behavior matches normal browsing.
  * Humanized, polite behavior: random delays, scrolling, mouse movement.
    This reduces server load and makes the flow more stable.
  * The scraper does NOT bypass CAPTCHA or any security control. It only
    automates actions you could perform manually on public data.

==========================
How to run
==========================
1) One-time setup:
       pip install -r requirements.txt
       # Chromium download is blocked by the corporate proxy, so the default
       # browser is Microsoft Edge (already installed and managed).

2) Run the desktop UI (recommended):
       python nadlan_ui.py     # opens a local web page in your browser
   Or run the console version:
       python nadlan_scraper.py

   Default mode (BROWSER_MODE = "own_profile"):
     * Opens an Edge window with a dedicated profile (kept between runs).
     * Log in and run the search until the results grid is shown.
     * Return to the console and press Enter - the script verifies the grid
       exists and starts. (If you are on the wrong page it warns and retries.)
     * When done an Excel file (output.xlsx) with two sheets is saved.

   Alternative mode (BROWSER_MODE = "existing"):
     * Launch Edge/Chrome with a debugging port, navigate to the results
       page, then run the script.
"""

from __future__ import annotations

import random
import re
import os
import sys
import time
from datetime import datetime
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from playwright.sync_api import (
    Page,
    Locator,
    sync_playwright,
)

# ======================================================================
#  Logger hook - lets a GUI capture the messages
# ======================================================================
#  All messages go through LOG(). By default they are printed to stderr
#  (console mode). The UI replaces the callback via set_logger().

def _default_log(msg: str) -> None:
    print(msg, file=sys.stderr)


LOG = _default_log


def set_logger(fn) -> None:
    """Replace the log function (e.g. for a GUI). None restores the default."""
    global LOG
    LOG = fn or _default_log


# ======================================================================
#  CONFIG - tune here without touching the logic
# ======================================================================

# --- Browser launch mode ----------------------------------------------
# "own_profile" (recommended): the script opens a browser window with its own
#                              dedicated profile, lets you log in and run the
#                              search, and waits for your confirmation (Enter)
#                              before starting. The login is kept between runs.
# "existing":                  attach to an already-running browser started
#                              with --remote-debugging-port.
BROWSER_MODE = "own_profile"

# Dedicated profile directory (kept between runs -> stores login/cookies).
# Stored under LOCALAPPDATA, NOT inside the OneDrive project folder, because
# OneDrive file syncing locks the profile and breaks the browser launch.
USER_PROFILE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "nadlan_scraper",
    "browser-profile",
)

# Browser channel used in "own_profile" mode:
#   "msedge" -> Microsoft Edge installed on the machine (default; the managed
#               corporate browser, allowed through the ZIM proxy).
#   "chrome" -> Google Chrome installed on the machine.
#   None     -> Chromium downloaded by Playwright (blocked by the ZIM policy).
BROWSER_CHANNEL = "msedge"

# Attach to an existing browser (only relevant when BROWSER_MODE = "existing").
CDP_URL = "http://localhost:9222"

# Start page to navigate to automatically in "own_profile" mode.
START_URL = (
    "https://nadlan.taxes.gov.il/svinfonadlan2010/"
    "startpageNadlanNewDesign.aspx"
)
HEADLESS = False

# Output file (Excel).
OUTPUT_FILE = "output.xlsx"

# Title text for the tables. {info} is replaced automatically with the current
# search details (property type, dates, city, etc.) read from the results page.
TITLE_TEMPLATE = "עסקאות שבוצעו - רשות המיסים   |   {info}"

# --- Selectors - matched to the exact svinfonadlan2010 markup ----------
# Results grid (RadGrid/asp GridView) on InfoNadlanPerutWithMap.aspx.
RESULTS_GRID_SELECTOR = "#ContentUsersPage_GridMultiD1"

# Entry links to each transaction's detail page (postback -> perutOfDira.aspx).
# The ids are LogShow_0 .. LogShow_11 (up to 12 transactions per page).
ROW_LINK_SELECTOR = "a[id^='ContentUsersPage_GridMultiD1_LogShow_']"

# Pager row at the bottom of the grid. Links are page numbers; the current
# page appears as a <span>, not a link.
PAGER_LINK_SELECTOR = "tr.tabelPages a"

# Current page label ("דף 1 מתוך 6").
PAGE_LABEL_SELECTOR = "#lblPage"

# Current search criteria box (property type, dates, city, build years).
SEARCH_INFO_SELECTOR = "#lblInfo"

# "Back" button on the detail page -> returns to the results grid.
BACK_BUTTON_SELECTOR = "#ContentUsersPage_btHazara"

# Marker that confirms we are on the detail page (perutOfDira.aspx).
DETAIL_MARKER = "#ContentUsersPage_lblGush"

# ----------------------------------------------------------------------
#  Detail page (perutOfDira.aspx) field map by stable id.
#  Values are read directly by the span id, so extraction is precise.
#  key = internal name | value = DOM id.
#  (The trailing comment on each line is the Hebrew field label.)
# ----------------------------------------------------------------------
DETAIL_FIELDS: dict[str, str] = {
    "ezor":            "ContentUsersPage_lblEzor",        # area
    "gush_chelka":     "ContentUsersPage_lblGush",        # block/parcel (full)
    "date":            "ContentUsersPage_lblTarIska",     # transaction date
    "yeshuv":          "ContentUsersPage_lblYeshuv",      # city
    "rechov":          "ContentUsersPage_lblRechov",      # street
    "bayit":           "ContentUsersPage_lblBayit",       # house number
    "knisa":           "ContentUsersPage_lblKnisa",       # entrance
    "dira":            "ContentUsersPage_lblDira",        # apartment
    "mcir_mozhar":     "ContentUsersPage_lblMcirMozhar",  # declared price (ILS)
    "mcir_mozhar_dlr": "ContentUsersPage_lblMcirMozharDlr",  # declared price (USD)
    "mcir_morach":     "ContentUsersPage_lblMcirMorach",  # assessed price (ILS)
    "mcir_morach_dlr": "ContentUsersPage_lblMcirMorachDlr",  # assessed price (USD)
    "shetach_bruto":   "ContentUsersPage_lblShetachBruto",  # gross area
    "shetach_neto":    "ContentUsersPage_lblShetachNeto",   # net area
    # The two area labels are dynamic: usually "שטח ברוטו"/"שטח נטו" but
    # sometimes "שטח ארנונה"/"שטח רשום". We read them to label remarks correctly.
    "name_shetach_bruto": "ContentUsersPage_lblNameShetachBruto",  # gross-area label
    "name_shetach_neto":  "ContentUsersPage_lblNameShetachNeto",   # net-area label
    "shnat_bniya":     "ContentUsersPage_lblShnatBniya",  # build year
    "mechir_cheder":   "ContentUsersPage_lblMechirCheder",  # price per room
    "mechir_lmr":      "ContentUsersPage_lblMechirLmr",   # price per sqm
    "mis_hadarim":     "ContentUsersPage_lblMisHadarim",  # number of rooms
    "koma":            "ContentUsersPage_lblKoma",        # floor
    "mis_komot":       "ContentUsersPage_lblMisKomot",    # number of floors
    "dirot_bnyn":      "ContentUsersPage_lblDirotBnyn",   # apartments in building
    "hanaya":          "ContentUsersPage_lblHanaya",      # attached parking
    "malit":           "ContentUsersPage_lblMalit",       # elevator
    "gag":             "ContentUsersPage_lblGag",         # roof
    "machsan":         "ContentUsersPage_lblMachsan",     # storage
    "hzer":            "ContentUsersPage_lblHzer",        # yard
    "migrash":         "ContentUsersPage_lblMigrash",     # plot
    "glrya":           "ContentUsersPage_lblGlrya",       # gallery
    "sug_iska":        "ContentUsersPage_lblSugIska",     # transaction type
    "tifkud_bnyn":     "ContentUsersPage_lblTifkudBnyn",  # building function
    "tifkud_ychida":   "ContentUsersPage_lblTifkudYchida",  # unit function
    "shuma_halakim":   "ContentUsersPage_lblShumaHalakim",  # assessment parts
    "mofa_gush":       "ContentUsersPage_lblMofaGush",    # block appearances
    "tava":            "ContentUsersPage_lblTava",        # per zoning plan
    "mahut_zchut":     "ContentUsersPage_lblMahutZchut",  # nature of rights
}

# Hebrew headers for the "extended" columns (sheet 2), in display order.
DETAIL_LABELS: dict[str, str] = {
    "ezor":            "אזור",
    "gush_chelka":     "גו\"ח",
    "date":            "תאריך העסקה",
    "yeshuv":          "ישוב",
    "rechov":          "רחוב",
    "bayit":           "בית",
    "knisa":           "כניסה",
    "dira":            "דירה",
    "mcir_mozhar":     "מחיר מוצהר (₪)",
    "mcir_mozhar_dlr": "מחיר מוצהר ($)",
    "mcir_morach":     "מחיר מוערך (₪)",
    "mcir_morach_dlr": "מחיר מוערך ($)",
    "shetach_bruto":   "שטח ברוטו",
    "shetach_neto":    "שטח נטו",
    "shnat_bniya":     "שנת בניה",
    "mechir_cheder":   "מחיר לחדר",
    "mechir_lmr":      "מחיר למ\"ר",
    "mis_hadarim":     "מספר חדרים",
    "koma":            "קומה",
    "mis_komot":       "מספר קומות",
    "dirot_bnyn":      "דירות בבנין",
    "hanaya":          "חניה צמודה ל",
    "malit":           "מעלית",
    "gag":             "גג",
    "machsan":         "מחסן",
    "hzer":            "חצר",
    "migrash":         "מגרש",
    "glrya":           "גלריה",
    "sug_iska":        "סוג עסקה",
    "tifkud_bnyn":     "תפקוד הבנין",
    "tifkud_ychida":   "תפקוד היחידה",
    "shuma_halakim":   "שומה חלקים",
    "mofa_gush":       "מופעי גו\"ח",
    "tava":            "לפי תבע",
    "mahut_zchut":     "מהות הזכות",
}

# Headers for the 11 primary columns (sheet 1), in the fixed required order.
MAIN_HEADERS = [
    "גוש/חלקה",
    "כתובת",
    "תאריך",
    "גודל נטו (כולל אקוויולנטי)",
    "קומה",
    "מחיר",
    "מס חדרים",
    "שנת בנייה",
    "מס דירות בבניין",
    "מחיר למ\"ר",
    "התייחוסיות",
]

# ----------------------------------------------------------------------
#  Extended sheet (sheet 2) layout - mirrors the "עסקאות נדל״ן" workbook:
#  a grouped 32-column table with two header rows (category groups on top,
#  column names below), calculation coefficients (default 1) and two
#  adjusted-value formula columns.
# ----------------------------------------------------------------------

# Category groups shown on the first header row (label, first col, last col).
EXT_GROUPS: list[tuple[str, int, int]] = [
    ("תאריך", 1, 1),
    ("מיקום", 2, 7),
    ("מאפייני הנכס", 8, 17),
    ("מאפייני המבנה", 18, 21),
    ("העסקה", 22, 26),
    ("תחשיב", 27, 32),
]

# Column names shown on the second header row (32 columns, in display order).
EXT_HEADERS = [
    "יום מכירה",                 # 1  A
    "כתובת",                     # 2  B
    "דירה",                      # 3  C
    "ישוב",                      # 4  D
    "גוש",                       # 5  E
    "חלקה",                      # 6  F
    "תת חלקה",                   # 7  G
    "סוג נכס",                   # 8  H
    "מהות",                      # 9  I
    "תפקוד יחידה",               # 10 J
    "חדרים",                     # 11 K
    "קומה",                      # 12 L
    "שטח ברוטו",                 # 13 M
    "שטח נטו",                   # 14 N
    "גג",                        # 15 O
    "מחסן",                      # 16 P
    "חצר",                       # 17 Q
    "מספר קומות",                # 18 R
    "דירות בבנין",               # 19 S
    "שנת בניה",                  # 20 T
    "תפקוד בנין",                # 21 U
    "שווי מכירה בש\"ח",          # 22 V
    "חלק נמכר",                  # 23 W
    "מחיר למ\"ר ברוטו",          # 24 X
    "מחיר למ\"ר נטו",            # 25 Y
    "מחיר מוצהר",                # 26 Z
    "מקדם זמן",                  # 27 AA
    "מקדם קומה",                 # 28 AB
    "מקדם שוליות",               # 29 AC
    "מקדם מצב פיזי",             # 30 AD
    "שווי מתואם למ\"ר ברוטו",    # 31 AE  = X * AA * AB * AC * AD
    "שווי מתואם למ\"ר נטו",      # 32 AF  = Y * AA * AB * AC * AD
]

# --- "Humanized behavior" parameters -----------------------------------
MIN_DELAY = 0.6      # seconds - minimum pause between actions
MAX_DELAY = 1.8      # seconds - maximum pause between actions
MAX_PAGES = 200      # safety cap (upper bound on number of pages)

# Global speed knob: every random pause is multiplied by this factor.
#   1.0  -> original, slow-but-polite pace.
#   0.3  -> ~3x faster (recommended; still keeps small random gaps).
#   0.0  -> no artificial pauses at all (fastest; least "human").
# Network/page-load waits are NOT affected - only the artificial delays.
SPEED_FACTOR = 0.3


# ======================================================================
#  "Humanized behavior" helpers
# ======================================================================

def human_pause(a: float = MIN_DELAY, b: float = MAX_DELAY) -> None:
    """Short random pause to avoid load and mimic a human browsing pace.
    Scaled by SPEED_FACTOR so the whole run can be sped up from one knob."""
    if SPEED_FACTOR <= 0:
        return
    time.sleep(random.uniform(a, b) * SPEED_FACTOR)


def human_move(page: Page) -> None:
    """Subtle random mouse movement."""
    try:
        page.mouse.move(
            random.randint(100, 900),
            random.randint(150, 700),
            steps=random.randint(5, 15),
        )
    except Exception:
        pass


def human_scroll(page: Page) -> None:
    """Light scroll to mimic reading."""
    try:
        page.mouse.wheel(0, random.randint(120, 400))
    except Exception:
        pass


def humanized_click(page: Page, locator: Locator) -> None:
    """Click preceded by mouse movement and a pause."""
    human_move(page)
    human_pause(0.2, 0.6)
    locator.scroll_into_view_if_needed(timeout=5000)
    locator.click()
    human_pause()


# ======================================================================
#  Value normalization
# ======================================================================

def clean_text(value: Optional[str]) -> str:
    if not value:
        return ""
    # Collapse whitespace and trim
    return re.sub(r"\s+", " ", value).strip()


def normalize_date(raw: str) -> str:
    """
    Convert a date to day.month.year format (17.04.2025).
    Supports a few common formats; if unrecognized it is returned cleaned.
    """
    raw = clean_text(raw)
    if not raw:
        return ""

    # Pull the part that looks like a date
    m = re.search(r"(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})", raw)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        try:
            dt = datetime(int(y), int(mo), int(d))
            return dt.strftime("%d.%m.%Y")
        except ValueError:
            pass

    # Year-month-day format
    m = re.search(r"(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})", raw)
    if m:
        y, mo, d = m.groups()
        try:
            dt = datetime(int(y), int(mo), int(d))
            return dt.strftime("%d.%m.%Y")
        except ValueError:
            pass

    return raw


def _year_of_date(raw: str) -> Optional[int]:
    """Extract the 4-digit year from a transaction date string (any format)."""
    normalized = normalize_date(raw)
    m = re.search(r"(\d{4})", normalized)
    return int(m.group(1)) if m else None


def normalize_gush_chelka(raw: str) -> str:
    """
    Return the block/parcel string as shown (full format, e.g.
    017015-0014-003-00). If it already matches the full format, keep it.
    """
    raw = clean_text(raw)
    # If already in the full hyphenated format - return as is
    m = re.search(r"\d{4,6}-\d{3,4}-\d{2,3}-\d{2}", raw)
    if m:
        return m.group(0)
    return raw


def strip_decimal(value: str) -> str:
    """'2.0' -> '2', '5.0' -> '5', but '2.5' stays '2.5'."""
    value = clean_text(value)
    m = re.fullmatch(r"(\d+)\.0+", value)
    return m.group(1) if m else value


def num_value(value: str) -> Optional[float]:
    """Extract the first numeric value from a string (supports commas)."""
    value = clean_text(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", value)
    return float(m.group(0)) if m else None


def is_valid(value: Optional[str]) -> bool:
    """Whether the value is real (not empty and not '--')."""
    return value is not None and clean_text(value) not in ("", "--", "---")


def fmt_shekel(value: str) -> str:
    """Price format: '1580000' / '1,580,000' -> '₪ 1,580,000'. 0/empty -> ''."""
    if not is_valid(value):
        return ""
    n = num_value(value)
    if n is None:
        return clean_text(value)
    if n == 0:
        return ""
    return "₪ " + f"{int(round(n)):,}"


# ======================================================================
#  Extract transaction details by stable id (perutOfDira.aspx)
# ======================================================================

def get_text_by_id(page: Page, element_id: str) -> str:
    """Read text by id; returns '' if it does not exist."""
    loc = page.locator(f"#{element_id}")
    if loc.count() == 0:
        return ""
    try:
        return clean_text(loc.first.inner_text())
    except Exception:  # noqa: BLE001
        return ""


def extract_detail_by_ids(page: Page) -> dict[str, str]:
    """Extract every detail-page field using the ids in DETAIL_FIELDS."""
    return {key: get_text_by_id(page, el_id) for key, el_id in DETAIL_FIELDS.items()}


# ======================================================================
#  Build output rows
# ======================================================================

def build_address(raw: dict[str, str]) -> str:
    """Assemble a readable address from: city, street+house, entrance, apt."""
    parts: list[str] = []
    if is_valid(raw.get("yeshuv")):
        parts.append(clean_text(raw["yeshuv"]))
    street = clean_text(raw["rechov"]) if is_valid(raw.get("rechov")) else ""
    if is_valid(raw.get("bayit")):
        street = (street + " " + clean_text(raw["bayit"])).strip()
    if street:
        parts.append(street)
    if is_valid(raw.get("knisa")):
        parts.append("כניסה " + clean_text(raw["knisa"]))
    if is_valid(raw.get("dira")):
        parts.append("דירה " + clean_text(raw["dira"]))
    return ", ".join(parts)


def build_floor(raw: dict[str, str]) -> str:
    """Floor in 'X out of Y' format (e.g. '2 מתוך 7')."""
    koma = strip_decimal(raw.get("koma", ""))
    komot = strip_decimal(raw.get("mis_komot", ""))
    if is_valid(koma) and is_valid(komot):
        return f"{koma} מתוך {komot}"
    return koma if is_valid(koma) else ""


def build_price_per_sqm(raw: dict[str, str]) -> str:
    """Price per sqm: the page field if > 0, otherwise price / net area."""
    page_val = num_value(raw.get("mechir_lmr", ""))
    if page_val and page_val > 0:
        return fmt_shekel(raw["mechir_lmr"])
    price = num_value(raw.get("mcir_mozhar", ""))
    neto = num_value(raw.get("shetach_neto", ""))
    if price and neto and neto > 0:
        return fmt_shekel(str(price / neto))
    return ""


def _is_partial_share(value: str) -> bool:
    """True when the 'assessment parts' value represents less than the whole.
    Handles fractions ('1/2', '3/4'), percentages ('50%') and plain numbers.
    A whole share ('1/1', '1', '100%') returns False."""
    value = clean_text(value)
    if not value:
        return False
    # Fraction a/b -> partial when numerator < denominator.
    m = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", value)
    if m:
        num, den = float(m.group(1)), float(m.group(2))
        if den > 0:
            return num < den
        return False
    # Percentage -> partial when below 100%.
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", value)
    if m:
        return float(m.group(1)) < 100
    # Plain number -> 1 (or 100) is whole; a value strictly between 0 and 1
    # (e.g. 0.5) is partial.
    n = num_value(value)
    if n is not None:
        if n in (1, 100):
            return False
        return 0 < n < 1
    return False


def build_remarks(raw: dict[str, str]) -> str:
    """Collect 'remarks': area/price differences, rights, physical extras, elevator, parking."""
    extras: list[str] = []

    # Sold "on paper": transaction happened before the building was built.
    deal_year = _year_of_date(raw.get("date", ""))
    build_year = num_value(raw.get("shnat_bniya", ""))
    if deal_year is not None and build_year is not None and build_year > 0:
        gap = int(build_year) - deal_year
        if gap > 0:
            unit = "שנה" if gap == 1 else "שנים"
            extras.append(f"נמכרה על הנייר ({gap} {unit} לפני הבנייה)")

    # Difference between the two area values -> flag it explicitly.
    # The label names are dynamic ("ברוטו"/"נטו" or "ארנונה"/"רשום").
    # Only a real difference counts: BOTH values must actually exist (not
    # empty / "--" / 0). If one side is missing it is NOT a mismatch.
    bruto = num_value(raw.get("shetach_bruto", ""))
    neto = num_value(raw.get("shetach_neto", ""))
    both_present = (
        is_valid(raw.get("shetach_bruto")) and is_valid(raw.get("shetach_neto"))
        and bruto is not None and neto is not None
        and bruto > 0 and neto > 0
    )
    if both_present and bruto != neto:
        b_lbl = clean_text(raw.get("name_shetach_bruto", "")) or "ברוטו"
        n_lbl = clean_text(raw.get("name_shetach_neto", "")) or "נטו"
        b_txt = strip_decimal(clean_text(raw.get("shetach_bruto", "")))
        n_txt = strip_decimal(clean_text(raw.get("shetach_neto", "")))
        extras.append(f"אי התאמה בין {b_lbl} ל{n_lbl} {b_lbl}: {b_txt} {n_lbl}: {n_txt}")

    # Difference between declared and assessed price -> flag it explicitly.
    mozhar = num_value(raw.get("mcir_mozhar", ""))
    morach = num_value(raw.get("mcir_morach", ""))
    if mozhar is not None and morach is not None and mozhar != morach:
        mo_txt = fmt_shekel(raw.get("mcir_mozhar", "")) or "0"
        ma_txt = fmt_shekel(raw.get("mcir_morach", "")) or "0"
        extras.append(f"הפרש מחיר (מוצהר {mo_txt} / מוערך {ma_txt})")

    # Not sold in full: the "assessment parts" field (שומה חלקים) shows the
    # share that changed hands. Anything less than the whole (e.g. "1/2",
    # "3/4", or a percentage below 100) means a partial sale -> flag it.
    halakim = clean_text(raw.get("shuma_halakim", ""))
    if is_valid(halakim) and _is_partial_share(halakim):
        extras.append(f"לא נמכרה בשלמות (שומה חלקים: {halakim})")

    # NOTE: only computed/notable remarks are kept here (sold on paper, area
    # mismatch, price mismatch). Fields already shown on the site itself
    # (ownership, roof, storage, yard, plot, gallery, elevator, parking) are
    # intentionally NOT repeated - they remain in the extended sheet.
    return ". ".join(extras)


def _num_or_blank(value: str):
    """Return an int/float for a numeric field, '' when empty/'--'."""
    if not is_valid(value):
        return ""
    n = num_value(value)
    if n is None:
        return clean_text(value)
    return int(n) if n == int(n) else n


def split_gush_chelka(raw: str):
    """Split a block/parcel string into (block, parcel, sub-parcel) integers.
    Reads the first three numeric groups (e.g. '014051-0091-000-00' ->
    14051, 91, 0). Missing parts come back as ''."""
    nums = re.findall(r"\d+", clean_text(raw))
    gush = int(nums[0]) if len(nums) > 0 else ""
    chelka = int(nums[1]) if len(nums) > 1 else ""
    tat = int(nums[2]) if len(nums) > 2 else ""
    return gush, chelka, tat


def build_street_address(raw: dict[str, str]) -> str:
    """Street + house number only (no city / entrance / apartment)."""
    street = clean_text(raw["rechov"]) if is_valid(raw.get("rechov")) else ""
    if is_valid(raw.get("bayit")):
        street = (street + " " + clean_text(raw["bayit"])).strip()
    return street


def _to_datetime(raw: str):
    """Parse a transaction date to a datetime (for a real Excel date cell).
    Falls back to the cleaned string when it cannot be parsed."""
    norm = normalize_date(raw)
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", norm)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        try:
            return datetime(y, mo, d)
        except ValueError:
            pass
    return norm


def _share_value(value: str) -> float:
    """Sold share as a fraction of the whole (1.0 = full sale).
    Handles fractions ('1/2'), percentages ('50%') and plain numbers."""
    value = clean_text(value)
    if not value:
        return 1.0
    m = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", value)
    if m:
        num, den = float(m.group(1)), float(m.group(2))
        return num / den if den else 1.0
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", value)
    if m:
        return float(m.group(1)) / 100
    n = num_value(value)
    if n is not None and 0 < n <= 1:
        return n
    return 1.0


def build_extended_row(raw: dict[str, str]) -> list:
    """Build the 32 columns of the extended sheet, matching the grouped
    workbook layout. Columns 31/32 are left as None here - the writer fills
    them with per-row Excel formulas.

    Field mapping is best-effort: the Tax-Authority detail page does not
    expose every column of the target workbook, so a few descriptive columns
    (סוג נכס / מהות / תפקוד יחידה) reuse the closest available field.
    """
    gush, chelka, tat = split_gush_chelka(raw.get("gush_chelka", ""))

    price = num_value(raw.get("mcir_mozhar", ""))
    if not price:
        price = num_value(raw.get("mcir_morach", ""))
    bruto = num_value(raw.get("shetach_bruto", ""))
    neto = num_value(raw.get("shetach_neto", ""))
    per_bruto = int(round(price / bruto)) if price and bruto else ""
    per_neto = int(round(price / neto)) if price and neto else ""

    return [
        _to_datetime(raw.get("date", "")),            # 1  יום מכירה
        build_street_address(raw),                    # 2  כתובת
        _num_or_blank(raw.get("dira", "")),           # 3  דירה
        clean_text(raw.get("yeshuv", "")),            # 4  ישוב
        gush,                                         # 5  גוש
        chelka,                                       # 6  חלקה
        tat,                                          # 7  תת חלקה
        clean_text(raw.get("tifkud_ychida", "")),     # 8  סוג נכס
        clean_text(raw.get("mahut_zchut", "")),       # 9  מהות
        clean_text(raw.get("tifkud_ychida", "")),     # 10 תפקוד יחידה
        _num_or_blank(raw.get("mis_hadarim", "")),    # 11 חדרים
        _num_or_blank(raw.get("koma", "")),           # 12 קומה
        _num_or_blank(raw.get("shetach_bruto", "")),  # 13 שטח ברוטו
        _num_or_blank(raw.get("shetach_neto", "")),   # 14 שטח נטו
        _num_or_blank(raw.get("gag", "")),            # 15 גג
        _num_or_blank(raw.get("machsan", "")),        # 16 מחסן
        _num_or_blank(raw.get("hzer", "")),           # 17 חצר
        _num_or_blank(raw.get("mis_komot", "")),      # 18 מספר קומות
        _num_or_blank(raw.get("dirot_bnyn", "")),     # 19 דירות בבנין
        _num_or_blank(raw.get("shnat_bniya", "")),    # 20 שנת בניה
        clean_text(raw.get("tifkud_bnyn", "")),       # 21 תפקוד בנין
        int(price) if price else "",                  # 22 שווי מכירה בש"ח
        _share_value(raw.get("shuma_halakim", "")),   # 23 חלק נמכר
        per_bruto,                                    # 24 מחיר למ"ר ברוטו
        per_neto,                                     # 25 מחיר למ"ר נטו
        _num_or_blank(raw.get("mcir_mozhar", "")),    # 26 מחיר מוצהר
        1,                                            # 27 מקדם זמן
        1,                                            # 28 מקדם קומה
        1,                                            # 29 מקדם שוליות
        1,                                            # 30 מקדם מצב פיזי
        None,                                         # 31 שווי מתואם ברוטו (formula)
        None,                                         # 32 שווי מתואם נטו (formula)
    ]


def build_main_row(raw: dict[str, str]) -> list[str]:
    """Build the 11 primary columns in the fixed required order."""
    return [
        normalize_gush_chelka(raw.get("gush_chelka", "")),   # block/parcel
        build_address(raw),                                  # address
        normalize_date(raw.get("date", "")),                 # date
        clean_text(raw.get("shetach_neto", "")),             # net area
        build_floor(raw),                                    # floor
        fmt_shekel(raw.get("mcir_mozhar", "")),              # price
        strip_decimal(raw.get("mis_hadarim", "")),           # number of rooms
        clean_text(raw.get("shnat_bniya", "")),              # build year
        clean_text(raw.get("dirot_bnyn", "")),               # apartments in building
        build_price_per_sqm(raw),                            # price per sqm
        build_remarks(raw),                                  # remarks
    ]


def build_search_info(raw: dict[str, str]) -> str:
    """Description text of the current search for the title (area / parcel / city)."""
    bits: list[str] = []
    if is_valid(raw.get("ezor")):
        bits.append(f"אזור: {clean_text(raw['ezor'])}")
    if is_valid(raw.get("gush_chelka")):
        bits.append(f"גו\"ח: {clean_text(raw['gush_chelka'])}")
    if is_valid(raw.get("yeshuv")):
        bits.append(f"ישוב: {clean_text(raw['yeshuv'])}")
    return "   |   ".join(bits) if bits else "חיפוש נוכחי"


# ======================================================================
#  Results page navigation (InfoNadlanPerutWithMap.aspx)
# ======================================================================

def wait_for_grid(page: Page, timeout: int = 20000) -> None:
    """Wait until the results grid is loaded."""
    page.wait_for_selector(RESULTS_GRID_SELECTOR, timeout=timeout)


def wait_for_detail(page: Page, timeout: int = 20000) -> None:
    """Wait until the detail page is loaded (identified by DETAIL_MARKER)."""
    page.wait_for_selector(DETAIL_MARKER, timeout=timeout)


def get_search_info(page: Page) -> str:
    """Read the current search criteria text from the results page (#lblInfo)."""
    loc = page.locator(SEARCH_INFO_SELECTOR)
    if loc.count() == 0:
        return ""
    try:
        txt = loc.first.inner_text()
    except Exception:  # noqa: BLE001
        return ""
    # Collapse to single spaces and replace newlines with a uniform separator
    txt = re.sub(r"\s*\n\s*", "   |   ", txt.strip())
    return re.sub(r"[ \t]+", " ", txt)


def get_total_pages(page: Page) -> int:
    """Total page count from '#lblPage' (\"דף 1 מתוך 6\"). Defaults to 1."""
    loc = page.locator(PAGE_LABEL_SELECTOR)
    if loc.count() > 0:
        m = re.search(r"מתוך\s*(\d+)", clean_text(loc.first.inner_text()))
        if m:
            return min(int(m.group(1)), MAX_PAGES)
    # Fallback: number of pager links + 1
    pager = page.locator(PAGER_LINK_SELECTOR)
    return max(1, min(pager.count() + 1, MAX_PAGES))


def get_current_page(page: Page) -> int:
    """Current page number from '#lblPage' (\"דף 1 מתוך 6\"). Defaults to 1."""
    loc = page.locator(PAGE_LABEL_SELECTOR)
    if loc.count() > 0:
        m = re.search(r"דף\s*(\d+)", clean_text(loc.first.inner_text()))
        if m:
            return int(m.group(1))
    return 1


def goto_page(page: Page, target: int) -> bool:
    """Navigate to a requested page in the grid pager. Returns True on success."""
    if get_current_page(page) == target:
        return True
    # Find the pager link whose text == the requested page number
    links = page.locator(PAGER_LINK_SELECTOR)
    for i in range(links.count()):
        link = links.nth(i)
        try:
            if clean_text(link.inner_text()) == str(target):
                humanized_click(page, link)
                wait_for_grid(page)
                human_pause(0.8, 1.6)
                return get_current_page(page) == target
        except Exception:  # noqa: BLE001
            continue
    return False


def count_rows_on_page(page: Page) -> int:
    """Number of transactions (LogShow links) on the current page."""
    return page.locator(ROW_LINK_SELECTOR).count()


def open_deal_by_index(page: Page, index: int) -> bool:
    """Click the entry link of a transaction by index (0-based) -> detail page."""
    link = page.locator(f"#ContentUsersPage_GridMultiD1_LogShow_{index}")
    if link.count() == 0:
        # Fallback: by LogShow order on the page
        links = page.locator(ROW_LINK_SELECTOR)
        if index >= links.count():
            return False
        link = links.nth(index)
    try:
        humanized_click(page, link.first)
        wait_for_detail(page)
        return True
    except Exception:  # noqa: BLE001
        return False


def go_back_to_results(page: Page) -> None:
    """Return from the detail page to the results grid ('Back' button)."""
    btn = page.locator(BACK_BUTTON_SELECTOR)
    try:
        if btn.count() > 0:
            humanized_click(page, btn.first)
        else:
            page.go_back()
    except Exception:  # noqa: BLE001
        page.go_back()
    try:
        wait_for_grid(page)
    except Exception:  # noqa: BLE001
        pass
    human_pause(0.8, 1.6)


def scrape_page_deals(
    page: Page,
    page_num: int,
    should_stop=None,
    on_deal=None,
) -> list[dict[str, str]]:
    """
    Process all transactions on the current page: for each index -> open
    detail, extract, go back, and make sure we returned to the right page
    before the next transaction.
    should_stop -- optional callback returning True to stop gracefully.
    on_deal    -- optional callback invoked with each collected deal right
                  after it is extracted (used for incremental saving).
    """
    deals: list[dict[str, str]] = []
    total = count_rows_on_page(page)
    LOG(f"[i] עמוד {page_num}: נמצאו {total} עסקאות.")

    for i in range(total):
        if should_stop is not None and should_stop():
            LOG("[i] עצירה לבקשת המשתמש.")
            break

        # Make sure we are on the right page (going back may reset to page 1)
        if get_current_page(page) != page_num:
            if not goto_page(page, page_num):
                LOG(f"[!] לא ניתן לחזור לעמוד {page_num}.")
                break

        human_scroll(page)

        if not open_deal_by_index(page, i):
            LOG(f"[!] לא ניתן לפתוח עסקה {i} בעמוד {page_num}.")
            continue

        human_pause(0.6, 1.2)
        try:
            raw = extract_detail_by_ids(page)
        except Exception as e:  # noqa: BLE001
            LOG(f"[!] שגיאת חילוץ (עמוד {page_num}, עסקה {i}): {e}")
            raw = {}

        go_back_to_results(page)

        if raw and any(is_valid(v) for v in raw.values()):
            deals.append(raw)
            if on_deal is not None:
                on_deal(raw)
            LOG(
                f"[i] נאספה עסקה {i + 1}/{total} (עמוד {page_num}): "
                f"{raw.get('gush_chelka', '')}"
            )

    return deals


# ======================================================================
#  Write the Excel file
# ======================================================================

def _write_sheet(ws, title: str, headers: list[str], rows: list[list[str]]) -> None:
    """Write a single sheet: merged title row, header row, and data rows."""
    ncols = len(headers)
    ws.sheet_view.rightToLeft = True

    # Title row (merged across all columns)
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    tcell = ws.cell(row=1, column=1)
    tcell.font = Font(bold=True, size=13, color="FFFFFF")
    tcell.fill = PatternFill("solid", fgColor="305496")
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Header row
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Data rows
    for ri, rowvals in enumerate(rows, start=3):
        for c, val in enumerate(rowvals, start=1):
            cell = ws.cell(row=ri, column=c, value=val)
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Auto column width (capped)
    for c in range(1, ncols + 1):
        maxlen = len(str(headers[c - 1]))
        for rowvals in rows:
            if c - 1 < len(rowvals):
                maxlen = max(maxlen, len(str(rowvals[c - 1])))
        ws.column_dimensions[get_column_letter(c)].width = min(max(maxlen + 2, 10), 45)

    ws.freeze_panes = "A3"


def _write_extended_sheet(ws, title: str, rows: list[list]) -> None:
    """Write the extended sheet in the grouped 32-column workbook layout:
    row 1 = merged title, row 2 = category groups, row 3 = column names,
    row 4+ = data (with formula columns for the two adjusted values)."""
    ncols = len(EXT_HEADERS)
    ws.sheet_view.rightToLeft = True

    # Title row (merged across all columns)
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    tcell = ws.cell(row=1, column=1)
    tcell.font = Font(bold=True, size=13, color="FFFFFF")
    tcell.fill = PatternFill("solid", fgColor="305496")
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Group header row (merged per category)
    group_fill = PatternFill("solid", fgColor="8EA9DB")
    for label, c0, c1 in EXT_GROUPS:
        cell = ws.cell(row=2, column=c0, value=label)
        if c1 > c0:
            ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c1)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = group_fill
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Column header row
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    for c, h in enumerate(EXT_HEADERS, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Data rows (start at row 4). Columns 31/32 hold the adjusted-value
    # formulas: X/Y (price per sqm) times the four coefficients (AA..AD).
    right = Alignment(horizontal="right", vertical="center")
    for ri, rowvals in enumerate(rows, start=4):
        for c, val in enumerate(rowvals, start=1):
            if c in (31, 32):
                continue
            cell = ws.cell(row=ri, column=c, value=val)
            cell.alignment = right
            if c == 1 and hasattr(val, "year"):
                cell.number_format = "DD/MM/YYYY"
        coef = f"AA{ri} * AB{ri} * AC{ri} * AD{ri}"
        ws.cell(row=ri, column=31, value=f"=(X{ri} * {coef})").alignment = right
        ws.cell(row=ri, column=32, value=f"=(Y{ri} * {coef})").alignment = right

    # Auto column width (capped)
    for c in range(1, ncols + 1):
        maxlen = len(str(EXT_HEADERS[c - 1]))
        for rowvals in rows:
            val = rowvals[c - 1] if c - 1 < len(rowvals) else None
            if val not in (None, ""):
                maxlen = max(maxlen, len(str(val)))
        ws.column_dimensions[get_column_letter(c)].width = min(max(maxlen + 2, 8), 30)

    ws.freeze_panes = "A4"


def make_output_path(directory: str, info: str = "") -> str:
    """
    Build a unique Excel path for a single run, so running the scraper several
    times on different searches produces several files instead of overwriting.
    Files are stored in an dedicated outputs folder to keep the project root
    clean and avoid polluting Git with generated Excel reports.
    Format: output_<city>_<YYYYMMDD_HHMMSS>.xlsx (city is optional).
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    label = _slug_from_info(info)
    name = f"output_{label}_{stamp}.xlsx" if label else f"output_{stamp}.xlsx"

    base_dir = directory
    if os.path.basename(os.path.normpath(directory)) != "outputs":
        base_dir = os.path.join(directory, "outputs")

    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, name)


def _slug_from_info(info: str) -> str:
    """Extract the city from the search info and make it filename-safe.
    The results-page label varies (ישוב / יישוב / עיר), so several patterns
    are tried before giving up."""
    info = clean_text(info)
    if not info:
        return ""
    # Try to pull the city name after any of the common labels.
    # Examples: "ישוב: עפולה", "יישוב : תל אביב", "עיר: חיפה".
    label = ""
    for pat in (
        r"יי?שוב\s*:\s*([^|]+)",   # ישוב / יישוב
        r"עיר\s*:\s*([^|]+)",         # עיר
    ):
        m = re.search(pat, info)
        if m:
            label = clean_text(m.group(1))
            if label:
                break
    if not label:
        return ""
    # Keep only the city itself: stop at the first separator that usually
    # precedes another field (comma / slash / digits like a house number).
    label = re.split(r"[,/]|\s\d", label)[0].strip()
    # Keep only safe characters; replace the rest with underscores.
    label = re.sub(r'[\\/:*?"<>|\s]+', "_", label)
    return label.strip("_")[:40]


def write_excel(deals_raw: list[dict[str, str]], path: str, info: str = "") -> None:
    """Create an Excel file with two sheets: main (11 cols) and extended (all)."""
    if not info and deals_raw:
        info = build_search_info(deals_raw[0])
    title = TITLE_TEMPLATE.format(info=info)

    wb = Workbook()

    # Sheet 1 -- the 11 primary columns
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "עסקאות שבוצעו"
    main_rows = [build_main_row(r) for r in deals_raw]
    _write_sheet(ws1, title, MAIN_HEADERS, main_rows)

    # Sheet 2 -- extended: grouped 32-column layout matching the target
    # workbook (category groups, coefficients and adjusted-value formulas).
    ws2 = wb.create_sheet("מורחב")
    ext_rows = [build_extended_row(r) for r in deals_raw]
    _write_extended_sheet(ws2, title, ext_rows)

    wb.save(path)


# ======================================================================
#  Main flow
# ======================================================================

def run(
    page: Page,
    on_progress=None,
    should_stop=None,
    output_path: Optional[str] = None,
    on_deal=None,
) -> tuple[list[dict[str, str]], str]:
    """
    Walk every results page and collect all transactions.
    Returns (deals, search_text).
    on_progress -- optional callback
                   (current_deal, total_deals, collected, page, total_pages,
                    page_deal, page_deals).
    should_stop -- optional callback returning True to stop gracefully.
    output_path -- optional Excel path; when set, the file is re-written after
                   every collected deal so progress is never lost on a crash.
    on_deal     -- optional callback (main_row, raw) for each unique deal,
                   used to feed a live table in the UI.
    """
    wait_for_grid(page)
    search_info = get_search_info(page)
    total_pages = get_total_pages(page)
    LOG(f"[i] פרטי חיפוש: {search_info}\n[i] סה\"כ עמודים: {total_pages}")

    # Determine the EXACT number of deals the search returned, so the progress
    # percentage reflects deal progress (not pages). In this grid every page
    # except the last is full, so reading the first-page and last-page row
    # counts yields the exact total without visiting every page in advance.
    rows_first = count_rows_on_page(page)
    total_deals = rows_first
    if total_pages > 1:
        try:
            if goto_page(page, total_pages):
                rows_last = count_rows_on_page(page)
                total_deals = (total_pages - 1) * rows_first + rows_last
            else:
                total_deals = rows_first * total_pages  # safe fallback
            goto_page(page, 1)  # return to the first page before scraping
        except Exception:  # noqa: BLE001
            total_deals = rows_first * total_pages  # safe fallback
    LOG(f"[i] סה\"כ עסקאות שנמצאו: {total_deals}")

    # Progress is reported per DEAL (not per page):
    #   on_progress(current_deal, total_deals, collected, page, total_pages,
    #               page_deal, page_deals)
    if on_progress is not None:
        on_progress(0, total_deals, 0, 0, total_pages, 0, 0)

    all_deals: list[dict[str, str]] = []
    seen: set[str] = set()
    cur = {"page": 0}
    # processed   = deals gone through overall (drives the %).
    # page_deal   = deal index within the current page.
    # page_deals  = total deals on the current page.
    counters = {"processed": 0, "page_deal": 0, "page_deals": 0}

    def emit_progress() -> None:
        if on_progress is not None:
            on_progress(
                counters["processed"],
                total_deals,
                len(all_deals),
                cur["page"],
                total_pages,
                counters["page_deal"],
                counters["page_deals"],
            )

    def handle_deal(r: dict[str, str]) -> None:
        # Count every deal we go through (used for the percentage), even if it
        # turns out to be a duplicate of one we already collected.
        counters["processed"] += 1
        counters["page_deal"] += 1
        # Dedup signature: parcel | date | apartment | declared price
        sig = "|".join([
            r.get("gush_chelka", ""),
            r.get("date", ""),
            r.get("dira", ""),
            r.get("mcir_mozhar", ""),
        ])
        if sig in seen:
            emit_progress()
            return
        seen.add(sig)
        all_deals.append(r)
        if on_deal is not None:
            try:
                on_deal(build_main_row(r), r)
            except Exception:  # noqa: BLE001
                pass
        # Incremental save -- keep what we have even if the run breaks later.
        if output_path:
            try:
                write_excel(all_deals, output_path, search_info)
            except Exception as e:  # noqa: BLE001
                LOG(f"[!] שמירת ביניים נכשלה (ממשיכים): {e}")
        emit_progress()

    for page_num in range(1, total_pages + 1):
        if should_stop is not None and should_stop():
            LOG("[i] עצירה לבקשת המשתמש -- שומר את מה שנאסף עד כה.")
            break

        cur["page"] = page_num
        LOG(f"\n=== עיבוד עמוד {page_num}/{total_pages} ===")
        if not goto_page(page, page_num):
            LOG(f"[!] לא ניתן לנווט לעמוד {page_num} -- מדלג.")
            continue

        # Reset per-page counters now that we are on this page.
        counters["page_deal"] = 0
        counters["page_deals"] = count_rows_on_page(page)
        emit_progress()

        scrape_page_deals(
            page, page_num, should_stop=should_stop, on_deal=handle_deal
        )

        emit_progress()

    # Final exact total now that every page has been visited.
    total_deals = max(total_deals, counters["processed"], len(all_deals))
    emit_progress()
    LOG(f"\n[i] סה\"כ נאספו {len(all_deals)} עסקאות.")
    return all_deals, search_info


def wait_for_user_ready(page: Page) -> bool:
    """
    Wait for explicit user confirmation (Enter) that they finished logging in
    and running the search, then verify the results grid exists. If not, allow
    retrying. Returns True once the grid is found and we can start.
    """
    while True:
        LOG(
            "\n" + "=" * 60 + "\n"
            "  התחבר/י לאתר ובצע/י את החיפוש בדפדפן שנפתח.\n"
            "  המתן/י עד שמוצגת טבלת 'עסקאות שבוצעו' (העמוד הראשון).\n"
            "  כשהכל מוכן -- חזור/י לכאן והקש/י Enter כדי להתחיל.\n"
            "  (להקלדת 'q' + Enter כדי לבטל)\n"
            + "=" * 60
        )
        answer = input()
        if answer.strip().lower() == "q":
            LOG("[i] בוטל ע\"י המשתמש.")
            return False

        if grid_is_ready(page):
            return True


def grid_is_ready(page: Page, timeout: int = 5000) -> bool:
    """Check (without input) that the results grid exists. Also used by the UI."""
    try:
        wait_for_grid(page, timeout=timeout)
        LOG("[i] טבלת התוצאות זוהתה -- אפשר להתחיל.")
        return True
    except Exception:  # noqa: BLE001
        LOG(
            "[!] לא זוהתה טבלת התוצאות (#ContentUsersPage_GridMultiD1).\n"
            "    ודא/י שאת/ה נמצא/ת בעמוד 'עסקאות שבוצעו' עם תוצאות, ונסה/י שוב."
        )
        return False


def open_browser_existing(p):
    """Attach to an existing browser (remote-debugging-port). Returns (browser, page)."""
    LOG(f"[i] מתחבר לדפדפן קיים: {CDP_URL}")
    browser = p.chromium.connect_over_cdp(CDP_URL)
    ctx0 = browser.contexts[0] if browser.contexts else browser.new_context()
    page = None
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if "nadlan" in pg.url.lower() or "svinfo" in pg.url.lower():
                page = pg
                break
        if page:
            break
    if page is None:
        page = ctx0.pages[0] if ctx0.pages else ctx0.new_page()
    page.bring_to_front()
    return browser, page


def _screen_size() -> tuple[int, int]:
    """Return the primary monitor's pixel size (width, height).
    Falls back to a sensible default if the OS query fails."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        # Make the process DPI-aware so we read true pixels, not scaled ones.
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:  # noqa: BLE001
            try:
                user32.SetProcessDPIAware()
            except Exception:  # noqa: BLE001
                pass
        w = user32.GetSystemMetrics(0)  # SM_CXSCREEN
        h = user32.GetSystemMetrics(1)  # SM_CYSCREEN
        if w > 0 and h > 0:
            return int(w), int(h)
    except Exception:  # noqa: BLE001
        pass
    return 1920, 1080


def _maximize_window(page) -> None:
    """Maximize the browser window via the CDP Browser.setWindowBounds API."""
    try:
        session = page.context.new_cdp_session(page)
        win = session.send("Browser.getWindowForTarget")
        session.send("Browser.setWindowBounds", {
            "windowId": win["windowId"],
            "bounds": {"windowState": "maximized"},
        })
    except Exception:  # noqa: BLE001
        pass


def open_browser_own_profile(p):
    """Open a browser with our dedicated profile and navigate to the start page.
    Returns (context, page)."""
    os.makedirs(USER_PROFILE_DIR, exist_ok=True)
    channel_txt = BROWSER_CHANNEL if BROWSER_CHANNEL else "Chromium (Playwright)"
    LOG(f"[i] פותח דפדפן ({channel_txt}) עם פרופיל ייעודי:\n    {USER_PROFILE_DIR}")
    # --start-maximized is not reliably honored by a persistent context, so we
    # also pass an explicit --window-size matching the real screen resolution.
    sw, sh = _screen_size()
    launch_kwargs = dict(
        headless=HEADLESS,
        locale="he-IL",
        # no_viewport lets the browser use its real, maximized window size.
        # A fixed `viewport` would override --start-maximized and leave the
        # page rendered at that size (window opens partial / cut off).
        no_viewport=True,
        ignore_https_errors=True,
        args=[
            "--start-maximized",
            "--window-position=0,0",
            f"--window-size={sw},{sh}",
            "--ignore-certificate-errors",
        ],
    )
    if BROWSER_CHANNEL:
        launch_kwargs["channel"] = BROWSER_CHANNEL
    context = p.chromium.launch_persistent_context(USER_PROFILE_DIR, **launch_kwargs)
    page = context.pages[0] if context.pages else context.new_page()
    # Belt-and-suspenders: maximize via CDP after launch as well.
    _maximize_window(page)
    try:
        page.goto(START_URL, wait_until="domcontentloaded")
    except Exception:  # noqa: BLE001
        pass
    page.bring_to_front()
    return context, page


def main() -> None:
    with sync_playwright() as p:
        context = None
        browser = None

        if BROWSER_MODE == "existing":
            browser, page = open_browser_existing(p)
            if not wait_for_user_ready(page):
                return
        else:
            context, page = open_browser_own_profile(p)
            if not wait_for_user_ready(page):
                context.close()
                return

        human_pause(0.5, 1.0)

        # Unique output file per run, based on the current search + timestamp.
        out_path = make_output_path(os.path.dirname(os.path.abspath(__file__)),
                                    get_search_info(page))
        deals, search_info = run(page, output_path=out_path)

        if not deals:
            LOG("[!] לא נאספו עסקאות. בדוק שאתה נמצא בעמוד התוצאות.")
            if context is not None:
                context.close()
            return

        write_excel(deals, out_path, search_info)

        LOG(
            f"\n[i] נשמרו {len(deals)} עסקאות לקובץ {os.path.basename(out_path)} "
            f"(2 גיליונות: 'עסקאות שבוצעו' ו-'מורחב')."
        )

        # Close the browser (only in own_profile mode that we opened)
        if context is not None:
            LOG("[i] הקש/י Enter לסגירת הדפדפן...")
            try:
                input()
            except Exception:  # noqa: BLE001
                pass
            context.close()


if __name__ == "__main__":
    main()
