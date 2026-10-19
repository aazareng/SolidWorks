"""
DRAWING_LINKS.py - overlay clickable hotspots onto a SolidWorks drawing PDF.

Reads the geometry/BOM out of a DXF export of a drawing sheet, works out the
DXF -> PDF coordinate transform, and stamps clickable regions over every
balloon and every BOM row of the *original* PDF.  Nothing is redrawn - the
PDF you get back is your PDF plus annotations.

Clicking a balloon or a BOM row SHOWS the part - what it is, which sheet calls
for it, how many the BOM asks for - with the quantity ready to adjust and an
Add button beside it.  "Request Quote" mails the cart, grouped by sheet,
straight to the parts desk - as a #POLYPACK-PARTS-RFQ v1 block that reads as
a plain list and imports into PolyApp's CO > PARTS QUOTE as-is (see quoteBody
in SECTION_JS for the grammar).  The page wears Polypack's own chrome.

MULTI-SHEET: give one DXF per sheet, in sheet order, and a PDF with one page
per sheet - DXF n pairs with page n.  A DXF that carries several paperspace
layouts is read as several sheets too.  The page gets a sheet switcher, each
sheet keeps its own BOM, and the cart spans the lot: one line per part number
however many sheets call for it, so a part on two sheets is quoted once.

Outputs.  Only the files people open go to --out (default: next to the source
PDF); the working file goes to --work (default: %TEMP%\\DRAWING_LINKS), so a
publish folder holds nothing but the deliverable:
    <base>.html          the quote page - drawing, hotspots, cart, RFQ
    <base>_section.json  only with --emit section: the same thing as a SECTION
                         for the Technical Manual to absorb (see below).
                         Written HIDDEN - it has to sit beside the page for the
                         manual to find it, but nobody opens it, so the publish
                         folder shows just the .html.  config hide_section_json
    <base>_links.json    (working) the hotspot map, for debugging - the page
                         does not read it
    <base>_linked.pdf    (working) only with --linked-pdf: the original PDF
                         plus hotspots that open the page at that part.  It
                         keeps the _linked suffix because --work already holds
                         the <base>.pdf it is built from.

THE TECHNICAL MANUAL.  The page's chrome - the red, the header, the contact
drawer, the left rail, the toast - is not written here any more: it lives in
POLYPACK_SHELL.py, shared with the manual and the Word converter, so all three
look like one product and the red is changed in one place.  What is written
here is the part that is actually about a drawing.

That is also what lets the manual stop framing this page.  `--emit section`
writes it as a SECTION instead - its CSS, markup, behaviour and data, all
namespaced under one id - and any number of sections merge into one file with
no <iframe>, one shared rail and one shared cart:

    python DRAWING_LINKS.py A9987.DXF --emit section
    python POLYPACK_SHELL.py A9987_section.json safety_section.json \\
           --out TM_4630.html --title 4630

`--emit page` (the default) is unchanged, which is what the SolidWorks macro
and the parts desk keep getting.

Usage (this is what the SolidWorks VBA macro calls):
    python DRAWING_LINKS.py "C:\\...\\A9987-0000-M.DXF" --open
    python DRAWING_LINKS.py "...\\A9987_S1.DXF" "...\\A9987_S2.DXF" --open

    --pdf PATH      source PDF (default: same folder/basename as the first
                    DXF, with any _S1 suffix removed)
    --out DIR       where the HTML and linked PDF go (default: beside the PDF)
    --work DIR      where the working files go (default: %TEMP%\\DRAWING_LINKS)
    --sheet-names   SolidWorks' own sheet names, in DXF order - these become
                    the tab labels on the page
    --zones PATH    the zones manifest the SolidWorks macro writes (zones.txt).
                    Two things come of it.  On the MAIN sheet every section
                    sheet gets a CLICK ZONE where its assembly sits in the
                    machine - hover shows the sheet, click opens it.  On every
                    section sheet each top-level component of that sheet's
                    assembly gets a SILHOUETTE - hover lights its own linework,
                    click opens its part card when the BOM has it.  The
                    silhouettes come out of the DXFs on their own; the manifest
                    names them and places the MAIN zones.  Pillow is needed.
                    A 2026-10-09 macro also lists every PART under the
                    section's reference assembly (REFPART lines): those make
                    the ALL PARTS panel under the parts list - shut until
                    opened, a preview on every row, and while it is open a
                    row lights its part on the sheet and the sheet lights the
                    row.  Without the flag the page is exactly what it was.
    --thumbs PATH   the BOM manifest the SolidWorks macro writes (thumbs.txt,
                    with the bitmaps beside it).  Carries two things per row:
                    the model's saved preview, shown in the part card, and the
                    row's SHOP cost, shown as a price on the card and as a
                    column in the parts list after price_multiplier is applied.
                    Pillow is needed for the pictures and not for the prices.
                    Without the flag the page is exactly what it was.
    --open          open the generated HTML when finished
    --viewer MODE   how the drawing is shown in the HTML:
                      auto   rasterise if PyMuPDF is installed, else embed
                      raster render the page to an image - exact alignment
                      pdf    embed the PDF itself - browser viewers pad the
                             page, so hotspots need the Align sliders
    --dpi N         raster resolution (default 200)
    --image-format  auto (WebP, ~1/6 the size of PNG) | webp | jpeg | png
    --quality N     webp/jpeg quality, 1-100 (default 85)
    --page N        which PDF page a lone DXF belongs to (default 0)
    --embed-pdf     inline the PDF into the HTML so the page stands alone;
                    by default it just references the PDF in --work, which
                    only matters when PyMuPDF is missing and the page cannot
                    rasterise the sheets
    --no-boxes      do not draw the visible hotspot outlines on the PDF
    --emit WHAT     page (default) | section | both - see THE TECHNICAL
                    MANUAL above
    --section-label the manual's rail label for the section it emits
    --config PATH   JSON config (see DEFAULT_CONFIG below)
    --quiet         only print errors

Run with no arguments for the drag-and-drop window.

Requires: ezdxf, pypdf.  Strongly recommended: pymupdf (exact HTML
alignment).  Optional: tkinterdnd2 (drag and drop).
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import shutil
import pathlib
import re
import sys
import tempfile
import webbrowser
from dataclasses import dataclass, field

import ezdxf

# The page chrome, shared with the Technical Manual and the Word converter so
# all three look like one product.  It used to be imported from POLYPACK_SHELL.py
# beside this file; this copy carries it EMBEDDED (see the POLYPACK_SHELL section
# near the end) so the one file in the PDM vault is the whole tool.  `shell` is
# this module, so the three call sites below read as they always did.
shell = sys.modules[__name__]

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    DecodedStreamObject,
    DictionaryObject,
    FloatObject,
    NameObject,
    NumberObject,
    TextStringObject,
)

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

DEFAULT_CONFIG = {
    # What a hotspot in the *PDF* does:
    #   "html"    open the generated quote page with that item added to the
    #             cart.  Works in any viewer that follows links.
    #   "message" pop a message box listing the part.  Acrobat only - the
    #             browsers' built-in PDF viewers do not run PDF JavaScript.
    #   "url"     open link_template.
    "pdf_action": "html",

    # What a hotspot in the *HTML* does: "cart" or "url".
    "html_action": "cart",

    # Used by the "url" actions.  Placeholders are the BOM column names
    # upper-cased with non-alphanumerics stripped, plus {ITEM} and {DRAWING}.
    # e.g. "https://polyapp/parts/{PART}"
    "link_template": "",

    # Where the PDF's links point when pdf_action == "html".  Empty means the
    # generated file itself, as an absolute file:// URL - which breaks if the
    # pair is moved.  Set this to the folder the pages are served from, e.g.
    # "https://polyapp/drawings/", once they live somewhere.
    "html_link_base": "",

    # --- Request Quote -------------------------------------------------
    # Where an RFQ goes.  The page remembers anything you type over it.
    "quote_email": "parts@polypack.com",
    # {SERIAL} is the machine serial read out of the drawing number - the four
    # characters after the leading letter, so A9987-0000-M -> 9987.
    "quote_subject": "{SERIAL}-PARTS RFQ",
    # Prepended to the emailed parts list - above the #POLYPACK-PARTS-RFQ
    # block, so it is read by people and skipped by the importer.
    "quote_intro": "Please quote the following parts:",

    # --- Masthead -------------------------------------------------------
    # The first entry's phone is the one kept in the header; the rest live
    # behind the Contact button.  "tel" is worked out from "phone" if absent.
    "contact": {
        "company": "POLYPACK",
        "tagline": "WE GOT IT COVERED",
        "website": "https://www.polypack.com/",
        # One block per PLACE you can reach, each carrying its own numbers,
        # address and mailbox.  Headquarters comes first because the header
        # keeps the first block's first number - the switchboard.
        "lines": [
            {"label": "Main Headquarters",
             "address": "Polypack, Inc.\n3301 Gateway Centre Blvd\n"
                        "Pinellas Park, FL 33782",
             "phones": ["(727) 578-5000", "(727) 578-1300"],
             "email": "info@polypack.com"},
            {"label": "Service", "phone": "(727) 578-5000 x221",
             "email": "service@polypack.com"},
            # masthead: this page is a parts RFQ, so Parts is the number in the
            # header - regardless of where it sits in the drawer above.
            {"label": "Parts", "phone": "(727) 578-5000 x245",
             "email": "parts@polypack.com", "masthead": True},
        ],
    },

    # Which BOM column holds the part number (matched case-insensitively,
    # first hit wins).  Drives {PART}, the cart lines and the quote email.
    "part_columns": ["PART #", "PART NUMBER", "PART", "PARTNO", "PART NO"],

    # Where a row's part number comes from.  "vault" = the PDM vault's own
    # value for the row's file - its name in the vault, which is what PDM's
    # Part Number column holds (the card has no part-number variable) - read
    # off thumbs.txt column 8; the drawing's Part # cell (SolidWorks' custom
    # property) only for a row the macro found no file for.  "table" = the
    # cell, as before 2026-10-13.  The -M drawing's BOM keeps its Part #
    # column HIDDEN (the macro still reads it to find each row's file); a
    # sheet whose table shows none gets `part_column_label` added after the
    # description, so both panels keep the column.
    "part_source": "vault",
    "part_column_label": "PART #",

    # Which BOM column holds the quantity - used as the starting cart qty.
    "qty_columns": ["QTY.", "QTY", "QUANTITY"],

    # Which BOM column holds the description, used in the quote email.
    "description_columns": ["DESCRIPTION", "DESC", "TITLE"],

    # Which BOM column holds the configuration / vendor reference.  Often the
    # only thing that pins down which variant to buy, so it rides along on the
    # cart line and into the RFQ - but only when the drawing actually uses it.
    "config_columns": ["CONFIG", "CONFIG.", "CONFIGURATION"],

    # How much of the window the drawing fills.  The sheet is always sized to
    # fit the window's height; this trims it further.  1.0 = as large as fits.
    "page_zoom": 0.95,

    # How each sheet is encoded into the page.  "auto" = WebP (roughly a fifth
    # the size of PNG at the same resolution, and indistinguishable on a
    # drawing); "jpeg" for a viewer too old for WebP; "png" for lossless.
    # Quality applies to webp/jpeg.
    "image_format": "auto",
    "image_quality": 85,

    # The model previews shown in the part card (see --thumbs).  The long edge
    # in pixels, and the WebP quality.  The card is 310px wide, so past about
    # 300 the extra pixels only cost page weight.
    "thumb_px": 260,
    "thumb_quality": 82,

    # --- Prices ---------------------------------------------------------
    # What arrives in the manifest is the SHOP cost of the row.  This is what
    # it is multiplied by before anybody sees it, so one publish serves both
    # audiences and changing the margin never means re-exporting a drawing:
    # 1.0 shows cost, 1.6 shows cost plus 60%.
    #
    # 0 (or None) suppresses prices ENTIRELY - no card line, no column, and the
    # figures are dropped rather than hidden, so a page built for a customer
    # has no cost in its source for anyone to go looking for.
    "price_multiplier": 1.0,
    "price_label": "Est. Price",
    # The two panels beside the sheet.  The parts list IS the -M drawing's
    # curated BOM - the parts worth keeping - so it is called what it is; the
    # second lists every part under the sheet's reference assembly, and only
    # shows when the macro wrote REFPART lines into zones.txt.
    "parts_label": "Critical parts",
    "allparts_label": "All parts",
    # Where a reference part has no outline of its own (the macro could not
    # layer it, or the export dropped the layer), its instance box can stand
    # in - one box per instance, never a union.  A box bigger than this share
    # of the page is a frame or a guard, not a part worth a rectangle, and is
    # left off; 0 or allparts_boxes False means no boxes at all.
    "allparts_boxes": True,
    "allparts_box_max": 0.06,
    "price_currency": "$",
    # A Price column in the parts list, beside the card line.  This is the one
    # that survives a missing balloon: the BOM row carries its own price.
    "price_in_bom": True,
    # Show the basis the macro costed on ("LENGTH", "SHEET", "STOCK#", "ASSY"...)
    # as the column's tooltip.  Useful in-house, noise to a customer.
    "price_show_basis": True,
    # The PolyApp data API, read once per publish for the MATERIALS table so a
    # length material is priced on the length the row uses, at the moment the
    # page is made.  The macro costs a stock row with no purchase Length "each":
    # 14002, a #40 chain bought at $2.70 a FOOT, priced every chain at one foot
    # whatever its length (1208.5 mm read $8.10).  "" = off - macro prices only.
    "polyapp_api": "http://polyflapp02:8000",

    # The PDM version of the file each BOM row resolved to, as the macro read
    # it off the vault (thumbs.txt column 9) - the version on disk when the
    # drawing was published, which is the one the sheet actually shows.  A
    # column in the parts list and a line on the part card.  "" / None for the
    # label turns both off; a drawing whose manifest carries no versions shows
    # neither regardless.
    "pdm_version_label": "PDM Ver",

    # Cut-to-length items (chain, extrusion, cable) carry a MEASUREMENT in the
    # quantity column rather than a piece count.  A quantity with decimals, or
    # one at least `min`, is read that way and shown with the conversion beside
    # it.  factor = source units per target unit; 304.8 mm per foot.
    "length_units": {"min": 100, "from": "mm", "to": "ft", "factor": 304.8},

    # Hotspot padding, in PDF points, added around balloons and BOM rows.
    "balloon_padding_pt": 1.0,
    "row_padding_pt": 0.0,

    # Visible outline drawn on the PDF.  RGB 0-1, and stroke width in points.
    "box_rgb_balloon": [0.90, 0.45, 0.00],
    "box_rgb_row": [0.00, 0.45, 0.85],
    "box_width_pt": 0.6,

    # Manual calibration, applied after the automatic fit.  Only touch these
    # if a drawing's hotspots land slightly off.
    "offset_pt": [0.0, 0.0],
    "scale_override": None,

    # The section manifest has to sit beside the page - that is where the
    # Technical Manual looks for it - but it is not a file anyone opens, so it
    # is written hidden and the publish folder shows only the .html.  os.listdir
    # still returns it, which is how the API enumerates manifests, so hiding it
    # costs the manual nothing.  Set false to see it in Explorer.
    "hide_section_json": True,
}

# Blocks SolidWorks writes for balloons and for tables.
BALLOON_BLOCK_PREFIX = "SW_NOTE"
TABLE_BLOCK_PREFIX = "SW_TABLEANNOTATION"

# Grid lines closer together than this (DXF units) are treated as one line.
GRID_TOL = 1e-3


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class Balloon:
    item: str
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class TableRow:
    item: str
    values: dict[str, str]
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class Table:
    columns: list[str]
    rows: list[TableRow] = field(default_factory=list)
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0


@dataclass
class Sheet:
    """One drawing sheet: the balloons and BOM of a single DXF space.

    `extents` holds every candidate for "where the sheet corners are", best
    guess first.  Which one is right is settled against the PDF page it pairs
    with - see _choose_extents.
    """
    name: str                      # layout name, or the file's own name
    source: str                    # the DXF it came from
    balloons: list[Balloon]
    table: Table | None
    extents: list[tuple[str, tuple[float, float], tuple[float, float]]] = \
        field(default_factory=list)
    ext_min: tuple[float, float] = (0.0, 0.0)
    ext_max: tuple[float, float] = (0.0, 0.0)
    ext_source: str = ""
    # The drawing view's geometry, one entry per top-level component of the
    # model the view shows, as flattened polylines in DXF units - see
    # _view_blocks.  Empty unless the export wrote view blocks.
    blocks: dict[str, list[list[tuple[float, float]]]] = field(default_factory=dict)
    # The reference assembly's linework, from the macro's REF-only export
    # (<stem>_REF.dxf beside the sheet's DXF) - see _attach_ref_blocks.  The
    # sheet's own DXF was written with the REF hidden, so these are the only
    # lines of it there are.  Same shape as `blocks`, keyed the same way.
    ref_blocks: dict[str, list[list[tuple[float, float]]]] = field(default_factory=dict)
    # The same export's lines grouped by drawing LAYER, for the layers the
    # macro gave the reference parts (AP<sheet>_<n>) - one entry per part
    # instance, whatever block they sit in.  See _layer_lines.
    ref_layers: dict[str, list[list[tuple[float, float]]]] = field(default_factory=dict)


# --------------------------------------------------------------------------
# DXF parsing
# --------------------------------------------------------------------------

def _join_fragments(left: str, right: str) -> str:
    """Glue two pieces of one cell back together.

    A long BOM value gets broken across lines by SolidWorks, and it breaks
    wherever it runs out of column - including the middle of a number, so a
    chain length of 1407.35 arrives as "1407.3" then "5".  Joining those with a
    space would read as two numbers and a quantity of 14073.5 would go out on
    an RFQ.  Digits either side of the break mean it was one number.
    """
    left, right = left.rstrip(), right.lstrip()
    if not left:
        return right
    if not right:
        return left
    if re.search(r"[\d.]$", left) and re.match(r"[\d]", right):
        return left + right
    return left + " " + right


def _unwrap(text: str) -> str:
    """Collapse the line breaks inside one cell's own text."""
    parts = [p for p in re.split(r"[\r\n]+", text) if p.strip()]
    if not parts:
        return ""
    out = parts[0].strip()
    for part in parts[1:]:
        out = _join_fragments(out, part)
    return out


def _plain_text(entity) -> str:
    """Text of a TEXT/MTEXT/ATTRIB entity with MTEXT formatting codes removed."""
    etype = entity.dxftype()
    if etype == "MTEXT":
        try:
            return _unwrap(entity.plain_text())
        except Exception:
            return _unwrap(str(entity.text))
    if etype in ("TEXT", "ATTRIB"):
        return _unwrap(str(entity.dxf.text))
    return ""


def _dedupe_sorted(values: list[float]) -> list[float]:
    """Sort and collapse near-identical coordinates (grid lines drawn twice)."""
    out: list[float] = []
    for v in sorted(values):
        if not out or abs(v - out[-1]) > GRID_TOL:
            out.append(v)
    return out


def _extract_balloon(insert) -> tuple[Balloon, bool] | None:
    """Pull the item number and bounding box out of one SW_NOTE block.

    SolidWorks writes plain sheet notes into SW_NOTE blocks too, so the second
    element of the result says whether this block actually looks like a
    balloon - i.e. it carries a shape around the text, not just text.
    """
    try:
        subs = list(insert.virtual_entities())
    except Exception:
        return None

    item = ""
    circles = []
    has_shape = False
    for sub in subs:
        stype = sub.dxftype()
        if stype == "CIRCLE":
            circles.append(sub)
            has_shape = True
        elif stype in ("TEXT", "MTEXT", "ATTRIB"):
            if not item:
                item = _plain_text(sub)
        else:
            # Square / hexagonal / triangular balloon outlines.
            has_shape = True

    if not item:
        return None
    # "3 (2)" or "3,4" -> "3": the leading token is the BOM item number.
    item = item.split()[0].strip().rstrip(",")

    if circles:
        # Balloons are usually a circle + a number; the circle is the hotspot.
        x0 = min(c.dxf.center.x - c.dxf.radius for c in circles)
        x1 = max(c.dxf.center.x + c.dxf.radius for c in circles)
        y0 = min(c.dxf.center.y - c.dxf.radius for c in circles)
        y1 = max(c.dxf.center.y + c.dxf.radius for c in circles)
        return Balloon(item, x0, y0, x1, y1), True

    # No circle - fall back to the block's own extents.
    try:
        from ezdxf import bbox

        extents = bbox.extents(subs, fast=True)
        if extents.has_data:
            return Balloon(item, extents.extmin.x, extents.extmin.y,
                           extents.extmax.x, extents.extmax.y), has_shape
    except Exception:
        pass
    return None


def _extract_table(insert) -> Table | None:
    """Rebuild a SW_TABLEANNOTATION block from its grid lines and cell text.

    The grid lines give exact cell rectangles, which is what we need for
    row hotspots - inferring rows from text positions alone is fragile.
    """
    try:
        subs = list(insert.virtual_entities())
    except Exception:
        return None

    h_lines: list[tuple[float, float, float]] = []  # (y, xmin, xmax)
    v_lines: list[tuple[float, float, float]] = []  # (x, ymin, ymax)
    cells: list[tuple[float, float, str]] = []      # (x, y, text)

    for sub in subs:
        stype = sub.dxftype()
        if stype == "LINE":
            a, b = sub.dxf.start, sub.dxf.end
            if abs(a.y - b.y) <= GRID_TOL:
                h_lines.append((a.y, min(a.x, b.x), max(a.x, b.x)))
            elif abs(a.x - b.x) <= GRID_TOL:
                v_lines.append((a.x, min(a.y, b.y), max(a.y, b.y)))
        elif stype in ("TEXT", "MTEXT"):
            text = _plain_text(sub)
            if text:
                pos = sub.dxf.insert
                cells.append((pos.x, pos.y, text))

    if len(h_lines) < 3 or len(v_lines) < 2 or not cells:
        return None

    bands_y = _dedupe_sorted([h[0] for h in h_lines])          # bottom -> top
    table_x0 = min(v[0] for v in v_lines)
    table_x1 = max(v[0] for v in v_lines)

    def columns_crossing(y_lo: float, y_hi: float) -> list[float]:
        """X of every vertical rule that spans this horizontal band."""
        mid = (y_lo + y_hi) / 2.0
        xs = [v[0] for v in v_lines if v[1] - GRID_TOL <= mid <= v[2] + GRID_TOL]
        return _dedupe_sorted(xs)

    def cells_in(y_lo: float, y_hi: float) -> list[tuple[float, float, str]]:
        return [c for c in cells if y_lo - GRID_TOL <= c[1] <= y_hi + GRID_TOL]

    # Walk bands from the top down: title band(s) first, then the header, then
    # the data rows.  A band with no interior vertical rule is a title.
    bands = list(zip(bands_y[:-1], bands_y[1:]))  # (y_lo, y_hi), bottom -> top
    bands.reverse()                                # top -> bottom

    columns: list[str] = []
    column_x: list[float] = []
    rows: list[TableRow] = []

    for y_lo, y_hi in bands:
        xs = columns_crossing(y_lo, y_hi)
        band_cells = cells_in(y_lo, y_hi)
        if not band_cells:
            continue
        if len(xs) <= 2 and not columns:
            continue  # spanning title row, e.g. "PARTS LIST: TUNNEL FAN"

        if not columns:
            # First multi-column band is the header.  A header can be split
            # over two text lines ("ITEM" above "#"), so join per column.
            column_x = xs[:-1]
            parts: list[str] = ["" for _ in column_x]
            for x, y, text in sorted(band_cells, key=lambda c: (-c[1], c[0])):
                idx = _column_index(x, xs)
                parts[idx] = _join_fragments(parts[idx], text)
            columns = [p.strip() or f"COL{i + 1}" for i, p in enumerate(parts)]
            continue

        values = {name: "" for name in columns}
        for x, y, text in sorted(band_cells, key=lambda c: (-c[1], c[0])):
            idx = _column_index(x, xs)
            if idx >= len(columns):
                idx = len(columns) - 1
            name = columns[idx]
            # stacked pieces of one cell - same wrap problem as within an MTEXT
            values[name] = _join_fragments(values[name], text)

        item = values[columns[0]].strip()
        if not item:
            continue
        rows.append(TableRow(item, values, table_x0, y_lo, table_x1, y_hi))

    if not columns or not rows:
        return None

    table = Table(columns=columns, rows=rows)
    table.x0, table.x1 = table_x0, table_x1
    table.y0 = min(r.y0 for r in rows)
    table.y1 = max(bands_y)
    return table


def _column_index(x: float, boundaries: list[float]) -> int:
    """Index of the cell containing x, given that band's vertical rules."""
    for i in range(len(boundaries) - 1):
        if x < boundaries[i + 1]:
            return max(i, 0)
    return max(len(boundaries) - 2, 0)


def _parse_space(space, name: str, source: str) -> Sheet | None:
    """Pull the balloons and the BOM out of one DXF space (modelspace or one
    paperspace layout).  Returns None if there is nothing clickable in it."""
    candidates: list[tuple[Balloon, bool]] = []
    tables: list[Table] = []

    for insert in space.query("INSERT"):
        block = insert.dxf.name or ""
        if block.startswith(BALLOON_BLOCK_PREFIX):
            candidate = _extract_balloon(insert)
            if candidate:
                candidates.append(candidate)
        elif block.startswith(TABLE_BLOCK_PREFIX):
            table = _extract_table(insert)
            if table:
                tables.append(table)

    table = _pick_bom_table(tables, {b.item for b, _ in candidates})
    bom_items = {row.item for row in table.rows} if table else set()
    # Keep a SW_NOTE block only if it is drawn like a balloon, or its text is
    # an item number the BOM knows about.  That drops sheet notes such as
    # "NOTES:" and the address block in the title bar.
    balloons = [b for b, has_shape in candidates
                if has_shape or b.item in bom_items]

    if not balloons and table is None:
        return None
    return Sheet(name=name, source=source, balloons=balloons, table=table)


VIEW_BLOCK_PREFIX = "DRAWING_VIEW"


def _block_key(name: str) -> str:
    """The component a view block stands for, as far as the export lets on.

    SolidWorks writes one block per top-level component of the view's model,
    named DRAWING_VIEW<n>_<instance name>, the name with every non-alphanumeric
    as "_", "<name>_<assembly>" when the name is short, the whole clipped to
    12 characters and deduped by a counter that overwrites the tail.  The first
    ten characters are what reliably survive, so that is the key - and the
    reason two components sharing a ten-character prefix fold into one
    silhouette.  The macro's manifest gives the full names back."""
    parts = name.split("_", 2)
    tail = parts[2] if len(parts) == 3 else name
    return tail[:10].lower()


def _view_blocks(space, tol: float) -> dict[str, list[list[tuple[float, float]]]]:
    """Every view block in a space, flattened to polylines, keyed by the
    block's own name - one block per instance, as SolidWorks wrote them.
    build_silhouettes folds them into parts (see _block_key for the key, and
    the instance boxes for telling instances apart).  `tol` is the flattening
    distance in DXF units."""
    from ezdxf import path as ezpath

    blocks: dict[str, list[list[tuple[float, float]]]] = {}
    for insert in space.query("INSERT"):
        name = insert.dxf.name or ""
        if not name.startswith(VIEW_BLOCK_PREFIX):
            continue
        out = blocks.setdefault(name, [])
        try:
            entities = list(insert.virtual_entities())
        except Exception:
            continue
        for entity in entities:
            try:
                pts = [(v.x, v.y) for v in ezpath.make_path(entity).flattening(tol)]
            except Exception:
                continue
            if len(pts) >= 2:
                out.append(pts)
    return {k: v for k, v in blocks.items() if v}


PART_LAYER_PREFIX = "AP"


def _layer_lines(space, tol: float, prefix: str = PART_LAYER_PREFIX
                 ) -> dict[str, list[list[tuple[float, float]]]]:
    """Every line in the space grouped by the LAYER it is on, for the layers
    named `prefix`<digits>_<digits> - the ones dl_PartLayersOn in the macro
    gives the reference parts.  Read from WHEREVER SolidWorks put them: a
    line drawn straight into the space, one inside any block however it is
    named, one inside a block inside a block.  (The first version looked only
    inside DRAWING_VIEW blocks, and an export whose components wear their own
    line fonts need not keep that shape.)  A line on layer "0" inside a block
    takes the block's own layer, as the DXF rules say."""
    from ezdxf import path as ezpath

    out: dict[str, list[list[tuple[float, float]]]] = {}
    pat = re.compile(r"^%s\d+_\d+$" % re.escape(prefix), re.IGNORECASE)

    def take(entity, host: str, depth: int) -> None:
        if entity.dxftype() == "INSERT":
            if depth > 6:
                return
            own = (entity.dxf.layer or "").strip()
            try:
                subs = list(entity.virtual_entities())
            except Exception:
                return
            for sub in subs:
                take(sub, own if own and own != "0" else host, depth + 1)
            return
        layer = (entity.dxf.layer or "").strip()
        if layer == "0" or not layer:
            layer = host
        if not pat.match(layer):
            return
        try:
            pts = [(v.x, v.y) for v in ezpath.make_path(entity).flattening(tol)]
        except Exception:
            return
        if len(pts) >= 2:
            out.setdefault(layer.upper(), []).append(pts)

    for entity in space:
        take(entity, "", 0)
    return out


def _attach_view_blocks(sheet: "Sheet", space) -> None:
    """Read the view blocks for a sheet once its extents are known (the
    flattening tolerance is a fraction of the sheet)."""
    size = 0.0
    for _, lo, hi in sheet.extents:
        size = max(size, hi[0] - lo[0], hi[1] - lo[1])
    if size <= 0:
        return
    try:
        sheet.blocks = _view_blocks(space, size / 4000.0)
    except Exception:
        sheet.blocks = {}


REF_LAYER_SUFFIX = "_REF.dxf"


def _attach_ref_blocks(sheet: "Sheet", dxf_path: str) -> str:
    """The REF layer the macro writes beside a sheet's DXF, when it wrote
    one: the same sheet exported with everything BUT the reference assembly
    hidden.  Its view blocks become `sheet.ref_blocks`; the sheet's own export
    was made with the REF hidden, so every part's silhouette is whole and the
    REF's outline lives only here.  Returns the path read, "" when there is
    none - an older macro, or a sheet whose assembly has no REF folder."""
    stem, ext = os.path.splitext(dxf_path)
    if ext.lower() != ".dxf":
        return ""
    ref_path = stem + REF_LAYER_SUFFIX
    if not os.path.isfile(ref_path):
        return ""
    size = 0.0
    for _, lo, hi in sheet.extents:
        size = max(size, hi[0] - lo[0], hi[1] - lo[1])
    if size <= 0:
        return ""
    try:
        doc = ezdxf.readfile(ref_path)
        sheet.ref_blocks = _view_blocks(doc.modelspace(), size / 4000.0)
        sheet.ref_layers = _layer_lines(doc.modelspace(), size / 4000.0)
    except Exception:
        sheet.ref_blocks = {}
        sheet.ref_layers = {}
        return ""
    return ref_path if sheet.ref_blocks else ""


def _space_extents(space) -> tuple[tuple[float, float], tuple[float, float]] | None:
    """Bounding box of everything drawn in a space."""
    from ezdxf import bbox

    try:
        found = bbox.extents(space, fast=True)
    except Exception:
        return None
    if not found.has_data:
        return None
    return ((found.extmin.x, found.extmin.y), (found.extmax.x, found.extmax.y))


def parse_dxf(dxf_path: str) -> list[Sheet]:
    """Every drawing sheet in one DXF file.

    A single-sheet SolidWorks export puts the sheet in modelspace; a
    multi-sheet export can also carry one paperspace layout per sheet.  Both
    are read, so it does not matter which way the export was set up.
    """
    doc = ezdxf.readfile(dxf_path)
    label = os.path.splitext(os.path.basename(dxf_path))[0]
    sheets: list[Sheet] = []

    msp = doc.modelspace()
    sheet = _parse_space(msp, label, dxf_path)
    if sheet is not None:
        # The header extents describe modelspace, and for a SolidWorks sheet
        # export they are exactly the sheet corners - the drawn bounds are not,
        # because the sheet format's border sits inside the sheet edge.
        head_min = doc.header.get("$EXTMIN", None)
        head_max = doc.header.get("$EXTMAX", None)
        if (head_min is not None and head_max is not None
                and float(head_max[0]) - float(head_min[0]) > 0
                and float(head_max[1]) - float(head_min[1]) > 0):
            sheet.extents.append((
                "header extents",
                (float(head_min[0]), float(head_min[1])),
                (float(head_max[0]), float(head_max[1]))))
        drawn = _space_extents(msp)
        if drawn is not None:
            sheet.extents.append(("drawn bounds", drawn[0], drawn[1]))
        if sheet.extents:
            _attach_view_blocks(sheet, msp)
            _attach_ref_blocks(sheet, dxf_path)
            sheets.append(sheet)

    for name in doc.layouts.names_in_taborder():
        if name.lower() == "model":
            continue
        layout = doc.layout(name)
        sheet = _parse_space(layout, name, dxf_path)
        if sheet is None:
            continue
        paper = _layout_paper(layout)
        if paper is not None:
            sheet.extents.append(("layout paper size", paper[0], paper[1]))
        drawn = _space_extents(layout)
        if drawn is not None:
            sheet.extents.append(("drawn bounds", drawn[0], drawn[1]))
        if sheet.extents:
            _attach_view_blocks(sheet, layout)
            sheets.append(sheet)

    # A file with nothing to link on it is still a sheet.  The macro exports one
    # DXF per sheet and one PDF page per sheet, and the two are paired by
    # POSITION - so a cover sheet with no balloons and no BOM must still take
    # its place in the list, or every sheet after it lands on the wrong page.
    # It comes back with its extents and nothing else; the page it becomes has
    # no hotspots, which is exactly what the drawing has.
    if not sheets:
        blank = Sheet(name=label, source=dxf_path, balloons=[], table=None)
        head_min = doc.header.get("$EXTMIN", None)
        head_max = doc.header.get("$EXTMAX", None)
        if (head_min is not None and head_max is not None
                and float(head_max[0]) - float(head_min[0]) > 0
                and float(head_max[1]) - float(head_min[1]) > 0):
            blank.extents.append((
                "header extents",
                (float(head_min[0]), float(head_min[1])),
                (float(head_max[0]), float(head_max[1]))))
        drawn = _space_extents(msp)
        if drawn is not None:
            blank.extents.append(("drawn bounds", drawn[0], drawn[1]))
        if blank.extents:
            _attach_view_blocks(blank, msp)
            _attach_ref_blocks(blank, dxf_path)
            sheets.append(blank)

    return sheets


def _layout_paper(layout):
    """The sheet rectangle a paperspace layout declares, if it declares one."""
    try:
        lo, hi = layout.dxf.get("limmin", None), layout.dxf.get("limmax", None)
        if lo is not None and hi is not None and hi[0] - lo[0] > 0 and hi[1] - lo[1] > 0:
            return (float(lo[0]), float(lo[1])), (float(hi[0]), float(hi[1]))
        width = float(layout.dxf.get("paper_width", 0) or 0)
        height = float(layout.dxf.get("paper_height", 0) or 0)
        if width > 0 and height > 0:
            return (0.0, 0.0), (width, height)
    except Exception:
        pass
    return None


def _choose_extents(sheet: Sheet, page_ratio: float, label: str, log) -> None:
    """Settle which candidate really is the sheet.

    The DXF and the PDF page are the same sheet, so the candidate whose
    proportions match the page is the one to trust.  That matters most for
    paperspace layouts, where the only thing on offer may be the drawn bounds -
    and a note hanging past the border would otherwise drag the whole overlay
    out of place.
    """
    scored = []
    for name, ext_min, ext_max in sheet.extents:
        width = ext_max[0] - ext_min[0]
        height = ext_max[1] - ext_min[1]
        if width <= 0 or height <= 0:
            continue
        scored.append((abs(width / height - page_ratio) / page_ratio,
                       name, ext_min, ext_max))
    if not scored:
        raise ValueError(f"{label}: the DXF has no usable extents.")

    scored.sort(key=lambda s: s[0])
    error, name, ext_min, ext_max = scored[0]
    sheet.ext_source, sheet.ext_min, sheet.ext_max = name, ext_min, ext_max

    if error > 0.005:
        alternatives = ", ".join(f"{n} {e * 100:.2f}%" for e, n, _, _ in scored)
        log(f"  WARNING: {label} - no DXF extent matches the page proportions "
            f"({alternatives}); hotspots will be offset. Export this sheet's "
            f"DXF on its own so the header extents describe it.")


def _pick_bom_table(tables: list[Table], balloon_items: set[str]) -> Table | None:
    """A sheet can carry several tables (BOM, revision, hole).  The BOM is the
    one whose item column lines up with the balloons; failing that, the
    biggest one."""
    if not tables:
        return None
    if balloon_items:
        scored = [(sum(1 for r in t.rows if r.item in balloon_items), t)
                  for t in tables]
        scored.sort(key=lambda s: s[0], reverse=True)
        if scored[0][0] > 0:
            return scored[0][1]
    return max(tables, key=lambda t: len(t.rows))


# --------------------------------------------------------------------------
# DXF -> PDF transform
# --------------------------------------------------------------------------

class Transform:
    """Maps DXF sheet coordinates onto PDF user-space points.

    A SolidWorks DXF export of a drawing sheet covers exactly the same sheet
    as the PDF export, so the sheet extents map onto the page box.  The two
    axis scales are compared as a sanity check - they should agree to well
    under a tenth of a percent.
    """

    def __init__(self, sheet: Sheet, page, config: dict):
        box = page.mediabox
        self.page_w = float(box.width)
        self.page_h = float(box.height)
        self.left = float(box.left)
        self.bottom = float(box.bottom)
        self.rotate = int(page.get("/Rotate", 0) or 0) % 360

        self.dxf_w = sheet.ext_max[0] - sheet.ext_min[0]
        self.dxf_h = sheet.ext_max[1] - sheet.ext_min[1]
        if self.dxf_w <= 0 or self.dxf_h <= 0:
            raise ValueError("DXF has no usable extents ($EXTMIN/$EXTMAX)")
        self.min_x, self.min_y = sheet.ext_min

        # Visible page size, i.e. after the viewer applies /Rotate.
        if self.rotate in (90, 270):
            self.vis_w, self.vis_h = self.page_h, self.page_w
        else:
            self.vis_w, self.vis_h = self.page_w, self.page_h

        self.scale_x = self.vis_w / self.dxf_w
        self.scale_y = self.vis_h / self.dxf_h
        self.residual = abs(self.scale_x - self.scale_y) / max(self.scale_x, 1e-12)

        override = config.get("scale_override")
        if override:
            self.scale_x = self.scale_y = float(override)

        self.off_x, self.off_y = (float(v) for v in config.get("offset_pt", (0, 0)))

    def __call__(self, x: float, y: float) -> tuple[float, float]:
        # Position within the visible page, origin bottom-left, y up.
        vx = (x - self.min_x) * self.scale_x + self.off_x
        vy = (y - self.min_y) * self.scale_y + self.off_y
        # Undo /Rotate to get back to PDF user space.
        if self.rotate == 90:
            ux, uy = self.page_w - vy, vx
        elif self.rotate == 180:
            ux, uy = self.page_w - vx, self.page_h - vy
        elif self.rotate == 270:
            ux, uy = vy, self.page_h - vx
        else:
            ux, uy = vx, vy
        return ux + self.left, uy + self.bottom

    def rect(self, x0: float, y0: float, x1: float, y1: float,
             pad: float = 0.0) -> tuple[float, float, float, float]:
        """Transformed, normalised rectangle in PDF points."""
        ax, ay = self(x0, y0)
        bx, by = self(x1, y1)
        return (min(ax, bx) - pad, min(ay, by) - pad,
                max(ax, bx) + pad, max(ay, by) + pad)

    def point(self, x: float, y: float) -> tuple[float, float]:
        """A DXF point as fractions of the visible page, y from the top."""
        n = self.normalised(self.rect(x, y, x, y))
        return n["x"], n["y"]

    def normalised(self, rect: tuple[float, float, float, float]) -> dict:
        """Rectangle as fractions of the *visible* page, y measured from the
        top - which is what CSS wants."""
        x0, y0, x1, y1 = rect
        if self.rotate in (90, 270):
            # Re-express in visible space before normalising.
            corners = [(x0, y0), (x1, y1)]
            vs = []
            for cx, cy in corners:
                ux, uy = cx - self.left, cy - self.bottom
                if self.rotate == 90:
                    vs.append((uy, self.page_w - ux))
                else:
                    vs.append((self.page_h - uy, ux))
            xs = [v[0] for v in vs]
            ys = [v[1] for v in vs]
            x0, x1 = min(xs), max(xs)
            y0, y1 = min(ys), max(ys)
        else:
            x0, x1 = x0 - self.left, x1 - self.left
            y0, y1 = y0 - self.bottom, y1 - self.bottom
        return {
            "x": x0 / self.vis_w,
            "y": (self.vis_h - y1) / self.vis_h,
            "w": (x1 - x0) / self.vis_w,
            "h": (y1 - y0) / self.vis_h,
        }


# --------------------------------------------------------------------------
# Hotspots
# --------------------------------------------------------------------------

def _placeholder_key(column: str) -> str:
    return "".join(ch for ch in column.upper() if ch.isalnum())


def _column_value(row: TableRow | None, config: dict, key: str) -> str:
    """First non-empty value among the columns named by config[key]."""
    if row is None:
        return ""
    lookup = {c.strip().upper(): c for c in row.values}
    for candidate in config.get(key, []):
        col = lookup.get(candidate.strip().upper())
        if col and row.values[col]:
            return row.values[col]
    return ""


def _part_of(row: TableRow | None, config: dict) -> str:
    return _column_value(row, config, "part_columns")


def _vault_part(file: str, pn: str = "") -> str:
    """The vault's part number: the card's StockedPartNumber when it has one,
    else the file's name - how the drawing's own Part # column reads."""
    return (pn or "").strip() or os.path.splitext(os.path.basename(file or ""))[0].upper()


def _vault_parts(sheet: Sheet, sheet_index: int, row_files: dict[tuple[int, str], str],
                 config: dict, log=print, row_pns: dict | None = None,
                 row_cells: dict | None = None) -> None:
    """Put the vault's part number into each BOM row's Part # cell - the
    card's StockedPartNumber, else the file name (_vault_part) - adding the
    column when the drawing's table does not show one (the -M BOM keeps it
    hidden).  Runs before the hotspots are built, so the cards, the cart and
    {PART} all read the one value.  A row the macro found no file for keeps
    whatever its cell said, blank on a hidden column.

    Every row whose drawing cell (the SolidWorks property, which the macro
    still reads off the hidden column) disagrees with the vault is listed in
    the log by item: the property is being phased out for the card, and a
    disagreement is a card or a property to fix."""
    if (config.get("part_source") or "vault").lower() != "vault" or not sheet.table:
        return
    table = sheet.table
    names = {c.strip().upper() for c in config.get("part_columns", [])}
    col = next((c for c in table.columns if c.strip().upper() in names), None)
    if col is None:
        col = config.get("part_column_label") or "PART #"
        desc = {c.strip().upper() for c in config.get("description_columns", [])}
        at = next((i + 1 for i, c in enumerate(table.columns) if c.strip().upper() in desc),
                  min(1, len(table.columns)))
        table.columns = table.columns[:at] + [col] + table.columns[at:]
        for row in table.rows:
            row.values = {c: row.values.get(c, "") for c in table.columns}
    found = stocked = 0
    differ = []
    for row in table.rows:
        key = (sheet_index, str(row.item).strip())
        f = row_files.get(key)
        if not f:
            continue
        pn = (row_pns or {}).get(key, "")
        part = _vault_part(f, pn)
        found += 1
        stocked += 1 if pn else 0
        cell = ((row_cells or {}).get(key) or row.values.get(col, "")).strip()
        if cell and cell.upper() != part.upper():
            differ.append(f"item {row.item}: drawing {cell} / vault {part}"
                          + ("" if pn else " (file name - card has no StockedPartNumber)"))
        row.values[col] = part
    log(f"    part numbers from the vault: {found} of {len(table.rows)} row(s), "
        f"{stocked} StockedPartNumber, {found - stocked} file name"
        + (f", {len(table.rows) - found} without a file keep the cell" if found < len(table.rows) else ""))
    if differ:
        log(f"    WARNING: {len(differ)} row(s) where the drawing's Part # and the vault disagree "
            f"(the page shows the vault's):")
        for d in differ:
            log(f"      {d}")


# Length forms and the mm in one unit of the UOM a price is quoted in - the
# same table the app's belt checklist uses (matUomMm / MAT_LEN_FORMS).
_LEN_FORMS = {"CHAIN", "ROLL", "TUBE", "BAR", "ROD", "HEX", "THREAD", "SPLINE", "MULTI"}
_UOM_MM = {"FT": 304.8, "FEET": 304.8, "FOOT": 304.8, "IN": 25.4, "INCH": 25.4,
           "M": 1000.0, "METER": 1000.0, "METRE": 1000.0, "MM": 1.0}


def _load_materials(config: dict, log) -> dict[str, dict]:
    """MATERIALS keyed by stocked part number (upper case), from the data API.

    Empty - and the macro's prices stand - when polyapp_api is blank or the
    server cannot be reached; the page still publishes.
    """
    base = (config.get("polyapp_api") or "").strip().rstrip("/")
    if not base:
        return {}
    import urllib.request
    try:
        with urllib.request.urlopen(f"{base}/api/table/MATERIAL?limit=100000", timeout=90) as r:
            rows = json.loads(r.read().decode("utf-8"))
    except Exception as e:  # noqa: BLE001 - any failure means "no repricing", never "no page"
        log(f"  WARNING: MATERIALS not read from {base} ({e}) - length materials keep the macro's price.")
        return {}
    out = {str(m.get("StockedPartNumber") or "").strip().upper(): m for m in rows}
    out.pop("", None)
    log(f"  MATERIALS: {len(out)} row(s) from {base} - length materials priced on their length.")
    return out


def _num(v) -> float:
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return 0.0


def _length_price(hotspot: dict, mats: dict[str, dict],
                  cells: list[str] | None = None) -> tuple[float, str] | None:
    """(cost, basis) for a row that is a length material, else None.

    The row names its stock # somewhere - the description, the part column, or
    a STOCK # column of the drawing's own BOM (cells: every value on the row;
    "40-CHAIN SOLID | 14002 | (8) | 1208.5" carries it in none of the named
    columns).  A chain,
    belt or hose imports its consumed LENGTH as the BOM quantity (mm), so the
    cost is RawCost per purchase unit x that length - the purchase unit being
    the row's Length, or one of the UOM it is quoted in when Length is blank.
    """
    for v in (hotspot.get("description"), hotspot.get("part"), *(cells or [])):
        key = str(v or "").strip().upper()
        m = mats.get(key)
        if not m:
            continue
        form = str(m.get("RawForm") or "").strip().upper()
        uom = str(m.get("UnitOfMeasure") or "").strip().upper()
        length, cost = _num(m.get("Length")), _num(m.get("RawCost"))
        per = length if length > 0 and form != "SHEET" else (_UOM_MM.get(uom, 0.0) if form in _LEN_FORMS else 0.0)
        q = _num(hotspot.get("qty"))
        if per > 0 and cost > 0 and q > 0 and (q >= 25 or q != int(q)):
            unit = f"{uom}" if length <= 0 and uom else f"{per:g} mm"
            return cost / per * q, f"LENGTH {key}: ${cost:.2f}/{unit} x {q:g} mm"
        return None
    return None


def _qty_of(row: TableRow | None, config: dict) -> float | int:
    """The BOM quantity, decimals and all.

    Cut-to-length items carry a real measurement (1407.35 mm of chain), so this
    cannot round to an integer.  Spaces are dropped first, which also repairs a
    wrap that _join_fragments did not catch.
    """
    raw = _column_value(row, config, "qty_columns").replace(",", "")
    found = re.search(r"\d+(?:\.\d+)?", raw.replace(" ", ""))
    if not found:
        return 1
    value = float(found.group(0))
    if value <= 0:
        return 1
    return int(value) if value.is_integer() else round(value, 4)


def _message_for(item: str, row: TableRow | None, drawing: str) -> str:
    lines = [f"ITEM {item}" if item else "ITEM"]
    if row is None:
        lines.append("(no matching BOM row on this sheet)")
    else:
        item_column = next(iter(row.values), None)
        for column, value in row.values.items():
            if value and column != item_column:  # the item number is the title
                lines.append(f"{column.strip()}: {value}")
    lines.append(f"DRAWING: {drawing}")
    return "\n".join(lines)


def _url_for(item: str, row: TableRow | None, drawing: str, config: dict) -> str:
    template = config.get("link_template") or ""
    if not template:
        return ""
    values = {"ITEM": item, "DRAWING": drawing, "PART": _part_of(row, config)}
    if row is not None:
        for column, value in row.values.items():
            values[_placeholder_key(column)] = value
    out = template
    for key, value in values.items():
        out = out.replace("{" + key + "}", str(value))
    return out


def build_hotspots(sheet: Sheet, transform: Transform, drawing: str,
                   config: dict, sheet_index: int = 0) -> list[dict]:
    rows_by_item: dict[str, TableRow] = {}
    if sheet.table:
        for row in sheet.table.rows:
            rows_by_item.setdefault(row.item, row)

    def common(item: str, row: TableRow | None) -> dict:
        return {
            "sheet": sheet_index,
            "item": item,
            "part": _part_of(row, config),
            "config": _column_value(row, config, "config_columns"),
            "description": _column_value(row, config, "description_columns"),
            "qty": _qty_of(row, config),
            "message": _message_for(item, row, drawing),
            "url": _url_for(item, row, drawing, config),
        }

    hotspots: list[dict] = []

    balloon_pad = float(config.get("balloon_padding_pt", 0.0))
    for balloon in sheet.balloons:
        rect = transform.rect(balloon.x0, balloon.y0, balloon.x1, balloon.y1,
                              balloon_pad)
        hotspots.append({
            "kind": "balloon",
            **common(balloon.item, rows_by_item.get(balloon.item)),
            "rect_pt": [round(v, 3) for v in rect],
            **{k: round(v, 6) for k, v in transform.normalised(rect).items()},
        })

    row_pad = float(config.get("row_padding_pt", 0.0))
    if sheet.table:
        for row in sheet.table.rows:
            rect = transform.rect(row.x0, row.y0, row.x1, row.y1, row_pad)
            hotspots.append({
                "kind": "row",
                **common(row.item, row),
                "values": row.values,
                "rect_pt": [round(v, 3) for v in rect],
                **{k: round(v, 6) for k, v in transform.normalised(rect).items()},
            })

    return hotspots


# --------------------------------------------------------------------------
# Zones - where each SECTION sits on MAIN, and each top-level component's
# SILHOUETTE on its own sheet
# --------------------------------------------------------------------------
#
# A -M drawing is curated: MAIN shows the machine, every other sheet shows one
# section, and the section is whatever -M assembly the sheet's view references
# - not necessarily a subassembly the designer made.  So the sections are read
# FROM THE SHEETS, and the machine's model (the TLA behind MAIN) is only where
# their parts are looked up.  For each section sheet the zone on MAIN is the
# convex hull of the projected 3D boxes of the parts MAIN shows that sit under
# that assembly's -X twin in the TLA.  It is where the section IS, whether or
# not its edges are visible, so a transparent guard in front of it swallows
# nothing.
#
# The macro writes what only SolidWorks knows, tab-separated (VBA writes it):
#   MAIN     <sheet> <view> <tla file> <sheetW> <sheetH> <ox0> <oy0> <ox1> <oy1> <16 transform terms>
#   SECTION  <sheet> <assembly file>
#   PART     <instance path> <file> <x0> <y0> <z0> <x1> <y1> <z1>   (visible in MAIN, box in TLA space)
#   ASSY     <instance path> <file> <x0> <y0> <z0> <x1> <y1> <z1>   (every subassembly, same box)
#   TOP      <sheet> <instance name> <file> <flags> [<bitmap>]      (top-level of that sheet's assembly;
#            flags S suppressed, H hidden, E envelope, R in the REF folder;
#            bitmap: the component's own preview, saved beside thumbs.txt)
# Sheet indices are 0-based DXF order, the same the hotspots carry.
#
# The silhouettes need none of that to be DRAWN - the DXF already holds one
# block per top-level component (see _block_key) - only to be NAMED.

REF_TOKEN = re.compile(r"(^!|(^|[^a-z0-9])ref([^a-z0-9]|$))", re.IGNORECASE)


def _read_manifest(path: str) -> list[str]:
    """The lines of a macro-written manifest, whatever it was encoded in.

    zones.txt is written through ADODB as UTF-8 with a BOM; thumbs.txt comes
    out of a FileSystemObject TextStream, which is the Windows code page and
    no BOM.  A part or assembly named with an em dash - ALSFS-0002-M—4696 is
    a real one - is byte 0x97 in the second, which UTF-8 refuses outright.
    So: UTF-8 first, and the machine's code page when that fails.
    """
    with open(path, "rb") as handle:
        data = handle.read()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            text = data.decode("mbcs")
        except (UnicodeDecodeError, LookupError):
            text = data.decode("cp1252", errors="replace")
    return text.splitlines()


def _load_zones(path: str, log=print) -> dict | None:
    """Read the macro's zones manifest.  None when there is nothing to read."""
    if not path:
        return None
    if not os.path.isfile(path):
        log(f"  WARNING: no zones manifest at {path} - no zones.")
        return None
    zones = {"main": None, "sections": {}, "parts": [], "assemblies": [], "top": {},
             "views": {}, "refparts": {}, "pdmpn": {}}
    if True:
        for raw in _read_manifest(path):
            f = raw.split("\t")
            kind = f[0].strip().upper() if f else ""
            try:
                if kind == "MAIN" and len(f) >= 26:
                    zones["main"] = {
                        "sheet": int(f[1]), "view": f[2], "file": f[3],
                        "w": float(f[4]), "h": float(f[5]),
                        "outline": [float(v) for v in f[6:10]],
                        "xf": [float(v) for v in f[10:26]],
                    }
                elif kind == "SECTION" and len(f) >= 3:
                    zones["sections"][int(f[1])] = f[2].strip()
                elif kind == "VIEW" and len(f) >= 25:
                    # a section sheet's model view: sheet size, view outline
                    # and model-to-sheet transform, like MAIN's own line -
                    # what places the instance boxes on the page
                    zones["views"][int(f[1])] = {
                        "view": f[2], "w": float(f[3]), "h": float(f[4]),
                        "outline": [float(v) for v in f[5:9]],
                        "xf": [float(v) for v in f[9:25]],
                    }
                elif kind in ("PART", "ASSY") and len(f) >= 9:
                    zones["parts" if kind == "PART" else "assemblies"].append({
                        "name": f[1].strip(), "file": f[2].strip().lower(),
                        "box": [float(v) for v in f[3:9]]})
                elif kind == "TOP" and len(f) >= 4:
                    flags = f[4].strip().upper() if len(f) > 4 else ""
                    zones["top"].setdefault(int(f[1]), []).append({
                        "name": f[2].strip(), "file": f[3].strip().lower(),
                        "suppressed": "S" in flags, "hidden": "H" in flags,
                        "envelope": "E" in flags, "ref": "R" in flags,
                        "bitmap": f[5].strip() if len(f) > 5 else "",
                        # the referenced configuration - a 2026-10 macro writes
                        # it; an older manifest leaves it "" and parts group
                        # by file alone
                        "config": f[6].strip() if len(f) > 6 else "",
                        # the instance's box in its assembly's space, when
                        # the macro could read one - six numbers after the
                        # configuration
                        "box": ([float(v) for v in f[7:13]]
                                if len(f) >= 13 and all(x.strip() for x in f[7:13]) else None)})
                elif kind == "PDMPN" and len(f) >= 4 and f[3].strip():
                    # the vault card's StockedPartNumber for a REFPART file
                    # and configuration (2026-10-13); none = the file name
                    zones["pdmpn"][(f[1].strip().lower(), f[2].strip().upper())] = f[3].strip()
                elif kind == "REFPART" and len(f) >= 10:
                    # a PART under the section's reference assembly, at any
                    # depth (a 2026-10-09 macro): the ALL PARTS panel.  Name,
                    # file, flags, bitmap, configuration, the drawing layer
                    # its lines took in the REF-only export, the file's vault
                    # version, description, then the instance's box in the
                    # view's assembly space.
                    flags = f[4].strip().upper()
                    zones["refparts"].setdefault(int(f[1]), []).append({
                        "name": f[2].strip(), "file": f[3].strip().lower(),
                        "hidden": "H" in flags, "bitmap": f[5].strip(),
                        "config": f[6].strip(), "layer": f[7].strip().upper(),
                        "ver": f[8].strip(), "desc": f[9].strip(),
                        "box": ([float(v) for v in f[10:16]]
                                if len(f) >= 16 and all(x.strip() for x in f[10:16]) else None)})
            except ValueError:
                continue
    return zones


def _twin_file(assembly: str) -> str:
    """The design assembly a curated one stands for: A4696-IC00-M.SLDASM ->
    a4696-ic00-x.sldasm, and ALSFS-0002-M-4696 -> alsfs-0002-x.sldasm."""
    stem = assembly.lower().rsplit(".", 1)[0]
    cut = stem.rfind("-m")
    return (stem[:cut] if cut > 0 else stem) + "-x.sldasm"


def _sw_transform(xf: list[float], p: tuple[float, float, float]) -> tuple[float, float]:
    """A SolidWorks MathTransform (9 rotation terms, 3 translation, scale)
    applied to a point - row-vector convention, p' = s * (p . R) + t."""
    x, y, z = p
    return (xf[12] * (x * xf[0] + y * xf[3] + z * xf[6]) + xf[9],
            xf[12] * (x * xf[1] + y * xf[4] + z * xf[7]) + xf[10])


def _hull(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    pts = sorted(set(points))
    if len(pts) < 3:
        return pts

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])
    lower: list = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper: list = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def _poly_area(poly) -> float:
    a = 0.0
    for i in range(len(poly)):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % len(poly)]
        a += x0 * y1 - x1 * y0
    return abs(a) / 2


def build_sections(zones: dict, labels: list[str], log=print) -> list[dict]:
    """The MAIN sheet's click zones, one per section sheet.  Fractions of the
    sheet, y from the top, like every other hotspot."""
    main = zones.get("main")
    if not main or not zones.get("sections"):
        return []
    W, H, xf = main["w"], main["h"], main["xf"]
    if W <= 0 or H <= 0:
        return []

    def frac(p3):
        x, y = _sw_transform(xf, p3)
        return (x / W, 1 - y / H)

    def corners(box):
        return [(x, y, z) for x in (box[0], box[3]) for y in (box[1], box[4]) for z in (box[2], box[5])]

    ox0, oy0, ox1, oy1 = main["outline"]
    out_w, out_h = (ox1 - ox0) / W, (oy1 - oy0) / H

    def placed(entries):
        out = []
        for p in entries:
            segs = [g.lower() for g in p["name"].split("/")]
            stems = [g.rsplit("-", 1)[0] if g.rsplit("-", 1)[-1].isdigit() else g for g in segs]
            pts = [frac(c) for c in corners(p["box"])]
            xs = [q[0] for q in pts]
            ys = [q[1] for q in pts]
            # An envelope or a stray body can carry a box the size of a
            # building; a box wider than the view itself is not placing anything.
            if out_w > 0 and (max(xs) - min(xs) > out_w * 1.05 or max(ys) - min(ys) > out_h * 1.05):
                continue
            out.append({"stems": stems, "file": p["file"], "pts": pts})
        return out
    parts = placed(zones["parts"])
    assemblies = placed(zones.get("assemblies", []))

    sections = []
    for idx in sorted(zones["sections"]):
        assembly = zones["sections"][idx]
        twin = _twin_file(assembly)
        stem = twin.rsplit(".", 1)[0]
        # The parts MAIN shows under the twin place the section where it is
        # drawn.  A section shown by nothing at all - behind a guard, or all
        # hardware - falls back to the twin assembly's own box, which is where
        # it would be drawn.
        src = [p for p in parts if stem in p["stems"]]
        rule = "shown"
        if not src:
            src = [a for a in assemblies if a["file"] == twin]
            rule = "boxed"
        label = labels[idx] if idx < len(labels) else f"Sheet {idx + 1}"
        if not src:
            log(f"  zones: {label} - {twin} is not in the machine; no zone on MAIN.")
            continue
        poly = _hull([q for p in src for q in p["pts"]])
        if len(poly) < 3:
            continue
        sections.append({"sheet": idx, "label": label, "assembly": assembly,
                         "parts": len(src), "rule": rule,
                         "poly": [[round(x, 5), round(y, 5)] for x, y in poly],
                         "area": round(_poly_area(poly), 6)})
    if sections:
        log(f"  zones: {len(sections)} of {len(zones['sections'])} section(s) placed on MAIN")
    return sections


def _comp_matches(key: str, comp: dict) -> bool:
    n = re.sub(r"[^a-z0-9]", "_", comp["name"].lower())
    k = re.sub(r"[^a-z0-9]", "_", key)
    return k == n[:10] or (len(n) < 10 and k.startswith(n + "_"))


def _is_ref(comp: dict) -> bool:
    """The reference assembly a drawing carries on top of everything - drawn
    transparent, filed in a feature folder named REF (the macro flags it R),
    or named !REF.  Never a click target."""
    return bool(comp.get("ref")) or bool(REF_TOKEN.search(comp["name"])) or \
        bool(REF_TOKEN.search(comp["file"].rsplit(".", 1)[0]))


def _png_uri(img) -> str:
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _draw_silhouette(px: list, px_w: int, px_h: int, colour: str) -> dict | None:
    """One component's linework, in page pixels, as the tinted overlay the page
    shows and the fattened quarter-scale mask it hit-tests - the place on the
    page as fractions, like the hotspots.  None when there is nothing to draw
    or Pillow is missing.  Shared by the silhouettes and the ALL PARTS list."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    xs = [x for pl in px for x, _ in pl]
    ys = [y for pl in px for _, y in pl]
    if not xs:
        return None
    x0, y0 = max(0, int(min(xs)) - 4), max(0, int(min(ys)) - 4)
    x1, y1 = min(px_w, int(max(xs)) + 5), min(px_h, int(max(ys)) + 5)
    cw, ch = x1 - x0, y1 - y0
    if cw <= 0 or ch <= 0:
        return None
    rgb = tuple(int(colour[i:i + 2], 16) for i in (1, 3, 5))
    over = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
    draw = ImageDraw.Draw(over)
    for pl in px:
        draw.line([(x - x0, y - y0) for x, y in pl], fill=rgb + (230,), width=3)
    mask = Image.new("L", (max(1, cw // 4), max(1, ch // 4)), 0)
    mdraw = ImageDraw.Draw(mask)
    for pl in px:
        mdraw.line([((x - x0) / 4, (y - y0) / 4) for x, y in pl], fill=255, width=5)
    return {"x": round(x0 / px_w, 6), "y": round(y0 / px_h, 6),
            "w": round(cw / px_w, 6), "h": round(ch / px_h, 6), "area": cw * ch,
            "overlay": _png_uri(over), "mask": _png_uri(mask)}


def build_silhouettes(sheet: Sheet, transform: Transform, comps: list[dict],
                      config: dict, log=print, row_files: dict | None = None,
                      images: dict | None = None, bitmap_dir: str = "",
                      view: dict | None = None
                      ) -> tuple[list[dict], str, list[str]]:
    """Every top-level component's own linework on this sheet, as a tinted
    overlay and a fattened hit mask.  Returns (silhouettes, ref overlay, ref
    names).  Coordinates are fractions of the page, like the hotspots.

    Needs Pillow; without it nothing is drawn and the caller says so once."""
    if not sheet.blocks and not sheet.ref_blocks:
        return [], "", []
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return [], "", []

    dpi = float(config.get("silhouette_dpi", 120) or 120)
    px_w = max(1, int(round(transform.vis_w * dpi / 72)))
    px_h = max(1, int(round(transform.vis_h * dpi / 72)))
    palette = config.get("silhouette_colours") or [
        "#d81b60", "#1e88e5", "#43a047", "#fb8c00", "#8e24aa", "#00acc1", "#6d4c41",
        "#c0ca33", "#e53935", "#3949ab", "#00897b", "#f4511e", "#5e35b1", "#039be5"]

    def encode(img) -> str:
        buf = io.BytesIO()
        img.save(buf, "PNG", optimize=True)
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")

    live = [c for c in comps if not c.get("suppressed")]
    ref_names = sorted({c["name"] for c in live if _is_ref(c)})

    # The BOM row a component stands for.  By FILE first, when the manifest
    # says which model each row resolved to - the only link that holds when the
    # BOM's part number is a purchased number (30014) and the file is the
    # vendor's name for it (TN-71-B3-B14).  Then by part number, compared with
    # everything but letters and digits dropped ("SITI_0.5HP_REDUCER" against
    # "SITI_0_5HP_REDUCER").
    def norm(s: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", (s or "").upper())
    row_items: dict[str, str] = {}
    if sheet.table:
        for row in sheet.table.rows:
            p = norm(_part_of(row, config))
            if p:
                row_items.setdefault(p, row.item)
    file_items: dict[str, str] = {}             # file -> item (first row)
    filecfg_items: dict[tuple[str, str], str] = {}  # (file, CONFIG) -> item
    for k, item in (row_files or {}).items():
        f, c = (k if isinstance(k, tuple) else (k, ""))
        file_items.setdefault(f.lower(), item)
        filecfg_items.setdefault((f.lower(), (c or "").upper()), item)

    # The component's own preview, when the macro saved one - keyed like the
    # BOM previews, shared by every silhouette that names the same bitmap.
    size = int(config.get("thumb_px", 260) or 260)
    quality = int(config.get("thumb_quality", 82) or 82)

    def preview_key(matched: list[dict]) -> str:
        if images is None or not bitmap_dir:
            return ""
        for c in matched:
            bmp = c.get("bitmap") or ""
            if not bmp:
                continue
            key = os.path.splitext(bmp)[0]
            if key not in images:
                encoded = _encode_thumb(os.path.join(bitmap_dir, bmp), size, quality)
                if encoded is None:
                    continue
                images[key] = encoded
            return key
        return ""

    out, ref_lines = [], []
    # The REF-only export, when the macro wrote one: every line of it is the
    # reference assembly, whatever its blocks are called.
    for polylines in sheet.ref_blocks.values():
        for pl in polylines:
            ref_lines.append([(fx * px_w, fy * px_h)
                              for fx, fy in (transform.point(x, y) for x, y in pl)])
    # One silhouette per PART, not per block.  SolidWorks writes one block per
    # instance, and a short instance name gets the assembly's name appended
    # before the 12-character clip - "11022-1_A469", "11022-2_A469", ... - so
    # the eight instances of one bushing came out as eight keys, eight
    # silhouettes and eight colours.  Blocks whose components are the same
    # file(s) are folded together here, so every instance of a part lights
    # with its partners, in one colour.  A block the manifest cannot name
    # stays on its own key.
    # WHICH instance a block is.  A long instance name loses its instance
    # number to the 12-character clip and the dedupe counter ("AXXDB-0002-0",
    # "AXXDB-0002-1" ...), so the name alone cannot say which of eight dancer
    # bars - in three configurations - a block draws.  The manifest can: the
    # macro writes each section view's transform (VIEW) and each top-level
    # instance's box in its assembly's space (on the TOP line), and the block
    # whose outline sits inside an instance's projected box IS that instance.
    # Without the VIEW line (an older macro) the block keeps every instance
    # its name could be, as before.
    def comp_rect(c: dict):
        if not view or not c.get("box"):
            return None
        W, H = view["w"], view["h"]
        if W <= 0 or H <= 0:
            return None
        b = c["box"]
        pts = [_sw_transform(view["xf"], (x, y, z))
               for x in (b[0], b[3]) for y in (b[1], b[4]) for z in (b[2], b[5])]
        xs = [x / W for x, _ in pts]
        ys = [1 - y / H for _, y in pts]
        return (min(xs), min(ys), max(xs), max(ys))

    def pick_instance(matched: list[dict], polylines) -> list[dict]:
        if len(matched) < 2 or not view:
            return matched
        pts = [transform.point(x, y) for pl in polylines for x, y in pl]
        if not pts:
            return matched
        bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts)
        bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
        barea = max((bx1 - bx0) * (by1 - by0), 1e-12)
        best, best_score = None, 0.0
        for c in matched:
            r = comp_rect(c)
            if r is None:
                continue
            ix = max(0.0, min(bx1, r[2]) - max(bx0, r[0]))
            iy = max(0.0, min(by1, r[3]) - max(by0, r[1]))
            inter = ix * iy
            rarea = max((r[2] - r[0]) * (r[3] - r[1]), 1e-12)
            # how much of the block the box covers, tempered by how much of
            # the box the block fills - a frame that covers everything does
            # not win every bracket
            score = (inter / barea) * (0.5 + 0.5 * min(1.0, inter / rarea))
            if score > best_score:
                best, best_score = c, score
        return [best] if best is not None and best_score > 0.25 else matched

    groups: dict[str, dict] = {}
    for name, polylines in sorted(sheet.blocks.items()):
        key = _block_key(name)
        matched = [c for c in live if _comp_matches(key, c)]
        if matched and all(_is_ref(c) for c in matched):
            for pl in polylines:
                ref_lines.append([(fx * px_w, fy * px_h)
                                  for fx, fy in (transform.point(x, y) for x, y in pl)])
            continue
        matched = pick_instance(matched, polylines)
        # ... and by CONFIGURATION within a file: a 16 and a 20 of the same
        # part are two things to order, so they are two silhouettes and two
        # colours, each wired to its own BOM row.  The manifest's TOP lines
        # carry the referenced configuration (a 2026-10 macro); an older one
        # leaves it blank and the file alone decides.
        parts = sorted({c["file"] + "@" + (c.get("config") or "").upper() for c in matched})
        gid = ("file:" + "|".join(parts)) if parts else ("key:" + key)
        g = groups.setdefault(gid, {"key": key, "polylines": [], "matched": []})
        g["polylines"].extend(polylines)
        for c in matched:
            if not any(c is m for m in g["matched"]):
                g["matched"].append(c)
    # Which files appear under more than one configuration on this sheet - those
    # labels say which configuration they are.
    cfgs_of_file: dict[str, set] = {}
    for g in groups.values():
        for c in g["matched"]:
            cfgs_of_file.setdefault(c["file"], set()).add((c.get("config") or "").upper())
    # The page indexes silhouettes by key, so two groups split off one block
    # key - the 0611 and the 0660 of one dancer bar - must not share it, or
    # the second overwrites the first and only one of them can ever light.
    used_keys: set = set()
    for gi, g in enumerate(groups.values()):
        key, polylines, matched = g["key"], g["polylines"], g["matched"]
        if key in used_keys:
            n = 2
            while f"{key}~{n}" in used_keys:
                n += 1
            key = f"{key}~{n}"
        used_keys.add(key)
        gcfgs = {(c.get("config") or "").upper() for c in matched}
        gcfg = next(iter(gcfgs)) if len(gcfgs) == 1 else ""
        px = []
        for pl in polylines:
            px.append([(fx * px_w, fy * px_h) for fx, fy in (transform.point(x, y) for x, y in pl)])
        colour = palette[gi % len(palette)]
        drawn = _draw_silhouette(px, px_w, px_h, colour)
        if drawn is None:
            continue
        files = sorted({c["file"] for c in matched})
        names = [c["name"] for c in matched]
        stems = sorted({f.rsplit(".", 1)[0].upper() for f in files} |
                       {re.sub(r"-\d+$", "", n).upper() for n in names})
        if len(files) == 1:
            label = files[0].rsplit(".", 1)[0].upper()
            if gcfg and len(cfgs_of_file.get(files[0], ())) > 1:
                label += " [" + gcfg + "]"
        elif len(names) == 1:
            label = names[0]
        else:
            label = key.upper() + ("" if files else "")
        # The row: this file in THIS configuration first, then the file in any,
        # then the part number.
        item = next((filecfg_items[(f, gcfg)] for f in files if (f, gcfg) in filecfg_items), "")
        if not item:
            item = next((file_items[f] for f in files if f in file_items), "")
        if not item:
            item = next((row_items[norm(p)] for p in stems if norm(p) in row_items), "")
        out.append({
            "key": key, "label": label, "colour": colour,
            "names": names[:24], "n_names": len(names), "files": files[:12],
            "parts": stems[:16], "item": item, "thumb": preview_key(matched),
            "blocks": len(polylines),
            **drawn,
        })
    ref_overlay = ""
    if ref_lines:
        over = Image.new("RGBA", (px_w, px_h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(over)
        for pl in ref_lines:
            draw.line(pl, fill=(120, 120, 120, 150), width=2)
        ref_overlay = encode(over)
    return out, ref_overlay, ref_names


def build_allparts(sheet: Sheet, transform: Transform, refparts: list[dict],
                   config: dict, log=print, row_files: dict | None = None,
                   images: dict | None = None, bitmap_dir: str = "",
                   view: dict | None = None, show_version: bool = True,
                   pdmpn: dict | None = None) -> list[dict]:
    """Every PART under this sheet's reference assembly - the design twin the
    -M assembly files under !REF - one entry per file and configuration, with
    its preview, how many of it there are, and where each instance sits on
    the sheet.  This is the ALL PARTS panel: the curated parts list is the
    BOM, and this is the rest of the module, which the BOM leaves out on
    purpose.

    Each entry carries `places`, one per instance where one could be found,
    smallest first.  Three sources, best first.  The macro put each part's
    drawing component on a layer of its own for the REF-only export
    (dl_PartLayersOn), and the REFPART line names it: that layer's lines are
    the instance's own outline.  Failing that, the export's block names are
    read the way the main export's are, and a block the manifest can match
    to exactly one part is that part's outline (all instances in one).
    Failing both, the instance box through the sheet's VIEW transform - a
    rectangle, one per instance, never a union, and only while it is small
    enough to mean something (config allparts_box_max).  A part with none of
    these is listed and pictured and simply has nowhere to light.

    Coordinates are fractions of the page, like the hotspots."""
    if not refparts:
        return []

    def norm(s: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", (s or "").upper())

    # One entry per file + configuration: a 16 and a 20 of one bushing are two
    # things, the eight instances of one bushing are one thing, eight times.
    groups: dict[tuple[str, str], dict] = {}
    order: list[tuple[str, str]] = []
    for c in refparts:
        gk = (c["file"], (c.get("config") or "").upper())
        g = groups.get(gk)
        if g is None:
            g = groups[gk] = {"file": c["file"], "config": c.get("config") or "", "comps": []}
            order.append(gk)
        g["comps"].append(c)
    cfgs_of_file: dict[str, set] = {}
    for f, cfg in groups:
        cfgs_of_file.setdefault(f, set()).add(cfg)

    # The BOM row a part stands for when the curated list has it too - by file
    # first, then by part number - so its row's card opens, as a silhouette's
    # does, and the row reads "(item n)".
    row_items: dict[str, str] = {}
    if sheet.table:
        for row in sheet.table.rows:
            p = norm(_part_of(row, config))
            if p:
                row_items.setdefault(p, row.item)
    file_items: dict[str, str] = {}
    filecfg_items: dict[tuple[str, str], str] = {}
    for k, item in (row_files or {}).items():
        f, c = (k if isinstance(k, tuple) else (k, ""))
        file_items.setdefault(f.lower(), item)
        filecfg_items.setdefault((f.lower(), (c or "").upper()), item)

    size = int(config.get("thumb_px", 260) or 260)
    quality = int(config.get("thumb_quality", 82) or 82)

    def preview_key(comps: list[dict]) -> str:
        if images is None or not bitmap_dir:
            return ""
        for c in comps:
            bmp = c.get("bitmap") or ""
            if not bmp:
                continue
            key = os.path.splitext(bmp)[0]
            if key not in images:
                encoded = _encode_thumb(os.path.join(bitmap_dir, bmp), size, quality)
                if encoded is None:
                    continue
                images[key] = encoded
            return key
        return ""

    def comp_rect(c: dict):
        if not view or not c.get("box"):
            return None
        W, H = view["w"], view["h"]
        if W <= 0 or H <= 0:
            return None
        b = c["box"]
        pts = [_sw_transform(view["xf"], (x, y, z))
               for x in (b[0], b[3]) for y in (b[1], b[4]) for z in (b[2], b[5])]
        xs = [x / W for x, _ in pts]
        ys = [1 - y / H for _, y in pts]
        r = (min(xs), min(ys), max(xs), max(ys))
        # A box wider than the sheet is an envelope or a fault, not a part.
        if r[2] - r[0] > 1 or r[3] - r[1] > 1:
            return None
        return r

    dpi = float(config.get("silhouette_dpi", 120) or 120)
    px_w = max(1, int(round(transform.vis_w * dpi / 72)))
    px_h = max(1, int(round(transform.vis_h * dpi / 72)))
    palette = config.get("silhouette_colours") or [
        "#d81b60", "#1e88e5", "#43a047", "#fb8c00", "#8e24aa", "#00acc1", "#6d4c41",
        "#c0ca33", "#e53935", "#3949ab", "#00897b", "#f4511e", "#5e35b1", "#039be5"]

    # The REF layer's blocks, folded onto the parts they name.  A block the
    # manifest matches to no part, or to two parts sharing a ten-character
    # prefix, draws nothing - its part keeps the box.
    # The manifest names an instance by its whole path under the reference
    # ("A4696-RP00-X-1/X4696-RP04-1"); the block is named for the leaf.  A
    # block two configurations of one part could be - the 16 and the 20 of a
    # bar share a key - goes to whichever instance's projected box the block
    # sits in, as build_silhouettes decides it; with no box to ask, it draws
    # nothing and both keep their boxes.
    leaves = [dict(c, name=c["name"].rsplit("/", 1)[-1]) for c in refparts]

    def pick(matched: list[dict], polylines) -> list[dict]:
        if len(matched) < 2 or not view:
            return matched
        pts = [transform.point(x, y) for pl in polylines for x, y in pl]
        if not pts:
            return matched
        bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts)
        bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
        barea = max((bx1 - bx0) * (by1 - by0), 1e-12)
        best, best_score = None, 0.0
        for c in matched:
            r = comp_rect(c)
            if r is None:
                continue
            ix = max(0.0, min(bx1, r[2]) - max(bx0, r[0]))
            iy = max(0.0, min(by1, r[3]) - max(by0, r[1]))
            inter = ix * iy
            rarea = max((r[2] - r[0]) * (r[3] - r[1]), 1e-12)
            score = (inter / barea) * (0.5 + 0.5 * min(1.0, inter / rarea))
            if score > best_score:
                best, best_score = c, score
        return [best] if best is not None and best_score > 0.25 else matched

    drawn: dict[tuple[str, str], list] = {}
    for name, polylines in sorted(sheet.ref_blocks.items()):
        key = _block_key(name)
        matched = [c for c in leaves if _comp_matches(key, c)]
        gks = {(c["file"], (c.get("config") or "").upper()) for c in matched}
        if len(gks) > 1:
            gks = {(c["file"], (c.get("config") or "").upper()) for c in pick(matched, polylines)}
        if len(gks) != 1:
            continue
        drawn.setdefault(next(iter(gks)), []).extend(polylines)

    boxes_on = bool(config.get("allparts_boxes", True))
    box_max = float(config.get("allparts_box_max", 0.06) or 0)
    page_px = float(px_w * px_h)
    stats = {"layer": 0, "block": 0, "box": 0, "big": 0}

    def to_px(polylines):
        return [[(fx * px_w, fy * px_h) for fx, fy in (transform.point(x, y) for x, y in pl)]
                for pl in polylines]

    out = []
    for gi, gk in enumerate(order):
        g = groups[gk]
        comps = g["comps"]
        stem = g["file"].rsplit(".", 1)[0].upper()
        # what the Part # column shows: the card's StockedPartNumber, else
        # the file name (pdmpn: the zones manifest's PDMPN lines).  Matching
        # stays on the file (`parts`) - one stocked blank is on many cards.
        pn = _vault_part(g["file"], (pdmpn or {}).get(gk) or (pdmpn or {}).get((g["file"], "")))
        label = pn
        if gk[1] and len(cfgs_of_file.get(g["file"], ())) > 1:
            label += " [" + gk[1] + "]"
        item = (filecfg_items.get((g["file"], gk[1])) or file_items.get(g["file"])
                or row_items.get(norm(stem), ""))
        colour = palette[gi % len(palette)]
        places = []
        # 1. the instance's own layer in the REF export
        unplaced = []
        for c in comps:
            lines = sheet.ref_layers.get((c.get("layer") or "").upper()) if c.get("layer") else None
            placed = _draw_silhouette(to_px(lines), px_w, px_h, colour) if lines else None
            if placed:
                placed["kind"] = "outline"
                places.append(placed)
                stats["layer"] += 1
            else:
                unplaced.append(c)
        # 2. a block of the export that names this part - every instance at once
        if unplaced and drawn.get(gk):
            placed = _draw_silhouette(to_px(drawn[gk]), px_w, px_h, colour)
            if placed:
                placed["kind"] = "outline"
                places.append(placed)
                stats["block"] += 1
                unplaced = []
        # 3. the instance box, one each, small ones only
        if unplaced and boxes_on and box_max > 0:
            for c in unplaced:
                r = comp_rect(c)
                if r is None:
                    continue
                x0, y0 = max(0.0, r[0]), max(0.0, r[1])
                x1, y1 = min(1.0, r[2]), min(1.0, r[3])
                if x1 <= x0 or y1 <= y0:
                    continue
                if (x1 - x0) * (y1 - y0) > box_max:
                    stats["big"] += 1
                    continue
                places.append({"kind": "box", "x": round(x0, 6), "y": round(y0, 6),
                               "w": round(x1 - x0, 6), "h": round(y1 - y0, 6),
                               "area": (x1 - x0) * (y1 - y0) * page_px,
                               "overlay": "", "mask": ""})
                stats["box"] += 1
        places.sort(key=lambda p: p["area"])
        out.append({
            "key": "ap%d" % (gi + 1), "label": label, "file": g["file"],
            "config": g["config"], "count": len(comps),
            "desc": next((c["desc"] for c in comps if c.get("desc")), ""),
            # the file's vault version, as a BOM row carries it (pdm_ver),
            # shown in the Config cell when the page shows versions at all
            "pdm_ver": (next((c["ver"] for c in comps if c.get("ver")), "")
                        if show_version else ""),
            "names": [c["name"] for c in comps][:24], "n_names": len(comps),
            "files": [g["file"]], "parts": [stem], "pn": pn, "item": item,
            "thumb": preview_key(comps), "colour": colour,
            "hidden": all(c.get("hidden") for c in comps),
            "places": places,
            "area": places[0]["area"] if places else 0,
            # what tells the page's card and chip this is a reference part,
            # on the entry (the panel's rows) and on every zone made of it
            "ap": True,
        })
    out.sort(key=lambda e: e["label"])
    if out:
        log(f"    {len(out)} part(s) under the reference assembly ({len(refparts)} instance(s)): "
            f"{sum(1 for z in out if z['thumb'])} pictured, "
            f"{stats['layer']} instance(s) outlined from their layer, "
            f"{stats['block']} part(s) from a named block, "
            f"{stats['box']} boxed, {stats['big']} box(es) too big to show, "
            f"{sum(1 for z in out if not z['places'])} with nowhere to light, "
            f"{sum(1 for z in out if z['pdm_ver'])} versioned, "
            f"{sum(1 for z in out if z['item'])} also on the parts list")
    return out


# --------------------------------------------------------------------------
# PDF output
# --------------------------------------------------------------------------

def _js_string(value: str) -> str:
    """Escape a Python string for embedding in PDF JavaScript source."""
    return (value.replace("\\", "\\\\").replace('"', '\\"')
            .replace("\r", "").replace("\n", "\\n"))


def _uri_action(uri: str) -> DictionaryObject:
    return DictionaryObject({
        NameObject("/Type"): NameObject("/Action"),
        NameObject("/S"): NameObject("/URI"),
        NameObject("/URI"): TextStringObject(uri),
    })


def _annotation(hotspot: dict, config: dict, html_url: str) -> DictionaryObject:
    rect = hotspot["rect_pt"]
    annot = DictionaryObject({
        NameObject("/Type"): NameObject("/Annot"),
        NameObject("/Subtype"): NameObject("/Link"),
        NameObject("/Rect"): ArrayObject([FloatObject(v) for v in rect]),
        NameObject("/Border"): ArrayObject(
            [NumberObject(0), NumberObject(0), NumberObject(0)]),
        NameObject("/F"): NumberObject(4),  # print
        # Shown as a tooltip by viewers that support it.
        NameObject("/Contents"): TextStringObject(hotspot["message"]),
    })

    action = config.get("pdf_action", "html")

    if action == "url" and hotspot.get("url"):
        annot[NameObject("/A")] = _uri_action(hotspot["url"])
    elif action == "html" and html_url:
        # Open the quote page on this sheet with the item dropped into the
        # cart.  A plain URI action, so it works in Acrobat and in the browser
        # viewers - unlike PDF JavaScript, which only Acrobat runs.
        from urllib.parse import quote

        annot[NameObject("/A")] = _uri_action(
            f"{html_url}#sheet={int(hotspot.get('sheet', 0)) + 1}"
            f"&add={quote(str(hotspot['item']), safe='')}")
    else:
        title = f"Item {hotspot['item']}"
        js = (f'app.alert({{cMsg: "{_js_string(hotspot["message"])}", '
              f'cTitle: "{_js_string(title)}", nIcon: 3, nType: 0}});')
        annot[NameObject("/A")] = DictionaryObject({
            NameObject("/Type"): NameObject("/Action"),
            NameObject("/S"): NameObject("/JavaScript"),
            NameObject("/JS"): TextStringObject(js),
        })
    return annot


def _outline_operators(hotspots: list[dict], config: dict) -> bytes:
    """Raw PDF content operators drawing an outline around each hotspot."""
    width = float(config.get("box_width_pt", 0.6))
    balloon_rgb = config.get("box_rgb_balloon", [0.9, 0.45, 0.0])
    row_rgb = config.get("box_rgb_row", [0.0, 0.45, 0.85])

    ops = [b"q", f"{width} w".encode("ascii"), b"[2 2] 0 d"]
    for hotspot in hotspots:
        rgb = balloon_rgb if hotspot["kind"] == "balloon" else row_rgb
        x0, y0, x1, y1 = hotspot["rect_pt"]
        ops.append(f"{rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} RG".encode("ascii"))
        ops.append(
            f"{x0:.3f} {y0:.3f} {x1 - x0:.3f} {y1 - y0:.3f} re S".encode("ascii"))
    ops.append(b"Q")
    return b"\n".join(ops) + b"\n"


def _append_content(writer: PdfWriter, page, data: bytes) -> None:
    """Append a content stream to a page without disturbing its graphics state.

    The existing content is wrapped in q ... Q so our operators start from the
    page's initial state rather than whatever the drawing left behind.
    """
    contents = page.get(NameObject("/Contents"))
    if contents is None:
        return
    existing = list(contents) if isinstance(contents, ArrayObject) else [
        page.raw_get(NameObject("/Contents"))]

    def stream(payload: bytes):
        obj = DecodedStreamObject()
        obj.set_data(payload)
        return writer._add_object(obj)

    page[NameObject("/Contents")] = ArrayObject(
        [stream(b"q\n")] + existing + [stream(b"Q\n" + data)])


def build_pdf(src_pdf: str, out_pdf: str, by_page: dict[int, list[dict]],
              config: dict, draw_boxes: bool, html_url: str) -> None:
    """Annotate the source PDF, one set of hotspots per page it applies to."""
    writer = PdfWriter(clone_from=src_pdf)

    for page_index, hotspots in sorted(by_page.items()):
        if not hotspots:
            continue
        page = writer.pages[page_index]

        if draw_boxes:
            _append_content(writer, page, _outline_operators(hotspots, config))

        annots = page.get(NameObject("/Annots"))
        if not isinstance(annots, ArrayObject):
            annots = ArrayObject()
            page[NameObject("/Annots")] = annots
        for hotspot in hotspots:
            annots.append(
                writer._add_object(_annotation(hotspot, config, html_url)))

    with open(out_pdf, "wb") as handle:
        writer.write(handle)


# --------------------------------------------------------------------------
# HTML output
# --------------------------------------------------------------------------
#
# The page is no longer written here.  The chrome - Polypack's tokens, the
# header, the contact drawer, the left rail, the toast - lives in
# POLYPACK_SHELL.py, because the Technical Manual and the Word converter want
# the same page and three copies of it would drift apart the first time someone
# changed the red.  What stays here is the part that is genuinely about a
# drawing: the sheet viewer, the hotspot overlay, the part card, the parts list
# and the cart.
#
# That split is what lets the manual stop putting this page in an <iframe>.  A
# SECTION carries its own CSS/markup/behaviour namespaced under its own id, so
# any number of them share one file without colliding; `--emit section` writes
# one to disk for the manual to absorb.  The default output is still the
# standalone page, byte-for-byte the same idea as before - one section, poured
# into the shell - so the SolidWorks VBA macro and the parts desk see no change.

SECTION_ID = "ipl"          # illustrated parts list - the namespace for all of it

SECTION_CSS = r"""
  /* ===== The sheet, and the hotspots stamped over it ===== */
  /* ===== THE VIEWPORT =====
     Zoom grows .page past this box, and THIS is the element that scrolls, so
     panning is native scrolling.  Deliberately not a second transform: the
     overlay already carries one (the alignment nudge), and two transforms on the
     same picture have to be kept in step forever.  Growing the box instead means
     the hotspots — positioned in percentages inside .page — scale for free and
     stay exactly where they were aligned, and the raster stays crisp at any zoom
     rather than being resampled.
     `flex: 1 1 auto; min-width: 0` is what confines the growth to this column:
     the 440px aside and the shell's rail never see it, however far in you go.
     --fitH is the height the sheet is fitted to, named once because the page
     width below is derived from it and the two must not drift. */
  /* The parts list scrolls on its own; the drawing stays put.  A long list
     used to make the PAGE tall, and scrolling down to its tail took the
     drawing off the top with it.  Sticky under the shell's header, with the
     same ceiling the rail uses, so the list's tail scrolls inside the panel
     column instead.  Stacked (narrow), it goes back to flowing. */
  aside { position: sticky; top: calc(var(--stickh) + 8px);
          max-height: calc(100vh - var(--stickh) - 22px);
          overflow-y: auto; overscroll-behavior: contain; }
  /* The panels keep their height inside the capped column - a flex child
     shrinks to fit by default, and a panel that shrinks clips its rows
     (overflow hidden) instead of pushing the column into scrolling. */
  aside > .panel { flex: 0 0 auto; }
  @media (max-width: 1100px) {
    aside { position: static; max-height: none; overflow: visible; }
  }
  .sheet { --fitH: calc(100vh - 190px);
           position: relative; flex: 1 1 auto; min-width: 0; display: flex;
           max-height: var(--fitH);
           overflow: auto; overscroll-behavior: contain; }
  .sheet.pannable { cursor: grab; }
  .sheet.panning  { cursor: grabbing; }
  /* flex-shrink 0, or a flex item told to be wider than its container just shrinks
     back to fit and there is never any overflow to scroll.  And margin:auto for
     the centring rather than justify-content: a CENTRED flex item wider than its
     box is clipped at the leading edge and cannot be scrolled back into view
     (the standard overflowing-flexbox trap).  margin:auto centres while there is
     room and stops centring when there is not - BUT only while it is the only
     thing centring.  An auto margin is treated as zero once the free space goes
     negative, at which point justify-content takes the alignment back and the
     trap is sprung again, so the .sheet above must not set it. */
  .page { position: relative; flex: 0 0 auto; margin: auto;
          /* the fit width, unchanged, times the zoom the reader drives.  min()
             reproduces the old sizing exactly at --vz:1 — whichever of the
             column's width or the window's height binds first. */
          width: calc(min(100%, var(--fitH) * var(--ar, 1.294) * var(--zoom, .8))
                      * var(--vz, 1));
          background: var(--panel); border: 1px solid var(--line);
          border-radius: 4px; overflow: hidden; }

  /* ===== Zoom controls in the header ===== */
  .zoombar { display: flex; align-items: center; gap: 4px; }
  /* The PDF download: there, but quiet - it is the drawing's second shape,
     not the page's purpose. */
  .pdfdl { color: var(--muted); font-size: 11px; letter-spacing: .6px; }
  .pdfdl:hover { color: var(--ink); }
  .zoombar .zpct { min-width: 44px; text-align: right; font-size: 11px;
                   color: var(--muted); font-variant-numeric: tabular-nums; }
  /* A drag that starts on a raster is answered by the browser's own "drag this
     image out" gesture, which steals the pointer stream mid-pan and leaves the
     sheet stuck under a ghost thumbnail; -webkit-user-drag kills it at the source
     and user-select stops the same drag selecting the sheet instead.  The embed
     is worse - a PDF plugin swallows every event over itself, which left only the
     thin margin round the page draggable - and it has nothing to click: the
     hotspots are in the overlay above it, so it gives up its events entirely. */
  .page embed, .page img { position: absolute; inset: 0; width: 100%; height: 100%;
                           border: 0; display: block; user-select: none;
                           -webkit-user-select: none; -webkit-user-drag: none; }
  .page embed { pointer-events: none; }
  .overlay { position: absolute; inset: 0; pointer-events: none;
             transform-origin: top left; }
  .hot { position: absolute; pointer-events: auto; cursor: pointer;
         border: 1.5px solid transparent; border-radius: 3px;
         transition: background .12s, border-color .12s; }
  .hot.balloon { border-radius: 50%; }
  .hot:hover { background: rgba(229, 138, 0, .34); border-color: var(--hot); }
  .hot.incart { background: rgba(26, 127, 69, .2); border-color: var(--ok); }
  .hot.active { background: rgba(225, 6, 0, .26); border-color: var(--red); }
  .hot.hover { background: rgba(229, 138, 0, .34); border-color: var(--hot); }
  .bomtable tr.hover td { background: rgba(229, 138, 0, .16); }
  .overlay img.sil.hover { opacity: 1; }
  /* `&` is this section's own root - the class the Align toggle sets */
  &.show-hot .hot.balloon { border-color: rgba(229, 138, 0, .8); }
  &.show-hot .hot.row { border-color: rgba(225, 6, 0, .5); }

  /* ===== Zones =====
     On MAIN, one polygon per section sheet - where that section sits in the
     machine.  On a section sheet, one tinted copy of each top-level component's
     own linework.  Neither takes pointer events: the page under them does the
     hit-testing (a polygon test, a mask lookup), so a balloon on top of a zone
     is still the balloon, and a transparent guard in front of a section swallows
     nothing.  Quiet until hovered; the Zones button shows them all. */
  .overlay svg.zones { position: absolute; inset: 0; width: 100%; height: 100%;
                       pointer-events: none; overflow: visible; }
  .overlay svg.zones polygon { fill: var(--zc); fill-opacity: 0; stroke: var(--zc);
                               stroke-opacity: 0; stroke-width: .25;
                               transition: fill-opacity .1s, stroke-opacity .1s; }
  .overlay svg.zones polygon.on { fill-opacity: .2; stroke-opacity: 1; stroke-width: .4; }
  .overlay svg.zones polygon.dim { fill-opacity: .04; stroke-opacity: .35; }
  &.show-zones .overlay svg.zones polygon { fill-opacity: .08; stroke-opacity: .6; }
  &.show-zones .overlay svg.zones polygon.on { fill-opacity: .2; stroke-opacity: 1; }
  .overlay img.sil, .overlay img.silref { right: auto; bottom: auto; pointer-events: none;
                                          opacity: 0; transition: opacity .1s; }
  .overlay img.sil.on { opacity: 1; }
  /* Held: the silhouette of the row whose card is open, lit from the parts
     list or the BOM table rather than from under the pointer.  A glow as well
     as full opacity, so it reads as chosen and not merely hovered. */
  .overlay img.sil.held { opacity: 1;
                          filter: drop-shadow(0 0 1.5px var(--red)) drop-shadow(0 0 5px rgba(225, 6, 0, .45)); }
  &.show-zones .overlay img.sil { opacity: .35; }
  &.show-zones .overlay img.sil.on { opacity: 1; }
  .overlay img.silref { inset: 0; width: 100%; height: 100%; }
  &.show-ref .overlay img.silref { opacity: 1; }

  /* The chip a silhouette shows while hovered: what you are pointing at, in one
     line, the way a BOM row's highlight tells you which row you are on.  The
     click opens the part card proper. */
  .zchip { position: fixed; z-index: 19; max-width: 340px; padding: 5px 9px;
           background: var(--ink); color: #fff; border-radius: 4px; font-size: 12px;
           line-height: 1.3; box-shadow: 0 6px 18px rgba(0,0,0,.25); pointer-events: none;
           white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .zchip b { font-family: var(--display); font-size: 13px; letter-spacing: .5px;
             margin-right: 6px; }
  .zchip .more { opacity: .7; margin-left: 6px; }

  /* The zone card: what a hover on MAIN names.  Same family as the part card,
     lighter - a section has nothing to add to a cart.  Pointer-transparent
     until pinned so it never sits between the pointer and the drawing. */
  .zcard { position: fixed; z-index: 19; width: 300px; background: var(--panel);
           border: 1px solid var(--line); border-top: 3px solid var(--zc, var(--red));
           border-radius: 4px; box-shadow: 0 12px 34px rgba(0,0,0,.22);
           padding: 9px 11px; pointer-events: none; }
  .zcard.pinned { pointer-events: auto; }
  .zcard .zhead { display: flex; align-items: baseline; gap: 8px; }
  .zcard .ztitle { font-family: var(--display); font-size: 17px; letter-spacing: .6px;
                   flex: 1; min-width: 0; overflow-wrap: anywhere; }
  .zcard .zclose { display: none; }
  .zcard.pinned .zclose { display: inline-flex; }
  .zcard .zsub { font-size: 11px; color: var(--muted); text-transform: uppercase;
                 letter-spacing: .5px; margin: 2px 0 8px; }
  .zcard .zbody { display: flex; gap: 10px; align-items: flex-start; }
  .zcard .zbody img { width: 130px; height: 96px; object-fit: contain; background: #fff;
                      border: 1px solid var(--line); border-radius: 3px; flex: 0 0 auto; }
  .zcard .ztext { font-size: 12.5px; line-height: 1.35; min-width: 0; }
  .zcard .zhint { font-size: 11px; color: var(--muted); margin-top: 4px; }
  .zcard ul { margin: 6px 0 0; padding-left: 16px; font-size: 11.5px;
              font-family: ui-monospace, Consolas, monospace; max-height: 120px; overflow: auto; }
  .zcard .zalso { margin-top: 8px; border-top: 1px solid var(--line); padding-top: 6px;
                  font-size: 11px; color: var(--muted); }
  .zcard .zalso div { display: flex; align-items: center; gap: 6px; padding: 3px 0;
                      font-size: 12.5px; color: var(--ink); cursor: pointer; }
  .zcard .zalso div:hover { color: var(--red); }
  .zcard .zalso i { width: 10px; height: 10px; border-radius: 2px; display: inline-block;
                    flex: 0 0 auto; }

  /* ===== Part card - what a click opens ===== */
  .card { position: fixed; z-index: 20; width: 310px; background: var(--panel);
          border: 1px solid var(--line); border-top: 3px solid var(--red);
          border-radius: 4px; box-shadow: 0 12px 34px rgba(0,0,0,.22); padding: 0; }
  .card .chead { display: flex; align-items: center; gap: 9px; padding: 9px 11px;
                 border-bottom: 1px solid var(--line); }
  .card .cballoon { flex: 0 0 auto; width: 26px; height: 26px; border-radius: 50%;
                    border: 2px solid var(--red); color: var(--red);
                    font-family: var(--display); font-size: 14px; display: flex;
                    align-items: center; justify-content: center; }
  .card .cpart { font-family: var(--display); font-size: 19px; letter-spacing: .8px;
                 flex: 1; min-width: 0; overflow-wrap: anywhere; }
  .card .cbody { padding: 10px 11px 0; }
  /* The model's saved preview.  Fixed height, not auto: the card is measured
     and placed the instant it opens, and a box that grows when the image
     finishes decoding would leave it hanging off its hotspot.  White behind it
     because that is what SolidWorks renders a preview against - contain, so a
     tall part is letterboxed rather than cropped. */
  .card .cthumb { display: block; width: 100%; height: 158px; object-fit: contain;
                  background: #fff; border: 1px solid var(--line);
                  border-radius: 3px; margin-bottom: 9px; }
  .card .cdesc { font-size: 13px; line-height: 1.35; margin-bottom: 6px; }
  .card .cmeta { font-size: 11px; color: var(--muted); text-transform: uppercase;
                 letter-spacing: .5px; margin-bottom: 10px; }
  /* The price sits directly above the stepper, which is the order the eye wants
     it in: what it is, what it costs, how many. */
  .card .cprice { display: flex; align-items: baseline; gap: 7px; font-size: 11px;
                  color: var(--muted); text-transform: uppercase; letter-spacing: .5px;
                  border-top: 1px solid var(--line); padding-top: 8px; margin-bottom: 10px; }
  .card .cprice b { font-family: var(--display); font-size: 18px; letter-spacing: .5px;
                    color: var(--ink); text-transform: none; }
  .card .cfoot { display: flex; align-items: center; gap: 7px; padding: 0 11px 11px; }
  .stepper { display: flex; align-items: center; border: 1px solid var(--line);
             border-radius: 4px; overflow: hidden; }
  .stepper button { border: 0; border-radius: 0; padding: 5px 10px; font-size: 15px;
                    background: var(--panel2); }
  .stepper input { width: 48px; border: 0; padding: 5px 4px; text-align: center;
                   font: inherit; background: var(--panel); color: var(--ink);
                   -moz-appearance: textfield; }
  .stepper input::-webkit-outer-spin-button,
  .stepper input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
  .card .cadd { flex: 1; }

  /* ===== Parts list + cart ===== */
  tbody tr { cursor: pointer; }
  tbody tr:hover, tbody tr.active { background: var(--red-soft); }
  /* Smaller than the shared 12.5px table, with tighter cells, so a narrow
     panel wraps descriptions less.  The cart keeps the shared size. */
  .bomtable { font-size: 11px; }
  .bomtable th { font-size: 9.5px; letter-spacing: .4px; }
  .bomtable th, .bomtable td { padding: 4px 6px; }
  /* The panel is 440px and overflow:hidden, so a wide BOM lost its last
     column(s) off the right edge.  Scroll sideways instead of clipping. */
  #bom { overflow-x: auto; }
  /* The drawing-wide total beside a row's own quantity: "1 (5)". */
  .bomtable .tot { color: var(--muted); margin-left: .35em; white-space: nowrap;
                   font-variant-numeric: tabular-nums; }
  td.sheetcol { color: var(--muted); font-size: 11.5px; white-space: nowrap; }
  th.pricecol, td.pricecol { text-align: right; white-space: nowrap; }
  .bomtable th:first-child, .bomtable td:first-child { width: 1%; padding-right: 6px; }
  td.pricecol.cfq, .cprice b.cfq { color: var(--muted); font-size: 11px; letter-spacing: .3px; cursor: help; }
  th.vercol, td.vercol { text-align: center; white-space: nowrap; }
  .bomtable .ver { color: var(--muted); margin-left: .35em; white-space: nowrap;
                   font-variant-numeric: tabular-nums; }
  .cart-empty { padding: 14px 12px; color: var(--muted); font-size: 12.5px; }
  .qty { width: 70px; padding: 3px 5px; font: inherit; text-align: right;
         border: 1px solid var(--line); border-radius: 3px; background: var(--panel2);
         color: var(--ink); }
  .conv { font-size: 10px; color: var(--muted); white-space: nowrap; margin-top: 2px; }
  .cart-foot { display: flex; gap: 8px; align-items: center; padding: 10px 12px;
               border-top: 1px solid var(--line); background: var(--shade); }
  /* Who is asking - optional, remembered across drawings, and written into the
     RFQ's header so the parts desk gets a complete request (see quoteBody). */
  .rfq-who { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; padding: 8px 12px 4px;
             border-bottom: 1px solid var(--line); }
  .rfq-who .lbl { grid-column: 1 / -1; font-size: 11px; color: var(--muted); }
  .rfq-who input { min-width: 0; font: inherit; font-size: 12px; padding: 4px 6px;
                   border: 1px solid var(--line); border-radius: 3px; background: var(--panel2);
                   color: var(--ink); }
  .rfq-who input.wide { grid-column: 1 / -1; }

  /* ===== All parts =====
     The rest of the module: every part under the sheet's reference assembly,
     the ones the curated list leaves out.  Shut until wanted; while it is open
     a row lights its part on the sheet and the sheet lights its row.  Shut, the
     reference assembly is the ghost it always was. */
  /* The table is the parts list's own (.bomtable, built by bomHead) - same
     columns, same type, same price column - so the two read as one list in
     two halves.  A pointer on the sheet lights a row the same way too. */
  #allparts { overflow-x: auto; }
  .overlay img.ap { right: auto; bottom: auto; pointer-events: none; opacity: 0;
                    transition: opacity .1s; }
  .overlay img.ap.on, .overlay img.ap.hover { opacity: 1; }
  .overlay img.ap.held { opacity: 1;
                         filter: drop-shadow(0 0 1.5px var(--red)) drop-shadow(0 0 5px rgba(225, 6, 0, .45)); }
  /* A part the REF layer could not outline: its instance box, on the spot. */
  .overlay .apbox { position: absolute; pointer-events: none; box-sizing: border-box; opacity: 0;
                    border: 2px solid var(--zc, var(--hot)); border-radius: 2px;
                    background: rgba(229, 138, 0, .12); transition: opacity .1s; }
  .overlay .apbox.on, .overlay .apbox.hover { opacity: 1; }
  .overlay .apbox.held { opacity: 1; box-shadow: 0 0 0 2px rgba(225, 6, 0, .35); }
  .overlay .apbox.dim { opacity: .3; }

  .align { display: none; gap: 8px; align-items: center; font-size: 11px;
           color: var(--muted); }
  .align.on { display: flex; }
  .align input[type=range] { width: 84px; vertical-align: middle; }
"""

# The card and the header controls are planted OUTSIDE the section's own column
# by the script below (the card is fixed-position; the Request Quote button and
# the Align sliders belong in the chrome).  They start here so the section owns
# every id it uses.
SECTION_HTML = r"""
<div class="sheet" id="sheets"></div>

<aside>
  <section class="panel collapsed" id="cart-panel">
    <h2 class="clickable" id="cart-head">
      <span class="chev">&#9662;</span> Cart
      <span class="sub" id="cart-sub">empty</span>
      <span class="spacer"></span>
      <button class="icon" id="cart-clear">Clear</button>
    </h2>
    <div class="panel-body">
      <div class="rfq-who" id="rfq-who">
        <span class="lbl">For the quote (optional - remembered on this PC)</span>
        <input id="rfq-company" placeholder="Company" autocomplete="organization">
        <input id="rfq-contact" placeholder="Your name &middot; email or phone" autocomplete="name">
        <input id="rfq-po" placeholder="PO #">
        <input id="rfq-need" type="date" title="Need by">
        <input id="rfq-ship" class="wide" placeholder="Ship: Ground / 2-Day / Overnight / Pick up">
      </div>
      <div id="cart-body"></div>
      <div class="cart-foot">
        <button id="cart-copy" disabled>Copy list</button>
        <span class="spacer"></span>
        <button class="primary" id="btn-quote-2" disabled>Request Quote</button>
      </div>
    </div>
  </section>

  <section class="panel" id="bom-panel">
    <h2 class="clickable" id="bom-head">
      <span class="chev">&#9662;</span> <span id="bom-title">Critical parts</span>
      <span class="spacer"></span>
      <span class="sub" id="bom-sheet"></span>
    </h2>
    <div class="panel-body"><div id="bom"></div></div>
  </section>

  <!-- The rest of the module: every part under the sheet's reference assembly,
       which the curated list above leaves out on purpose.  Hidden outright when
       the manifest carried none (an older macro). -->
  <section class="panel collapsed" id="all-panel" hidden>
    <h2 class="clickable" id="all-head">
      <span class="chev">&#9662;</span> <span id="all-title">All parts</span>
      <span class="spacer"></span>
      <span class="sub" id="all-sub"></span>
    </h2>
    <div class="panel-body"><div id="allparts"></div></div>
  </section>
</aside>

<div id="header-bits" hidden>
  <!-- Ctrl+wheel is the primary gesture; these are here because a trackpad, a
       touchscreen and a mouse with no wheel all exist, and because a control that
       shows the current zoom is how you discover the gesture is available. NOT
       .align-only - unlike the nudge sliders, zoom applies to every viewer mode. -->
  <div class="zoombar" id="zoombar">
    <button class="icon" id="z-out" title="Zoom out (Ctrl + wheel down)">&#8722;</button>
    <button class="icon" id="z-fit" title="Fit the sheet (Ctrl + 0)">Fit</button>
    <button class="icon" id="z-in" title="Zoom in (Ctrl + wheel up)">&#43;</button>
    <span class="zpct" id="z-pct">100%</span>
  </div>
  <button id="toggle-align" class="align-only">Align&hellip;</button>
  <div class="align align-only" id="align-panel">
    <label>x <input type="range" id="ax" min="-6" max="6" step="0.05" value="0"></label>
    <label>y <input type="range" id="ay" min="-6" max="6" step="0.05" value="0"></label>
    <label>zoom <input type="range" id="az" min="90" max="110" step="0.05" value="100"></label>
    <button id="align-reset">Reset</button>
  </div>
  <!-- The drawing itself, every sheet in one PDF: the export the page was
       rastered from, which the publish copies beside the HTML.  Hidden when
       that copy is not there (see DATA.pdf_file). -->
  <button class="icon pdfdl" id="btn-pdf" hidden
          title="Download every sheet as one PDF">PDF</button>
  <button class="primary" id="btn-quote" disabled>Request Quote<span class="count"
          id="cart-count">0</span></button>
</div>

<div class="zchip" id="zchip" hidden></div>

<div class="zcard" id="zcard" hidden>
  <div class="zhead">
    <span class="ztitle" id="z-title"></span>
    <button class="icon zclose" id="z-close" title="Close">&#10005;</button>
  </div>
  <div class="zsub" id="z-sub"></div>
  <div class="zbody">
    <img id="z-img" alt="" hidden>
    <div class="ztext"><div id="z-text"></div><div class="zhint" id="z-hint"></div></div>
  </div>
  <ul id="z-list" hidden></ul>
  <div class="zalso" id="z-also" hidden></div>
</div>

<div class="card" id="card" hidden>
  <div class="chead">
    <span class="cballoon" id="c-item"></span>
    <span class="cpart" id="c-part"></span>
    <button class="icon" id="c-close" title="Close">&#10005;</button>
  </div>
  <div class="cbody">
    <img class="cthumb" id="c-thumb" alt="" hidden>
    <div class="cdesc" id="c-desc"></div>
    <div class="cmeta" id="c-meta"></div>
    <div class="cprice" id="c-price" hidden></div>
  </div>
  <div class="cfoot">
    <div class="stepper">
      <button id="c-minus" title="Less">&#8722;</button>
      <input type="number" id="c-qty" min="1" step="1" value="1">
      <button id="c-plus" title="More">&#43;</button>
    </div>
    <button class="primary cadd" id="c-add">Add</button>
  </div>
</div>
"""

# Body of function (ROOT, DATA, SHELL).  $ and $$ come from SHELL and resolve
# inside this section's own root against its own namespace - which is why two of
# these can share a page with the manual's other sections and see none of them.
SECTION_JS = r"""
  var SHEETS = DATA.sheets;
  var THUMBS = DATA.thumbs || {};
  var toast = SHELL.toast, store = SHELL.store, load = SHELL.load, el = SHELL.el;
  var CART_KEY = "drawing-cart:" + DATA.drawing;
  var ALIGN_KEY = "drawing-align:" + DATA.drawing;
  var OPEN_KEY = "drawing-cart-open:" + DATA.drawing;
  var active = 0;

  /* The Align sliders only earn their place when the BROWSER is rendering the
     PDF - a rasterised sheet is aligned exactly, by construction, and there is
     nothing to nudge.  Hidden here rather than in CSS because these controls
     move into the header, away from the class that would have keyed it. */
  var browserPdf = DATA.viewer === "pdf";

  /* The header controls belong in the chrome, not in this section's column, so
     they are lifted into the header slot - which the shell tags with this
     section's own root class, so their CSS still reaches them, and hides along
     with the section.  The card is NOT moved: it is position:fixed, so it
     already floats over everything, and leaving it inside the root means it
     disappears when this section does. */
  var bits = $("header-bits");
  while (bits.firstChild) {
    var child = bits.firstChild;
    var cls = child.classList;
    var keep = !cls || !cls.contains("align-only") || browserPdf;
    if (keep) {
      SHELL.headerSlot(child);
    } else {
      bits.removeChild(child);
    }
  }
  bits.remove();

  /* Every sheet as one PDF.  The publish copies the exported PDF beside the
     HTML and names it in DATA.pdf_file, relative to the page, so the link
     works from the share and from a copy of the folder alike.  No copy, no
     button. */
  var pdfBtn = $("btn-pdf");
  if (pdfBtn && DATA.pdf_file) {
    pdfBtn.hidden = false;
    pdfBtn.addEventListener("click", function () {
      var a = document.createElement("a");
      a.href = DATA.pdf_file;
      a.download = DATA.pdf_file.split("/").pop();
      document.body.appendChild(a);
      a.click();
      a.remove();
    });
  }

  /* ---- quantities -------------------------------------------------------
     Most are piece counts.  Cut-to-length items are a measurement instead, so
     they keep their decimals and get the conversion shown beside them - the
     raw figure is always what travels, the conversion is only ever a hint. */
  var LEN = DATA.length_units || {};

  function isLength(q) {
    if (!LEN.factor) return false;
    return (q % 1 !== 0) || (LEN.min && q >= LEN.min);
  }
  function num(q) {
    return String(Math.round(q * 10000) / 10000);
  }
  function qtyText(q) {
    if (!isLength(q)) return num(q);
    return num(q) + " " + (LEN.from || "") +
           " (" + (q / LEN.factor).toFixed(2) + " " + (LEN.to || "") + ")";
  }

  /* ---- prices ------------------------------------------------------------
     A hotspot carries `price` only when the publish costed its row, and the
     multiplier was already applied on the way in - there is no cost figure in
     this page to work backwards from.

     `on` is decided by looking rather than by a flag: a drawing where nothing
     priced should show no price column and no card line at all, and that is
     not something the config can know in advance. */
  var PRICE = DATA.price || {};
  var priced = SHEETS.some(function (s) {
    return s.hotspots.some(function (h) { return typeof h.price === "number"; });
  });

  /* The PDM version, the same way: a column only when some row carries one.
     The text is the macro's, shown as written - not parsed, not padded. */
  var VER_LABEL = DATA.pdm_version_label || "";
  var versioned = !!VER_LABEL && SHEETS.some(function (s) {
    return s.hotspots.some(function (h) { return !!h.pdm_ver; }) ||
           (s.allparts || []).some(function (z) { return !!z.pdm_ver; });
  });

  /* Grouped by thousands and always two decimals.  Not toLocaleString: this
     page is read on whatever machine happens to open it, and a browser set to a
     European locale would render 1234.5 as "1.234,50" behind a "$". */
  function money(v) {
    return (PRICE.currency || "$") +
           v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  /* A cart line belongs to the SHEET it came from - that is how the RFQ is
     grouped, and item 3 on one sheet is not item 3 on another.

     Within a sheet the line is the PART, not the balloon: a part called out in
     four places is one thing to buy, and the four balloons plus its BOM row all
     have to answer to the same key so that clicking any of them lights up the
     rest and the cart holds one line.

     But a part number is not always one part.  The same number appears on a
     BOM twice whenever it is ordered in two CONFIGURATIONS - the same weldment
     left- and right-hand, the same extrusion at two cut lengths - and those are
     two different things to buy, with their own quantities.  Keyed on the part
     number alone they collapsed into one: clicking either lit up both, and Add
     to Cart wrote the second over the first, so the quote asked for one of them
     and silently dropped the other.  The configuration is therefore part of the
     key whenever the BOM gives one.  (Everything downstream was already built
     for this - the card prints "Config X", the cart grows a CONFIG column when
     any line has one, the RFQ mails it - it just could never happen.) */
  function baseKey(h) {
    var part = (h.part || "").trim().toUpperCase();
    var cfg = (h.config || "").trim().toUpperCase();
    /* No part number to group on - fall back to the balloon, which is at least
       unique on the sheet.  Appended only when there IS a configuration, so a
       drawing without that column keys exactly as it always did and a cart
       saved by an older build still matches. */
    if (!part) return "S:" + h.sheet + ":I:" + h.item;
    return "S:" + h.sheet + ":P:" + part + (cfg ? ":C:" + cfg : "");
  }

  /* The same trap without a CONFIG column: two BOM rows, one part number, and
     nothing on the sheet to tell them apart but their item numbers.  Merging
     those is the safe reading of a real duplicate, but it is the WRONG reading
     of two genuine callouts, and only the drawing knows which it is.  So the
     item number is added to the key for exactly those part numbers that appear
     under more than one item on a sheet - the ambiguous ones - and every other
     part keeps the short key it had.  Computed once, over the BOM rows only:
     balloons repeat their row's item and would report a false collision. */
  var AMBIGUOUS = (function () {
    var firstItem = {}, ambiguous = {};
    (SHEETS || []).forEach(function (s) {
      (s.hotspots || []).forEach(function (h) {
        if (h.kind !== "row") return;
        var k = baseKey(h), item = String(h.item);
        if (!(k in firstItem)) firstItem[k] = item;
        else if (firstItem[k] !== item) ambiguous[k] = true;
      });
    });
    return ambiguous;
  })();

  /* ---- hover, both ways (2026-10-08) ------------------------------------
     A BOM row under the pointer lights its balloon(s) on the sheet and the
     component's silhouette; a balloon under the pointer lights its row.  The
     click still opens the card - this is the glance before the click. */
  function rowHover(h, on) {
    $$('.hot[data-key="' + CSS.escape(keyOf(h)) + '"]').forEach(function (n) { n.classList.toggle("hover", on); });
    var rec = ZONES[h.sheet];
    if (rec) rec.sils.forEach(function (z) {
      if (String(z.item) === String(h.item) && rec.els[zoneId(z)]) rec.els[zoneId(z)].classList.toggle("hover", on);
    });
  }
  function hotHover(h, on) {
    $$('tr[data-key="' + CSS.escape(keyOf(h)) + '"]').forEach(function (n) { n.classList.toggle("hover", on); });
  }
  function keyOf(h) {
    var k = baseKey(h);
    return AMBIGUOUS[k] ? k + ":I:" + h.item : k;
  }

  // everything clickable that stands for a part, on the drawing or in the lists
  var CLICKABLE = ".hot, .bomtable tbody tr";

  // ---- the drawing -------------------------------------------------------
  var pdfUrl = DATA.pdf_href || "";
  if (DATA.pdf_base64) {
    var bin = atob(DATA.pdf_base64);
    var buf = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    pdfUrl = URL.createObjectURL(new Blob([buf], { type: "application/pdf" }));
  }

  var host = $("sheets");
  host.style.setProperty("--zoom", DATA.page_zoom || 0.8);
  /* Declared BEFORE the pages are built: buildZones runs inside that loop, and
     a var assigned further down is still undefined here. */
  var SVG_NS = "http://www.w3.org/2000/svg";
  var ZONES = {};                                  // sheet idx -> { sections, sils, els }
  SHEETS.forEach(function (s, idx) {
    var page = el("div", "page");
    page.id = NS + "page-" + idx;
    page.style.aspectRatio = s.w + " / " + s.h;
    page.style.setProperty("--ar", s.w / s.h);
    if (idx !== 0) page.hidden = true;

    if (s.image) {
      var img = document.createElement("img");
      img.src = "data:" + s.image_mime + ";base64," + s.image;
      img.alt = s.label;
      img.draggable = false;      // the attribute too: -webkit-user-drag is not universal
      page.appendChild(img);
    } else {
      var em = document.createElement("embed");
      em.type = "application/pdf";
      em.src = pdfUrl + "#page=" + (s.page + 1) +
               "&toolbar=0&navpanes=0&scrollbar=0&statusbar=0&view=Fit";
      page.appendChild(em);
    }

    var overlay = el("div", "overlay");
    overlay.id = NS + "overlay-" + idx;
    s.hotspots.forEach(function (h) {
      var box = el("div", "hot " + h.kind);
      box.style.left = (h.x * 100) + "%";
      box.style.top = (h.y * 100) + "%";
      box.style.width = (h.w * 100) + "%";
      box.style.height = (h.h * 100) + "%";
      box.title = h.message;
      box.dataset.key = keyOf(h);
      box.addEventListener("click", function (ev) { openCard(h, ev.currentTarget); });
      box.addEventListener("mouseenter", function () { hotHover(h, true); });
      box.addEventListener("mouseleave", function () { hotHover(h, false); });
      overlay.appendChild(box);
    });
    buildZones(s, idx, page, overlay);
    page.appendChild(overlay);
    host.appendChild(page);
  });

  /* ---- zones ------------------------------------------------------------
     MAIN carries one polygon per section sheet; a section sheet carries one
     tinted silhouette per top-level component.  Both are drawn into the overlay
     so the Align nudge moves them with the hotspots, and both are
     pointer-transparent: the PAGE listens, and decides what is under the
     pointer itself - a point-in-polygon test for the sections, a lookup in the
     silhouette's own fattened linework for the components.  Smallest first, so
     a bracket inside a frame is the bracket; everything else under the pointer
     is listed on the card, one click away. */
  function buildZones(s, idx, page, overlay) {
    var sections = (s.sections || []).slice().sort(function (a, b) { return a.area - b.area; });
    var sils = (s.silhouettes || []).slice().sort(function (a, b) { return a.area - b.area; });
    var aps = s.allparts || [];
    if (!sections.length && !sils.length && !aps.length) return;
    var rec = { sections: sections, sils: sils, aps: [], els: {} };
    ZONES[idx] = rec;
    if (sections.length) {
      var svg = document.createElementNS(SVG_NS, "svg");
      svg.setAttribute("class", "zones");
      svg.setAttribute("viewBox", "0 0 100 100");
      svg.setAttribute("preserveAspectRatio", "none");
      var palette = ["#d81b60", "#1e88e5", "#43a047", "#fb8c00", "#8e24aa", "#00acc1",
                     "#6d4c41", "#c0ca33", "#e53935", "#3949ab", "#00897b", "#f4511e"];
      sections.forEach(function (z, i) { z.colour = palette[z.sheet % palette.length]; });
      sections.slice().reverse().forEach(function (z) {       // biggest first, so small sit on top
        var pl = document.createElementNS(SVG_NS, "polygon");
        pl.setAttribute("points", z.poly.map(function (q) {
          return (q[0] * 100) + "," + (q[1] * 100); }).join(" "));
        pl.style.setProperty("--zc", z.colour);
        svg.appendChild(pl);
        rec.els["s" + z.sheet] = pl;
      });
      overlay.appendChild(svg);
    }
    sils.forEach(function (z) {
      var im = document.createElement("img");
      im.className = "sil";
      im.src = z.overlay;
      im.alt = "";
      im.draggable = false;
      im.style.left = (z.x * 100) + "%";
      im.style.top = (z.y * 100) + "%";
      im.style.width = (z.w * 100) + "%";
      im.style.height = (z.h * 100) + "%";
      overlay.appendChild(im);
      rec.els[z.key] = im;
      var m = new Image();
      m.onload = function () {
        var c = document.createElement("canvas");
        c.width = m.width; c.height = m.height;
        var g = c.getContext("2d");
        g.drawImage(m, 0, 0);
        z.hit = g.getImageData(0, 0, c.width, c.height).data;
        z.hw = c.width; z.hh = c.height;
      };
      m.src = z.mask;
    });
    /* The reference assembly's parts: one ZONE per place an instance was
       found - its own outline off the REF export's layer, a block's outline
       shared by every instance, or a small instance box - and nothing where
       none was.  A zone carries its entry's fields plus `ekey`, the entry it
       belongs to, so a row lights every instance and an instance lights its
       row.  Drawn and hit-tested like the silhouettes, but only while the
       All parts panel is open (see zoneStack). */
    rec.aps = [];
    aps.forEach(function (e) {
      (e.places || []).forEach(function (p, i) {
        var z = Object.assign({}, e, p, { ap: true, ekey: e.key, key: e.key + "~" + i });
        delete z.places;
        rec.aps.push(z);
        var n;
        if (z.overlay) {
          n = document.createElement("img");
          n.className = "ap";
          n.src = z.overlay;
          n.alt = "";
          n.draggable = false;
          var am = new Image();
          am.onload = function () {
            var c = document.createElement("canvas");
            c.width = am.width; c.height = am.height;
            var g = c.getContext("2d");
            g.drawImage(am, 0, 0);
            z.hit = g.getImageData(0, 0, c.width, c.height).data;
            z.hw = c.width; z.hh = c.height;
          };
          am.src = z.mask;
        } else {
          n = el("div", "apbox");
          n.style.setProperty("--zc", z.colour || "");
        }
        n.style.left = (z.x * 100) + "%";
        n.style.top = (z.y * 100) + "%";
        n.style.width = (z.w * 100) + "%";
        n.style.height = (z.h * 100) + "%";
        overlay.appendChild(n);
        rec.els[z.key] = n;
      });
    });
    rec.aps.sort(function (a, b) { return a.area - b.area; });
    if (s.ref_overlay) {
      var ref = document.createElement("img");
      ref.className = "silref";
      ref.src = s.ref_overlay;
      ref.alt = "";
      ref.draggable = false;
      overlay.appendChild(ref);
    }
    page.addEventListener("mousemove", function (ev) { zoneMove(ev, idx, page); });
    page.addEventListener("mouseleave", function () {
      if (!zpinned) { zoneLight(idx, []); zhover = ""; }
      zcardHide(); chipHide();
    });
    page.addEventListener("click", function (ev) { zoneClick(ev, idx, page); });
  }

  function zoneId(z) { return z.key !== undefined ? z.key : "s" + z.sheet; }

  function inPoly(poly, x, y) {
    var hit = false;
    for (var i = 0, j = poly.length - 1; i < poly.length; j = i++) {
      var xi = poly[i][0], yi = poly[i][1], xj = poly[j][0], yj = poly[j][1];
      if ((yi > y) !== (yj > y) && x < (xj - xi) * (y - yi) / (yj - yi) + xi) hit = !hit;
    }
    return hit;
  }

  /* Everything under the pointer on this page, smallest first. */
  function zoneStack(ev, idx, page) {
    var rec = ZONES[idx];
    if (!rec) return [];
    if (ev.target && ev.target.closest && ev.target.closest(".hot")) return [];
    var r = page.getBoundingClientRect();
    var fx = (ev.clientX - r.left) / r.width, fy = (ev.clientY - r.top) / r.height;
    var out = rec.sections.filter(function (z) { return inPoly(z.poly, fx, fy); });
    rec.sils.forEach(function (z) {
      if (!z.hit) return;
      var mx = Math.floor((fx - z.x) / z.w * z.hw), my = Math.floor((fy - z.y) / z.h * z.hh);
      if (mx < 0 || my < 0 || mx >= z.hw || my >= z.hh) return;
      if (z.hit[(my * z.hw + mx) * 4] > 0) out.push(z);
    });
    /* The reference assembly's parts count only while their panel is open:
       shut, the ghost is a ghost and the pointer passes through it. */
    if (allOpen()) (rec.aps || []).forEach(function (z) {
      if (!z.w || !z.h) return;
      if (fx < z.x || fy < z.y || fx > z.x + z.w || fy > z.y + z.h) return;
      if (z.overlay) {
        if (!z.hit) return;
        var ax = Math.floor((fx - z.x) / z.w * z.hw), ay = Math.floor((fy - z.y) / z.h * z.hh);
        if (ax < 0 || ay < 0 || ax >= z.hw || ay >= z.hh) return;
        if (z.hit[(ay * z.hw + ax) * 4] > 0) out.push(z);
      } else out.push(z);
    });
    return out;
  }

  /* The All parts row for whatever reference part the pointer is on. */
  function apRows(stack) {
    var top = stack.length && stack[0].ap ? stack[0].ekey : null;
    $$("tr[data-ap]").forEach(function (tr) {
      tr.classList.toggle("hover", tr.dataset.ap === top);
    });
  }
  /* Every zone of one All parts entry - all its instances. */
  function apZones(idx, ekey) {
    var rec = ZONES[idx];
    return rec ? (rec.aps || []).filter(function (z) { return z.ekey === ekey; }) : [];
  }

  function zoneLight(idx, stack) {
    var rec = ZONES[idx];
    if (!rec) return;
    var top = stack.length ? zoneId(stack[0]) : null, inStack = {};
    stack.forEach(function (z) { inStack[zoneId(z)] = true; });
    var held = (silHeld && silHeld.idx === idx) ? silHeld.keys : {};
    Object.keys(rec.els).forEach(function (id) {
      var on = id === top || !!held[id];
      rec.els[id].classList.toggle("on", on);
      rec.els[id].classList.toggle("held", !!held[id]);
      rec.els[id].classList.toggle("dim", top !== null && !on && !!inStack[id]);
    });
    apRows(stack);
  }

  /* The BOM row a silhouette stands for, if the sheet's parts list has it: the
     component's file name is the part number, as a rule. */
  function rowFor(idx, z) {
    var hit = null;
    if (z.item) {
      SHEETS[idx].hotspots.forEach(function (h) {
        if (!hit && h.kind === "row" && String(h.item) === String(z.item)) hit = h;
      });
      if (hit) return hit;
    }
    var want = {};
    (z.parts || []).forEach(function (p) { want[p.replace(/[^A-Z0-9]/g, "")] = true; });
    SHEETS[idx].hotspots.forEach(function (h) {
      if (h.kind !== "row") return;
      var part = (h.part || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (part && want[part] && !hit) hit = h;
    });
    return hit;
  }

  /* The silhouettes lit for the row whose card is open - clicked in the BOM
     table on the sheet, in the parts list beside it, or in the cart - and
     HELD lit whatever the pointer does on the way, until the card closes.
     The same lookup rowFor makes, run the other way: every silhouette whose
     row this is. */
  var silHeld = null;                                   // { idx, keys: {key: true} }

  function silsFor(idx, h) {
    var rec = ZONES[idx];
    if (!rec || !rec.sils) return [];
    return rec.sils.filter(function (z) { return rowFor(idx, z) === h; });
  }
  function silHoldKeys(idx, keys) {
    silRelease();
    if (!keys.length) return;
    silHeld = { idx: idx, keys: {} };
    keys.forEach(function (k) { silHeld.keys[k] = true; });
    zoneLight(idx, []);
  }
  function silHold(h) {
    var idx = typeof h.sheet === "number" ? h.sheet : active;
    silHoldKeys(idx, silsFor(idx, h).map(function (z) { return z.key; }));
  }
  function silRelease() {
    if (!silHeld) return;
    var idx = silHeld.idx;
    silHeld = null;
    zoneLight(idx, []);
  }
  /* Bring the held silhouette into view - a row picked from the list beside
     the sheet may stand for a part scrolled off the other side of it. */
  function silReveal() {
    if (!silHeld) return;
    var rec = ZONES[silHeld.idx], first = Object.keys(silHeld.keys)[0];
    var im = rec && rec.els[first];
    if (im && im.scrollIntoView) im.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  /* What a silhouette opens: the row's own card when the parts list has the
     part, and otherwise the SAME card for the component itself - its file name
     as the part number, its instances as the description, quantity one - so a
     weldment or a guard that never got a BOM line can still be picked and
     quoted from the drawing. */
  function hotspotFor(idx, z) {
    var row = rowFor(idx, z);
    if (row) return row;
    var desc = z.ap ? (z.desc || "")
             : (z.n_names > 1 ? z.n_names + " instances: " : "") + (z.names || []).join(", ");
    /* A part that is on no BOM row still prices when the entry itself carries a
       price (the manual puts one on every ALL PARTS entry it can estimate). */
    return { kind: "sil", sheet: idx, item: "\u2013", part: z.label, config: z.ap ? (z.config || "") : "",
             description: desc || (z.files || []).join(", ") ||
                          (z.ap ? "(part of the reference assembly)" : "(top-level component)"),
             qty: z.ap ? (z.count || 1) : 1, message: "", url: "", thumb: z.thumb || "",
             pdm_ver: z.ap ? (z.pdm_ver || "") : "",
             price: typeof z.price === "number" ? z.price : undefined, price_basis: z.price_basis || "" };
  }

  var zchip = $("zchip");
  function chipShow(idx, stack, ev) {
    var z = stack[0], row = rowFor(idx, z);
    zchip.textContent = "";
    zchip.appendChild(el("b", "", row ? (row.part || z.label) : z.label));
    zchip.appendChild(document.createTextNode(
      row ? ("Item " + row.item + (row.description ? "  \u00b7  " + row.description : ""))
          : z.ap ? ((z.count || 1) + " in the reference assembly" + (z.desc ? "  \u00b7  " + z.desc : ""))
          : (z.n_names ? z.n_names + " instance(s), not on the parts list" : "not on the parts list")));
    if (stack.length > 1) zchip.appendChild(el("span", "more", "+" + (stack.length - 1) + " under"));
    zchip.hidden = false;
    chipPlace(ev);
  }
  function chipPlace(ev) {
    var w = zchip.offsetWidth, hgt = zchip.offsetHeight;
    var left = ev.clientX + 16, top = ev.clientY + 18;
    if (left + w > window.innerWidth - 8) left = ev.clientX - w - 16;
    if (top + hgt > window.innerHeight - 8) top = ev.clientY - hgt - 12;
    zchip.style.left = Math.max(8, left) + "px";
    zchip.style.top = Math.max(8, top) + "px";
  }
  function chipHide() { zchip.hidden = true; }

  var zcard = $("zcard"), zpinned = false, zhover = "";

  function zcardHide() { if (!zpinned) { zcard.hidden = true; zhover = ""; } }
  function zcardClose() {
    zpinned = false;
    zcard.classList.remove("pinned");
    zcard.hidden = true;
    zhover = "";
    Object.keys(ZONES).forEach(function (i) { zoneLight(+i, []); });
  }
  $("z-close").onclick = zcardClose;

  function zcardFill(idx, stack) {
    var z = stack[0], isSection = z.key === undefined;
    zcard.style.setProperty("--zc", z.colour || "");
    var img = $("z-img"), text = $("z-text"), hint = $("z-hint");
    var list = $("z-list"), also = $("z-also");
    list.textContent = ""; also.textContent = "";
    list.hidden = true; also.hidden = true;
    if (isSection) {
      $("z-title").textContent = sheetLabel(z.sheet);
      $("z-sub").textContent = z.assembly;
      var sh = SHEETS[z.sheet];
      if (sh && sh.image) { img.src = "data:" + sh.image_mime + ";base64," + sh.image; img.hidden = false; }
      else { img.removeAttribute("src"); img.hidden = true; }
      text.textContent = "Sheet " + (z.sheet + 1) + " of " + SHEETS.length;
      hint.textContent = stack.length > 1 ? "Click to choose  ·  Shift+click opens" : "Click to open the sheet";
    } else {
      var row = rowFor(idx, z);
      $("z-title").textContent = z.label;
      $("z-sub").textContent = (z.files || []).join(", ") || ("view block " + z.key);
      var pic = row && row.thumb ? THUMBS[row.thumb] : (z.thumb ? THUMBS[z.thumb] || "" : "");
      if (pic) { img.src = pic; img.hidden = false; }
      else { img.removeAttribute("src"); img.hidden = true; }
      text.textContent = row ? ("Item " + row.item + (row.description ? "  ·  " + row.description : ""))
                       : z.ap ? ((z.count || 1) + " in the reference assembly" + (z.desc ? "  ·  " + z.desc : ""))
                             : (z.n_names ? z.n_names + " instance(s) at the top of this assembly"
                                          : "Not a top-level component the manifest names");
      hint.textContent = "Click for the part card";
      if (z.names && z.names.length > 1) {
        z.names.forEach(function (n) { list.appendChild(el("li", "", n)); });
        if (z.n_names > z.names.length) list.appendChild(el("li", "", "… " + (z.n_names - z.names.length) + " more"));
        list.hidden = false;
      }
    }
    if (stack.length > 1) {
      also.appendChild(document.createTextNode("Also under the pointer:"));
      stack.slice(1).forEach(function (q) {
        var d = el("div", "");
        var sw = el("i", "");
        sw.style.background = q.colour || "";
        d.appendChild(sw);
        d.appendChild(document.createTextNode(q.key === undefined ? sheetLabel(q.sheet) : q.label));
        d.addEventListener("click", function (ev) {
          ev.stopPropagation();
          zoneOpen(idx, q, [q].concat(stack.filter(function (o) { return o !== q; })), d, true);
        });
        also.appendChild(d);
      });
      also.hidden = false;
    }
  }

  function zcardPlace(ev) {
    var w = zcard.offsetWidth, hgt = zcard.offsetHeight;
    var left = ev.clientX + 18, top = ev.clientY + 18;
    if (left + w > window.innerWidth - 8) left = ev.clientX - w - 18;
    if (top + hgt > window.innerHeight - 8) top = ev.clientY - hgt - 18;
    zcard.style.left = Math.max(8, left) + "px";
    zcard.style.top = Math.max(8, top) + "px";
  }

  function zoneMove(ev, idx, page) {
    if (zpinned || pan) return;
    var stack = zoneStack(ev, idx, page);
    var key = stack.map(zoneId).join(",");
    if (key !== zhover) {
      zhover = key;
      zoneLight(idx, stack);
      zcard.hidden = true;
      chipHide();
      if (stack.length) {
        if (stack[0].key === undefined) { zcardFill(idx, stack); zcard.hidden = false; }
        else chipShow(idx, stack, ev);
      }
    }
    if (stack.length) {
      if (stack[0].key === undefined) zcardPlace(ev); else chipPlace(ev);
    }
  }

  /* What a click on a zone does.  A section opens its sheet.  A component
     opens its PART CARD when the parts list has it - the same card a balloon
     opens, anchored on the silhouette - and otherwise pins the zone card. */
  function zoneOpen(idx, z, stack, anchor, force) {
    if (z.key === undefined) {
      if (stack.length === 1 || force) { zcardClose(); showSheet(z.sheet); return; }
      zpinned = true;
      zcardFill(idx, stack);
      zcard.classList.add("pinned");
      zcard.hidden = false;
      zoneLight(idx, stack);
      return;
    }
    zcardClose();
    chipHide();
    openCard(hotspotFor(idx, z), ZONES[idx].els[z.key] || anchor);
    silHoldKeys(idx, [z.key]);
  }

  /* Silhouettes overlap - a bracket on a frame, a motor on a mount.  The first
     click takes the smallest; clicking the same spot again takes the next one
     under it, round and round, so nothing is unreachable. */
  var zlastKey = "", zlastAt = 0;

  function zoneClick(ev, idx, page) {
    if (zpinned) {
      if (!zcard.contains(ev.target)) zcardClose();
      return;
    }
    var stack = zoneStack(ev, idx, page);
    if (!stack.length) { zlastKey = ""; return; }
    if (stack[0].key === undefined) {                      // MAIN: a section
      if (stack.length > 1 && !ev.shiftKey) {
        zpinned = true;
        zcardFill(idx, stack);
        zcard.classList.add("pinned");
        zcard.hidden = false;
        zcardPlace(ev);
        zoneLight(idx, stack);
        return;
      }
      zoneOpen(idx, stack[0], stack, ev.target, ev.shiftKey);
      return;
    }
    var key = idx + ":" + stack.map(zoneId).join(",");
    zlastAt = (key === zlastKey) ? (zlastAt + 1) % stack.length : 0;
    zlastKey = key;
    zoneOpen(idx, stack[zlastAt], stack, ev.target, false);
  }


  /* ---- zoom and pan -----------------------------------------------------
     Ctrl+wheel zooms about the POINTER, which is the only anchor that feels
     right: the detail you are pointing at is the detail you want to keep, and
     zooming about the centre of the box walks it off screen in two notches.
     Plain wheel is left alone, so once the sheet is bigger than the window it
     still scrolls the way every other scrollable thing does.
     Zoom is applied by writing --vz, which widens .page (see the CSS).  The
     .sheet viewport then overflows and scrolls, which means panning is REAL
     scrolling - drag, wheel, scrollbars, arrow keys and Home/End all work and
     none of them had to be written here.  It also means the hotspots, which are
     positioned in percentages inside .page, scale with the sheet and stay
     exactly where the Align nudge put them, and a rasterised sheet is re-drawn
     at the new size by the browser rather than being resampled from a bitmap.
     The floor is 1 = fitted: the fit already shows the whole sheet, and smaller
     than that is just a smaller picture of the same thing. */
  var VZ_MIN = 1, VZ_MAX = 8, vz = 1, swallowClick = false, pan = null;
  var zpct = $("z-pct");

  function applyZoom() {
    host.style.setProperty("--vz", vz);
    if (zpct) zpct.textContent = Math.round(vz * 100) + "%";
    /* The grab cursor claims that dragging will do something, so it is set from
       whether there is actually anything to scroll rather than from vz > 1 - a
       tall narrow sheet on a wide window can be zoomed and still have slack. */
    var can = host.scrollWidth > host.clientWidth + 1 ||
              host.scrollHeight > host.clientHeight + 1;
    host.classList.toggle("pannable", can);
    if (!can) host.classList.remove("panning");
  }

  /* Zoom to `next`, keeping the content under viewport point (cx, cy) put.  The
     scroll offset that achieves this scales with the zoom, so the anchor is
     converted to content units BEFORE the change and read back after.  Null
     coordinates mean the centre of the viewport, which is what the buttons want. */
  function zoomTo(next, cx, cy) {
    next = Math.min(VZ_MAX, Math.max(VZ_MIN, next));
    if (next === vz) return;
    var r = host.getBoundingClientRect();
    var ax = (cx == null) ? r.width / 2 : cx - r.left;
    var ay = (cy == null) ? r.height / 2 : cy - r.top;
    var px = host.scrollLeft + ax, py = host.scrollTop + ay;
    var k = next / vz;
    vz = next;
    applyZoom();
    host.scrollLeft = px * k - ax;
    host.scrollTop = py * k - ay;
  }

  function zoomFit() {
    vz = 1;
    applyZoom();
    host.scrollLeft = 0;
    host.scrollTop = 0;
  }

  /* passive:false is the whole point - a wheel listener is passive by default on
     the platforms that matter, and a passive one cannot preventDefault, so the
     browser would zoom the entire page (chrome, rail, panels and all) instead of
     the sheet.  deltaY arrives in three different units depending on device and
     OS, so only its SIGN is trusted; the magnitude is clamped and the step is
     ours, which keeps a trackpad's flood of tiny deltas and a mouse's chunky
     notches feeling like the same gesture. */
  host.addEventListener("wheel", function (ev) {
    if (!ev.ctrlKey) return;
    ev.preventDefault();
    var mag = Math.abs(ev.deltaY) || 1;
    if (ev.deltaMode === 1) mag *= 16;         // lines
    else if (ev.deltaMode === 2) mag *= 400;   // pages
    var factor = Math.exp((ev.deltaY > 0 ? -1 : 1) * Math.min(mag, 200) * 0.0025);
    zoomTo(vz * factor, ev.clientX, ev.clientY);
  }, { passive: false });

  /* Drag to pan - only while there is something to pan, and never starting on a
     hotspot: those are the entire point of the page and have to stay clickable.
     A drag that actually moved swallows the click it would otherwise finish with,
     so panning off a hotspot does not also open that part's card. */
  host.addEventListener("pointerdown", function (ev) {
    if (ev.button !== 0 || !host.classList.contains("pannable")) return;
    if (ev.target && ev.target.closest && ev.target.closest(".hot")) return;
    /* This is what turns the press into a pan and not a selection or an image
       drag - the default action of a pointerdown on a picture is to start one. */
    ev.preventDefault();
    pan = { x: ev.clientX, y: ev.clientY,
            sl: host.scrollLeft, st: host.scrollTop, moved: false };
    host.classList.add("panning");
    if (host.setPointerCapture) {
      try { host.setPointerCapture(ev.pointerId); } catch (e) {}
    }
  });
  host.addEventListener("pointermove", function (ev) {
    if (!pan) return;
    var dx = ev.clientX - pan.x, dy = ev.clientY - pan.y;
    if (!pan.moved && Math.abs(dx) + Math.abs(dy) > 3) pan.moved = true;
    host.scrollLeft = pan.sl - dx;
    host.scrollTop = pan.st - dy;
  });
  function endPan(ev) {
    if (!pan) return;
    if (pan.moved) swallowClick = true;
    pan = null;
    host.classList.remove("panning");
    if (host.releasePointerCapture && ev.pointerId != null) {
      try { host.releasePointerCapture(ev.pointerId); } catch (e) {}
    }
  }
  host.addEventListener("pointerup", endPan);
  host.addEventListener("pointercancel", endPan);
  /* Capture phase: the hotspot's own click handler is on the hotspot, so this has
     to run before the event reaches it. */
  host.addEventListener("click", function (ev) {
    if (!swallowClick) return;
    swallowClick = false;
    ev.stopPropagation();
    ev.preventDefault();
  }, true);

  var zin = $("z-in"), zout = $("z-out"), zfit = $("z-fit");
  if (zin) zin.onclick = function () { zoomTo(vz * 1.25, null, null); };
  if (zout) zout.onclick = function () { zoomTo(vz / 1.25, null, null); };
  if (zfit) zfit.onclick = function () { zoomFit(); };

  /* Ctrl+0 pairs with the wheel gesture.  It is only taken while this section is
     the open one AND the sheet is actually zoomed - at 100% there is nothing to
     reset, so the browser keeps its own shortcut rather than having it silently
     eaten by a page that had no use for it. */
  document.addEventListener("keydown", function (ev) {
    if (!ev.ctrlKey || ev.altKey || ev.metaKey) return;
    if (ev.key !== "0" && ev.key !== "Insert") return;
    if (vz === 1 || !ROOT.classList.contains("on")) return;
    ev.preventDefault();
    zoomFit();
  });

  /* The fit size is derived from the window, so what can be panned changes when
     the window does even though the zoom has not. */
  window.addEventListener("resize", applyZoom);
  applyZoom();

  /* ---- sheet list -------------------------------------------------------
     The sheets ARE this section's rail rows, so they go in the rail the shell
     owns rather than a switcher of their own.  On a one-section page that rail
     is the whole left column, which is the sheet switcher this page has always
     had; inside the manual the same rows nest under Illustrated Parts. */
  var tabs = SHELL.railSubs();
  if (tabs && SHEETS.length > 1) {
    /* CLAIM the rows, do not add more.  The shell has already drawn one rail
       row per sheet off this section's `nav` - that is the whole point of
       declaring them there, so the manual can paint the rail before this
       script runs.  Appending a second set put every sheet in the rail twice.
       Row idx is sheet idx; only fill in what the shell could not know, which
       is the hotspot count and the switch behaviour.  The shell's own click
       handler is an onclick property, so addEventListener stacks on top of it
       rather than replacing it - nested in the manual, one click still opens
       the section AND lands on the sheet. */
    SHEETS.forEach(function (s, idx) {
      var b = tabs.children[idx];
      if (!b) {                      // no nav row to claim - make one
        b = document.createElement("button");
        tabs.appendChild(b);
      }
      b.id = NS + "tab-" + idx;
      b.title = s.hotspots.length + " hotspot(s) on this sheet";
      b.textContent = "";            // the shell wrote a bare label; respan it
      b.appendChild(el("span", "lbl", s.label));
      var n = el("span", "n");
      n.id = NS + "tabn-" + idx;
      b.appendChild(n);
      b.addEventListener("click", function () { showSheet(idx); });
    });
  }

  function showSheet(idx) {
    if (idx < 0 || idx >= SHEETS.length) return;
    active = idx;
    closeCard();
    if (typeof zcardClose === "function") zcardClose();
    clearSelection();
    SHEETS.forEach(function (s, i) {
      var p = $("page-" + i);
      if (p) p.hidden = i !== idx;
      var t = $("tab-" + i);
      if (t) t.classList.toggle("on", i === idx);
    });
    renderBom();
    renderAll();
    applyAlign();
    /* A new sheet is a new drawing to read, and inheriting the last one's zoom and
       scroll position lands you in the middle of it with no idea where you are. */
    zoomFit();
  }

  /* How many of each part the WHOLE drawing calls for: the same part number
     and configuration summed over every sheet's parts list, shown grey in
     parentheses beside the row's own quantity - a bushing called out once
     here and four times elsewhere reads "1 (5)".  Keyed on part + config
     because a different configuration is a different thing to order. */
  var TOTALS = null;
  function partKey(h) {
    /* by file where the macro found one: parts machined from one stocked
       blank share a StockedPartNumber and are still different parts */
    return (h.file || h.part || "").toUpperCase().replace(/[^A-Z0-9]/g, "") + "|" +
           (h.config || "").toUpperCase().trim();
  }
  function totals() {
    if (TOTALS) return TOTALS;
    TOTALS = {};
    SHEETS.forEach(function (s, i) {
      (s.hotspots || []).forEach(function (h) {
        if (h.kind !== "row" || !h.part) return;
        var k = partKey(h), t = TOTALS[k] || (TOTALS[k] = { qty: 0, where: [] });
        var q = Number(h.qty);
        t.qty += isFinite(q) ? q : 1;
        t.where.push(sheetLabel(i) + " item " + h.item + ": " + qtyText(h.qty || 1));
      });
    });
    return TOTALS;
  }

  /* The parts list's header, from the sheet's own BOM columns: "#" for the
     item column, "(ver)" on CONFIG when the rows carry a vault version (its
     own column when there is no CONFIG to ride in), and the price column
     appended - that one did not come off the sheet, it came out of the
     costing database, and wears the payload's label and note.  Shared by the
     parts list and the All parts table, so the two are one table in two
     halves.  Returns the <thead> and which columns are which. */
  /* For a sheet with no BOM of its own: the columns a -M sheet's BOM carries. */
  var BOM_DEFAULT_COLS = ["ITEM", "DESCRIPTION", "PART NUMBER", "CONFIG", "QTY."];
  function bomHead(cols) {
    var showPrice = priced && PRICE.in_bom !== false;
    var thead = document.createElement("thead");
    var htr = document.createElement("tr");
    cols.forEach(function (c, i) {
      var label = /^item\s*#?\.?$/i.test(c.trim()) ? "#" : c;
      htr.appendChild(el("th", i === 0 ? "num" : "", label));
    });
    var ccol = -1;
    cols.forEach(function (c, i) { if (ccol < 0 && /^config/i.test(c.trim())) ccol = i; });
    var VER_TIP = VER_LABEL + ": the vault version of this part's file when the drawing was published";
    if (versioned && ccol >= 0) {        // the version rides in this cell: say so in the head
      var cth = htr.children[ccol];
      if (cth) { cth.textContent = cth.textContent.replace(/\s*\(ver\)$/i, "") + "(ver)"; cth.title = VER_TIP; }
    }
    if (versioned && ccol < 0) {
      var vth = el("th", "vercol", VER_LABEL);
      vth.title = VER_TIP;
      htr.appendChild(vth);
    }
    if (showPrice) {
      var pth = el("th", "pricecol", PRICE.label || "Est. Price");
      if (PRICE.note) pth.title = PRICE.note;
      htr.appendChild(pth);
    }
    thead.appendChild(htr);
    /* The quantity column, for the drawing-wide total; the last column when
       the BOM has no column named for it. */
    var qcol = -1;
    cols.forEach(function (c, i) { if (qcol < 0 && /qty|quant/i.test(c)) qcol = i; });
    if (qcol < 0) qcol = cols.length - 1;
    return { thead: thead, ccol: ccol, qcol: qcol, showPrice: showPrice, VER_TIP: VER_TIP };
  }

  // ---- parts list of the active sheet ------------------------------------
  function renderBom() {
    var s = SHEETS[active];
    var bom = $("bom");
    bom.textContent = "";
    $("bom-sheet").textContent = SHEETS.length > 1 ? s.label : "";

    var rows = s.hotspots.filter(function (h) { return h.kind === "row"; });
    if (!rows.length) {
      bom.appendChild(el("p", "cart-empty", "No BOM table was found on this sheet."));
      return;
    }
    var head = bomHead(s.columns);
    var showPrice = head.showPrice, ccol = head.ccol, qcol = head.qcol, VER_TIP = head.VER_TIP;
    var table = el("table", "bomtable");
    table.appendChild(head.thead);
    var all = totals();

    var tbody = document.createElement("tbody");
    rows.forEach(function (h) {
      var tr = document.createElement("tr");
      tr.dataset.key = keyOf(h);
      tr.title = "Click for details";
      s.columns.forEach(function (c, i) {
        var td = el("td", i === 0 ? "num" : "", (h.values && h.values[c]) || "");
        if (i === qcol && h.part && all[partKey(h)]) {
          var t = all[partKey(h)];
          var tot = el("span", "tot", "(" + num(t.qty) + ")");
          tot.title = num(t.qty) + " on the whole drawing: " + t.where.join("; ");
          td.appendChild(tot);
        }
        if (i === ccol && versioned && h.pdm_ver) {
          var ver = el("span", "ver", "(" + h.pdm_ver + ")");
          ver.title = VER_TIP;
          td.appendChild(ver);
        }
        tr.appendChild(td);
      });
      if (versioned && ccol < 0) tr.appendChild(el("td", "vercol", h.pdm_ver || ""));
      /* Blank, not a dash and not a zero, on a row nothing could be found for.
         "$0.00" beside a part is a claim that it is free. */
      if (showPrice) {
        /* No price: the payload's "not available" word when it carries one
           ("Quote" - call for quote - in the manual), else blank. */
        var ptd = el("td", "pricecol",
                     typeof h.price === "number" ? money(h.price) : (PRICE.na || ""));
        if (typeof h.price === "number") { if (h.price_basis) ptd.title = "priced on " + h.price_basis; }
        else if (PRICE.na) { ptd.classList.add("cfq"); if (PRICE.na_note) ptd.title = PRICE.na_note; }
        tr.appendChild(ptd);
      }
      tr.addEventListener("click", function (ev) { openCard(h, ev.currentTarget); silReveal(); });
      tr.addEventListener("mouseenter", function () { rowHover(h, true); });
      tr.addEventListener("mouseleave", function () { rowHover(h, false); });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    bom.appendChild(table);
    markCart();
  }

  /* ---- all parts of the active sheet's reference assembly ---------------
     The parts list's own table, continued: the same columns off the sheet's
     BOM (#, Description, Part #, Config(ver), Qty, Est. Price), one row per
     part and configuration, the quantity being how many the module holds.
     The item number is filled only for a part the curated list has too.  A
     row opens the same card a balloon does - the BOM row's when the curated
     list has the part, else the part's own, with its picture - and holds its
     place on the sheet lit while the card is open. */
  function renderAll() {
    var s = SHEETS[active], box = $("allparts");
    if (!box) return;
    box.textContent = "";
    var aps = s.allparts || [];
    $("all-sub").textContent = aps.length
      ? (aps.length + (aps.length === 1 ? " part" : " parts") + (SHEETS.length > 1 ? "  \u00b7  " + s.label : ""))
      : (SHEETS.length > 1 ? "none on " + s.label : "none");
    if (!aps.length) {
      box.appendChild(el("p", "cart-empty", "This sheet's assembly carries no reference model."));
      return;
    }
    var cols = (s.columns && s.columns.length) ? s.columns : BOM_DEFAULT_COLS;
    var head = bomHead(cols);
    var table = el("table", "bomtable");
    table.appendChild(head.thead);
    /* What each of the sheet's columns asks for, by its heading. */
    function cell(c, z, h) {
      var t = c.trim();
      if (/^item/i.test(t)) return h.kind === "row" ? String(h.item) : "";
      if (/desc/i.test(t)) return z.desc || "";
      if (/part/i.test(t)) return z.pn || (z.parts && z.parts[0]) || z.label;
      if (/^config/i.test(t)) return z.config || "";
      if (/qty|quant/i.test(t)) return String(z.count || 1);
      return "";
    }
    var tbody = document.createElement("tbody");
    var idx = active;
    aps.forEach(function (z) {
      var h = hotspotFor(idx, z);
      var tr = document.createElement("tr");
      tr.dataset.ap = z.key;
      tr.dataset.key = keyOf(h);
      var zs = apZones(idx, z.key);
      tr.title = zs.length ? "Hover to find it on the sheet; click for details" : "Click for details";
      cols.forEach(function (c, i) {
        var td = el("td", i === 0 ? "num" : "", cell(c, z, h));
        if (i === head.qcol) td.title = (z.count || 1) + " in the reference assembly";
        if (i === head.ccol && versioned && h.pdm_ver) {
          var ver = el("span", "ver", "(" + h.pdm_ver + ")");
          ver.title = head.VER_TIP;
          td.appendChild(ver);
        }
        tr.appendChild(td);
      });
      if (versioned && head.ccol < 0) tr.appendChild(el("td", "vercol", h.pdm_ver || ""));
      if (head.showPrice) {
        var ptd = el("td", "pricecol",
                     typeof h.price === "number" ? money(h.price) : (PRICE.na || ""));
        if (typeof h.price === "number") { if (h.price_basis) ptd.title = "priced on " + h.price_basis; }
        else if (PRICE.na) { ptd.classList.add("cfq"); if (PRICE.na_note) ptd.title = PRICE.na_note; }
        tr.appendChild(ptd);
      }
      tr.addEventListener("click", function (ev) {
        openCard(h, ev.currentTarget);
        var keys = silsFor(idx, h).map(function (q) { return q.key; });
        apZones(idx, z.key).forEach(function (q) { keys.push(q.key); });
        silHoldKeys(idx, keys);
        silReveal();
      });
      tr.addEventListener("mouseenter", function () { apLight(idx, z, true); });
      tr.addEventListener("mouseleave", function () { apLight(idx, z, false); });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    box.appendChild(table);
    markCart();
  }
  function apLight(idx, z, on) {
    var rec = ZONES[idx];
    if (!rec) return;
    apZones(idx, z.key).forEach(function (q) {
      if (rec.els[q.key]) rec.els[q.key].classList.toggle("hover", on);
    });
  }

  /* ---- the part card ----------------------------------------------------
     A click SHOWS the part - what it is, which sheet calls for it, how many
     the BOM asks for - and adding to the cart is the second, deliberate act.
     The quantity starts at the called-out figure and can be changed right
     here, so the common case is one glance and one button. */
  var card = $("card");
  var cardFor = null;
  var cardAnchor = null;

  function openCard(h, anchor) {
    cardFor = h;
    cardAnchor = anchor;
    highlight(keyOf(h));
    silHold(h);

    $("c-item").textContent = h.item;
    $("c-part").textContent = h.part || ("Item " + h.item);
    $("c-desc").textContent = h.description || "(no description)";

    /* The model's own preview, when the macro exported one.  A part whose file
       was saved without a preview graphic simply has none: the box is removed
       rather than left empty, so the card closes up around the text.  The src
       is dropped as well as hidden - an <img src=""> re-requests the page. */
    var thumb = $("c-thumb");
    var picture = h.thumb ? THUMBS[h.thumb] : "";
    if (picture) {
      thumb.src = picture;
      thumb.alt = h.part || ("Item " + h.item);
      thumb.hidden = false;
    } else {
      thumb.removeAttribute("src");
      thumb.alt = "";
      thumb.hidden = true;
    }

    var meta = [];
    /* The configuration only: the vault version rides on the BOM row and in the
       quote (VER), not on the card. */
    if (h.config) meta.push("Config " + h.config);
    if (SHEETS.length > 1) meta.push(SHEETS[h.sheet].label);
    meta.push("BOM qty " + qtyText(h.qty || 1));
    $("c-meta").textContent = meta.join("  ·  ");

    /* The price, per each, when this row was costed.  Hidden rather than zeroed
       on a row that was not: a blank "Estimated Price:" reads as free. */
    var money_ = $("c-price");
    if (typeof h.price === "number") {
      money_.textContent = (PRICE.label || "Est. Price") + ":";
      money_.appendChild(el("b", "", money(h.price)));
      money_.title = h.price_basis ? ("priced on " + h.price_basis) : "";
      money_.hidden = false;
    } else if (PRICE.na) {
      money_.textContent = (PRICE.label || "Est. Price") + ":";
      money_.appendChild(el("b", "cfq", PRICE.na));
      money_.title = PRICE.na_note || "";
      money_.hidden = false;
    } else {
      money_.textContent = "";
      money_.removeAttribute("title");
      money_.hidden = true;
    }

    var inCart = cartIndex(keyOf(h));
    var q = $("c-qty");
    q.step = isLength(h.qty || 1) ? "any" : "1";
    q.value = num(inCart >= 0 ? cart[inCart].qty : (h.qty || 1));
    $("c-add").textContent = inCart >= 0 ? "Update" : "Add";

    card.hidden = false;
    position(anchor);
  }

  function position(anchor) {
    var r = anchor.getBoundingClientRect();
    var w = card.offsetWidth, hgt = card.offsetHeight;
    var left = r.right + 10;
    if (left + w > window.innerWidth - 8) left = r.left - w - 10;
    if (left < 8) left = Math.max(8, (window.innerWidth - w) / 2);
    var top = r.top + r.height / 2 - hgt / 2;
    top = Math.min(Math.max(8, top), window.innerHeight - hgt - 8);
    card.style.left = left + "px";
    card.style.top = top + "px";
  }

  function closeCard() {
    card.hidden = true;
    cardFor = null;
    cardAnchor = null;
    silRelease();
  }

  // The card is fixed-position but pinned to a hotspot, so it has to follow it
  // when the window moves under it rather than shutting in the reader's face.
  function reposition() {
    if (card.hidden || !cardAnchor) return;
    if (!cardAnchor.isConnected || !cardAnchor.offsetParent) { closeCard(); return; }
    position(cardAnchor);
  }
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);

  $("c-close").onclick = closeCard;
  $("c-minus").onclick = function () { bump(-1); };
  $("c-plus").onclick = function () { bump(1); };
  function bump(d) {
    var q = $("c-qty");
    q.value = num(Math.max(0.01, Math.round(((parseFloat(q.value) || 1) + d) * 10000) / 10000));
  }
  $("c-qty").addEventListener("keydown", function (ev) {
    if (ev.key === "Enter") $("c-add").click();
  });
  $("c-add").onclick = function () {
    if (!cardFor) return;
    var raw = parseFloat($("c-qty").value) || 0;
    var qty = Math.max(0.01, Math.round(raw * 10000) / 10000);
    setLine(cardFor, qty);
    toast("Item " + cardFor.item + " × " + qtyText(qty) + " in the cart");
    closeCard();
  };
  document.addEventListener("keydown", function (ev) {
    if (ev.key === "Escape") closeCard();
  });
  document.addEventListener("mousedown", function (ev) {
    if (card.hidden) return;
    if (card.contains(ev.target)) return;
    if (ev.target.closest && ev.target.closest(CLICKABLE)) return;
    closeCard();
  });

  // ---- cart --------------------------------------------------------------
  /* "S:" lines came off a sheet; "X:" lines were handed in by the page around us
     (the Technical Manual's Electrical BOM, spare kit, critical parts - see the
     cart API at the foot) and have no sheet: sheet -1 and a `src` saying where. */
  var cart = load(CART_KEY, []).filter(function (l) {
    return l && typeof l.key === "string" &&
           ((l.key.indexOf("S:") === 0 && typeof l.sheet === "number") ||
            l.key.indexOf("X:") === 0);
  });
  /* A saved line carries the fields the key is made of, so the key is RECOMPUTED
     rather than trusted: a cart parked in a browser before configurations were
     keyed would otherwise hold a stale key, match nothing on the sheet, and show
     neither its green in-cart mark nor its quantity when the part was clicked. */
  cart.forEach(function (l) { if (l.key.indexOf("X:") !== 0) l.key = keyOf(l); });
  var seenLine = {};              // two old lines can recompute to one key
  cart = cart.filter(function (l) {
    if (seenLine[l.key]) return false;
    seenLine[l.key] = true;
    return true;
  });

  function cartIndex(key) {
    for (var i = 0; i < cart.length; i++) if (cart[i].key === key) return i;
    return -1;
  }

  function setLine(h, qty) {
    var key = keyOf(h);
    var i = cartIndex(key);
    if (i < 0) {
      cart.push({ key: key, sheet: h.sheet, item: h.item, part: h.part || "",
                  config: h.config || "", description: h.description || "",
                  pdm_ver: h.pdm_ver || "", qty: qty });
    } else {
      cart[i].qty = qty;
    }
    renderCart();
  }

  function setQty(key, qty) {
    var i = cartIndex(key);
    if (i < 0) return;
    qty = Math.max(0, Math.round((Number(qty) || 0) * 10000) / 10000);
    if (qty === 0) cart.splice(i, 1); else cart[i].qty = qty;
    renderCart();
  }

  function sheetLabel(idx) { return SHEETS[idx] ? SHEETS[idx].label : "?"; }

  // CONFIG is blank on plenty of drawings; only give it a column when it earns one
  function hasConfig() {
    return cart.some(function (l) { return (l.config || "").trim(); });
  }

  function markCart() {
    $$(CLICKABLE).forEach(function (n) {
      if (n.dataset.key) n.classList.toggle("incart", cartIndex(n.dataset.key) >= 0);
    });
    SHEETS.forEach(function (s, i) {
      var n = $("tabn-" + i);
      if (!n) return;
      var c = cart.filter(function (l) { return l.sheet === i; }).length;
      n.textContent = c || "";
      n.style.display = c ? "" : "none";
    });
  }

  function renderCart() {
    store(CART_KEY, cart);
    notifyCart();

    /* The badge counts LINES, not quantities: adding 1407.35 mm of chain to
       2 bearings gives 1409.35 of nothing.  Pieces are only totalled over the
       items that are actually countable. */
    var pieces = cart.reduce(function (n, l) {
      return isLength(l.qty) ? n : n + l.qty;
    }, 0);
    var cuts = cart.filter(function (l) { return isLength(l.qty); }).length;

    $("cart-count").textContent = cart.length;
    var sub = [];
    if (cart.length) {
      sub.push(cart.length + (cart.length === 1 ? " line" : " lines"));
      if (pieces) sub.push(num(pieces) + " pcs");
      if (cuts) sub.push(cuts + (cuts === 1 ? " cut length" : " cut lengths"));
    }
    $("cart-sub").textContent = sub.join(" · ") || "empty";

    var empty = cart.length === 0;
    $("btn-quote").disabled = empty;
    $("btn-quote-2").disabled = empty;
    $("cart-copy").disabled = empty;

    var body = $("cart-body");
    body.textContent = "";

    if (empty) {
      body.appendChild(el("p", "cart-empty",
        "Click a balloon on the drawing, or a row in the parts list."));
      markCart();
      return;
    }

    var multi = SHEETS.length > 1;
    var withConfig = hasConfig();
    var t = document.createElement("table");
    t.innerHTML = "<thead><tr>" + (multi ? "<th>Sheet</th>" : "") +
                  '<th class="num">Item</th><th>Part #</th>' +
                  (withConfig ? "<th>Config</th>" : "") +
                  '<th>Description</th><th class="num">Qty</th><th></th>' +
                  "</tr></thead>";
    var tb = document.createElement("tbody");

    cart.forEach(function (line) {
      var tr = document.createElement("tr");
      tr.dataset.key = line.key;
      var fromManual = line.sheet < 0;
      tr.title = fromManual ? (line.src || "Added from the manual")
                            : "Show on " + sheetLabel(line.sheet);
      tr.addEventListener("click", function (ev) {
        if (ev.target.tagName === "INPUT" || ev.target.tagName === "BUTTON") return;
        if (fromManual) return;               // nothing on a sheet to show
        showSheet(line.sheet);
        highlight(line.key);
        var h = null;
        (SHEETS[line.sheet].hotspots || []).forEach(function (x) {
          if (!h && keyOf(x) === line.key) h = x; });
        if (h) { silHold(h); silReveal(); }
      });

      if (multi) tr.appendChild(el("td", "sheetcol",
                                   fromManual ? (line.src || "Manual") : sheetLabel(line.sheet)));
      tr.appendChild(el("td", "num", line.item));
      tr.appendChild(el("td", "", line.part || "—"));
      if (withConfig) tr.appendChild(el("td", "", line.config || "—"));
      tr.appendChild(el("td", "", line.description || "—"));

      var tdQty = el("td", "num");
      var input = document.createElement("input");
      input.type = "number";
      input.min = "0";
      input.step = isLength(line.qty) ? "any" : "1";
      input.className = "qty";
      input.value = num(line.qty);
      input.addEventListener("change", function () { setQty(line.key, input.value); });
      tdQty.appendChild(input);
      if (isLength(line.qty)) {
        tdQty.appendChild(el("div", "conv",
          (LEN.from || "") + " = " + (line.qty / LEN.factor).toFixed(2) +
          " " + (LEN.to || "")));
      }
      tr.appendChild(tdQty);

      var tdDel = el("td", "num");
      var del = el("button", "icon", "×");
      del.title = "Remove";
      del.addEventListener("click", function () { setQty(line.key, 0); });
      tdDel.appendChild(del);
      tr.appendChild(tdDel);

      tb.appendChild(tr);
    });
    t.appendChild(tb);
    body.appendChild(t);
    markCart();
  }

  function clearSelection() {
    $$(".active").forEach(function (n) { n.classList.remove("active"); });
  }

  function highlight(key) {
    clearSelection();
    $$('[data-key="' + CSS.escape(key) + '"]')
      .forEach(function (n) { n.classList.add("active"); });
  }

  // ---- cart panel, shut until wanted -------------------------------------
  var cartPanel = $("cart-panel");
  if (load(OPEN_KEY, false) === true) cartPanel.classList.remove("collapsed");
  $("cart-head").addEventListener("click", function (ev) {
    if (ev.target.id === NS + "cart-clear") return;
    var open = cartPanel.classList.toggle("collapsed") === false;
    store(OPEN_KEY, open);
  });
  // ---- the parts panels: Critical parts open, All parts shut until wanted --
  var LABELS = DATA.labels || {};
  if (LABELS.parts) $("bom-title").textContent = LABELS.parts;
  if (LABELS.allparts) $("all-title").textContent = LABELS.allparts;
  var BOM_KEY = "drawing-bom-open:" + DATA.drawing;
  var ALL_KEY = "drawing-all-open:" + DATA.drawing;
  var bomPanel = $("bom-panel");
  if (load(BOM_KEY, true) === false) bomPanel.classList.add("collapsed");
  $("bom-head").addEventListener("click", function () {
    store(BOM_KEY, bomPanel.classList.toggle("collapsed") === false);
  });
  var allPanel = $("all-panel");
  var HAS_ALL = SHEETS.some(function (s) { return (s.allparts || []).length > 0; });
  allPanel.hidden = !HAS_ALL;
  if (HAS_ALL && load(ALL_KEY, false) === true) allPanel.classList.remove("collapsed");
  $("all-head").addEventListener("click", function () {
    var open = allPanel.classList.toggle("collapsed") === false;
    store(ALL_KEY, open);
    /* Shut, the reference parts go back to being the ghost: nothing stays lit
       but what a card holds. */
    if (!open) {
      Object.keys(ZONES).forEach(function (i) { zoneLight(+i, []); });
      zhover = "";
      chipHide();
    }
  });
  function allOpen() { return HAS_ALL && !allPanel.classList.contains("collapsed"); }

  $("cart-clear").onclick = function (ev) {
    ev.stopPropagation();
    if (!cart.length) return;
    cart = [];
    clearSelection();
    closeCard();
    renderCart();
    toast("Cart cleared");
  };

  /* ---- RFQ email: #POLYPACK-PARTS-RFQ v1 ---------------------------------
     The body is BOTH the thing a person reads and the thing PolyApp imports
     (CO > PARTS QUOTE pastes it and prices it), so one format does both jobs
     rather than a readable list plus an opaque code block - long carts already
     press on the mailto: limit.  The grammar, kept in step with the parser
     under the same version tag:

         #POLYPACK-PARTS-RFQ v1          the block starts here; anything above
         KEY: value                      one header field per line, any order
         --- SHEET 1 ---                 a group (only on a multi-sheet drawing,
                                         or for lines the manual handed in)
         * QTY n  |  PART p  |  CONFIG c  |  ITEM i     one line per part;
           free-text description         fields are KEY value, "  |  " apart,
                                         any order, CONFIG/ITEM/FROM optional
         N line items
         #END PPQ1 CRC xxxx              CRC-16/CCITT-FALSE, see quoteCrc

     Deliberately NOT a space-padded table.  Mail clients render a plain-text
     body in a proportional font, so padding that lines up here lands ragged
     there, and worse, differently ragged on every drawing.  Labelling each
     field instead means nothing depends on character widths: it reads the
     same in Outlook, on a phone, and pasted into a purchase order - and the
     parser keys off the labels, so a reflowed or ">"-quoted copy still reads.
     Grouped under the sheet each line came from, because a sheet is a real
     division of the machine. */
  var RFQ_FORMAT = "v1";

  /* The footer's check covers the PART and QTY of every line, in order, and
     nothing else - so a customer's reflow, signature or ">" quoting does not
     trip it, while a hand-edited quantity does.  A mismatch is information,
     not an error: the import shows "edited by customer" and goes on.
     CRC-16/CCITT-FALSE (poly 0x1021, init 0xFFFF) over the UTF-8 bytes of
     "PART|QTY" lines joined by "\n", PART upper-cased, QTY as printed. */
  function quoteCrc(lines) {
    var text = lines.map(function (l) {
      return String(l.part || "").trim().toUpperCase() + "|" + qtyText(l.qty);
    }).join("\n");
    var bytes;
    try { bytes = unescape(encodeURIComponent(text)); } catch (e) { bytes = text; }
    var crc = 0xFFFF;
    for (var i = 0; i < bytes.length; i++) {
      crc ^= (bytes.charCodeAt(i) & 0xFF) << 8;
      for (var b = 0; b < 8; b++) {
        crc = (crc & 0x8000) ? ((crc << 1) ^ 0x1021) : (crc << 1);
        crc &= 0xFFFF;
      }
    }
    return ("0000" + crc.toString(16).toUpperCase()).slice(-4);
  }

  /* Who is asking.  Remembered under one key for every drawing and manual on
     this PC - a customer quoting from three drawings types their name once.
     Each filled field becomes a header line; blank ones are left out. */
  var WHO_KEY = "polypack-rfq-who";
  var WHO_FIELDS = [["company", "COMPANY"], ["contact", "CONTACT"], ["po", "PO"],
                    ["need", "NEED BY"], ["ship", "SHIP"]];
  var who = load(WHO_KEY, {}) || {};
  WHO_FIELDS.forEach(function (f) {
    var inp = $("rfq-" + f[0]);
    if (!inp) return;
    inp.value = who[f[0]] || "";
    inp.addEventListener("input", function () {
      who[f[0]] = inp.value.trim();
      store(WHO_KEY, who);
    });
  });

  function quoteLine(l, extra) {
    var bits = ["QTY " + qtyText(l.qty), "PART " + (l.part || "(none)")];
    if (l.config) bits.push("CONFIG " + l.config);
    /* The vault version of the part's file when the drawing was published - for
       the parts desk (it says which revision was quoted), never on the card. */
    if (l.pdm_ver) bits.push("VER " + l.pdm_ver);
    if (l.item) bits.push("ITEM " + l.item);
    if (extra) bits.push(extra);
    var out = ["* " + bits.join("  |  ")];
    // one description line, and never one that could pass for a field line
    if (l.description) out.push("  " + String(l.description).replace(/\s+/g, " ").trim());
    out.push("");
    return out;
  }

  function quoteBody() {
    var lines = [], ordered = [];
    if (DATA.quote.intro) lines.push(DATA.quote.intro, "");
    lines.push("#POLYPACK-PARTS-RFQ " + RFQ_FORMAT);
    if (DATA.serial) lines.push("SERIAL: " + DATA.serial);
    lines.push("DRAWING: " + DATA.drawing);
    // inside the Technical Manual the shell's key is "tm:<serial>" and the
    // page title names the book; on the standalone drawing page neither applies
    if (/^tm:/.test(SHELL.key || "") && document.title) {
      lines.push("MANUAL: " + document.title.replace(/\s+/g, " ").trim());
    }
    WHO_FIELDS.forEach(function (f) {
      if (who[f[0]]) lines.push(f[1] + ": " + who[f[0]]);
    });
    lines.push("");

    SHEETS.forEach(function (s, i) {
      var group = cart.filter(function (l) { return l.sheet === i; });
      if (!group.length) return;
      if (SHEETS.length > 1) lines.push("--- " + s.label.toUpperCase() + " ---");
      group.forEach(function (l) {
        ordered.push(l);
        lines.push.apply(lines, quoteLine(l, ""));
      });
    });
    // lines the manual handed in (its Electrical BOM, spare kit, critical parts)
    var ext = cart.filter(function (l) { return l.sheet < 0; });
    if (ext.length) {
      if (ext.length !== cart.length || SHEETS.length > 1) lines.push("--- FROM THE MANUAL ---");
      ext.forEach(function (l) {
        ordered.push(l);
        lines.push.apply(lines, quoteLine(l, l.src ? "FROM " + l.src : ""));
      });
    }

    lines.push(ordered.length === 1 ? "1 line item" : ordered.length + " line items");
    lines.push("#END PPQ1 CRC " + quoteCrc(ordered));
    return lines.join("\n").replace(/\n{3,}/g, "\n\n").replace(/\n+$/, "");
  }

  function copyText(text) {
    function fallback() {
      var ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand("copy"); toast("Parts list copied"); }
      catch (e) { toast("Could not copy - use Request Quote"); }
      document.body.removeChild(ta);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        function () { toast("Parts list copied"); }, fallback);
    } else {
      fallback();
    }
  }

  /* Straight into the mail client - the recipient and subject were settled
     when the page was built, so there is nothing left to ask.  A cart big
     enough to overflow a mailto: also lands on the clipboard, because a
     silently trimmed parts list is worse than no email. */
  function sendQuote() {
    if (!cart.length) return;
    var body = quoteBody();
    var subject = (DATA.quote.subject || "{SERIAL}-PARTS RFQ")
      .replace("{SERIAL}", DATA.serial).replace("{DRAWING}", DATA.drawing);

    if (encodeURIComponent(body).length > 1800) {
      copyText(body);
      toast("Long list - also copied, in case your mail client trims it");
    }
    var a = document.createElement("a");
    a.href = "mailto:" + encodeURIComponent(DATA.quote.email || "") +
             "?subject=" + encodeURIComponent(subject) +
             "&body=" + encodeURIComponent(body);
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  $("btn-quote").onclick = sendQuote;
  $("btn-quote-2").onclick = sendQuote;
  $("cart-copy").onclick = function () {
    if (cart.length) copyText(quoteBody());
  };

  /* ---- alignment nudge --------------------------------------------------
     Only reachable when the browser is drawing the PDF itself, so every one of
     these may legitimately be absent - the controls were dropped above rather
     than merely hidden.  applyAlign still runs (the stored transform is a real
     setting, and showSheet calls it either way); it simply has no sliders to
     write back to. */
  var align = load(ALIGN_KEY, { x: 0, y: 0, z: 100 });
  var ax = $("ax"), ay = $("ay"), az = $("az");

  function applyAlign() {
    var t = "translate(" + align.x + "%," + align.y + "%) scale(" +
            (align.z / 100) + ")";
    $$(".overlay").forEach(function (o) { o.style.transform = t; });
    if (ax) { ax.value = align.x; ay.value = align.y; az.value = align.z; }
    store(ALIGN_KEY, align);
  }
  if (ax) {
    ax.oninput = function () { align.x = +ax.value; applyAlign(); };
    ay.oninput = function () { align.y = +ay.value; applyAlign(); };
    az.oninput = function () { align.z = +az.value; applyAlign(); };
    $("align-reset").onclick = function () {
      align = { x: 0, y: 0, z: 100 };
      applyAlign();
    };
    var panel = $("align-panel");
    var toggleAlign = $("toggle-align");
    toggleAlign.onclick = function () {
      var on = panel.classList.toggle("on");
      toggleAlign.classList.toggle("on", on);
      ROOT.classList.toggle("show-hot", on);
    };
  }

  /* ---- deep link: ...#sheet=2&add=3 -------------------------------------
     The sheet comes first because an item number only means something on a
     given sheet.  A deep link opens the part card rather than adding blind. */
  function consumeHash() {
    var hash = location.hash || "";
    var ms = /[#&]sheet=(\d+)/.exec(hash);
    var ma = /[#&]add=([^&]*)/.exec(hash);
    if (!ms && !ma) return;
    var idx = ms ? Math.max(0, Math.min(SHEETS.length - 1, +ms[1] - 1)) : active;
    SHELL.show(SHELL.id);
    showSheet(idx);
    if (ma) {
      var item = decodeURIComponent(ma[1]);
      var hit = null;
      SHEETS[idx].hotspots.forEach(function (h) {
        if (h.item === item && (!hit || h.kind === "row")) hit = h;
      });
      if (hit) {
        var anchor = ROOT.querySelector(
          "#" + NS + "page-" + idx + ' .hot[data-key="' + CSS.escape(keyOf(hit)) + '"]');
        openCard(hit, anchor || $("page-" + idx));
      } else {
        toast("Item " + item + " is not on " + sheetLabel(idx));
      }
    }
    history.replaceState(null, "", location.pathname + location.search);
  }

  showSheet(0);
  applyAlign();
  renderCart();
  consumeHash();
  window.addEventListener("hashchange", consumeHash);

  /* ---- the cart, for the page around us (2026-10-02) ------------------------
     Folded into the Technical Manual, this drawing sits beside an Electrical BOM,
     a spare-parts kit and a critical-parts list, and the reader wants ONE request,
     not one per page.  The manual's own + buttons add through here; a line from
     outside has no sheet (sheet -1), a "X:" key and a `src` naming where it came
     from.  `watch` lets the manual's header badge follow this cart.  Returned
     from init, so POLYPACK.api(id).cart reaches it; nothing else about the page
     changes. */
  var cartWatchers = [];
  function notifyCart() {
    (cartWatchers || []).forEach(function (f) { try { f(cart.slice()); } catch (e) {} });
  }
  return {
    cart: {
      format: RFQ_FORMAT,
      lines: function () { return cart.slice(); },
      add: function (line) {
        var part = String((line && line.part) || "").trim();
        if (!part) return null;
        var key = "X:" + part + (line.config ? ":C:" + line.config : "");
        var i = cartIndex(key);
        var qty = Math.max(0, Number(line.qty) || 1);
        if (i < 0) {
          cart.push({ key: key, sheet: -1, item: "", part: part, config: line.config || "",
                      description: line.description || line.desc || "", qty: qty,
                      src: line.src || "" });
        } else {
          cart[i].qty = qty;
        }
        renderCart();
        return key;
      },
      has: function (part) {
        return cart.some(function (l) { return l.part === part; });
      },
      setQty: setQty,
      remove: function (key) { setQty(key, 0); },
      clear: function () { cart = []; renderCart(); },
      quote: sendQuote,
      body: quoteBody,
      watch: function (fn) { cartWatchers.push(fn); }
    }
  };
"""



def _encode_via_pillow(pix, fmt: str, quality: int) -> tuple[str, str] | None:
    """Encode a PyMuPDF pixmap with Pillow, which has codecs PyMuPDF lacks."""
    try:
        from PIL import Image
    except ImportError:
        return None

    image = Image.frombytes("RGBA" if pix.alpha else "RGB",
                            (pix.width, pix.height), pix.samples)
    if image.mode != "RGB":
        image = image.convert("RGB")

    if fmt in ("auto", "webp"):
        buffer = io.BytesIO()
        try:
            image.save(buffer, "WEBP", quality=quality, method=6)
            return "image/webp", base64.b64encode(buffer.getvalue()).decode("ascii")
        except Exception:
            if fmt == "webp":
                return None            # asked for it explicitly, so do not silently swap

    buffer = io.BytesIO()
    # 4:4:4 - chroma subsampling smears the red title-block linework
    image.save(buffer, "JPEG", quality=quality, optimize=True,
               progressive=True, subsampling=0)
    return "image/jpeg", base64.b64encode(buffer.getvalue()).decode("ascii")


def _rasterise(pdf_path: str, page_index: int, dpi: int,
               image_format: str = "auto", quality: int = 85
               ) -> tuple[str, str] | None:
    """Render one page to a base64 image.  Returns (mime, base64) or None if
    PyMuPDF is not installed.

    WebP by default.  A drawing carrying shaded views is photographic enough
    that PNG cannot compress it - the same 200 dpi page comes out about 85%
    smaller as WebP, at full resolution and with no visible difference in the
    title block.  Dropping the dpi instead would cost legibility for less.
    """
    try:
        import pymupdf  # type: ignore
    except ImportError:
        try:
            import fitz as pymupdf  # type: ignore
        except ImportError:
            return None

    fmt = (image_format or "auto").lower()
    with pymupdf.open(pdf_path) as doc:
        pix = doc[page_index].get_pixmap(dpi=dpi)

        if fmt != "png":
            encoded = _encode_via_pillow(pix, fmt, quality)
            if encoded is not None:
                return encoded
            if fmt in ("auto", "jpeg", "jpg"):   # no Pillow - PyMuPDF still does JPEG
                return "image/jpeg", base64.b64encode(
                    pix.tobytes("jpeg", jpg_quality=quality)).decode("ascii")

        return "image/png", base64.b64encode(pix.tobytes("png")).decode("ascii")


def _encode_thumb(path: str, size: int, quality: int) -> str | None:
    """One preview bitmap as a data URI, shrunk to fit `size` on its long edge.

    SolidWorks stores its preview as an uncompressed bitmap - a couple of
    hundred KB each, and a BOM's worth of those would outweigh the drawing they
    illustrate.  At card size WebP costs a few KB, with the same fall back to
    JPEG the sheet rasters use for a viewer too old for it.
    """
    if not os.path.isfile(path):
        return None
    try:
        from PIL import Image
    except ImportError:
        return None

    try:
        with Image.open(path) as image:
            image = image.convert("RGB")
            # A preview can be one flat colour - a configuration saved with
            # nothing in view - and a white box on the card is worse than no
            # box.  The macro skips those too; this is the belt to its braces.
            if all(lo == hi for lo, hi in image.getextrema()):
                return None
            image.thumbnail((size, size), Image.LANCZOS)
            buffer = io.BytesIO()
            try:
                image.save(buffer, "WEBP", quality=quality, method=6)
                mime = "image/webp"
            except Exception:
                buffer = io.BytesIO()
                image.save(buffer, "JPEG", quality=quality, optimize=True,
                           subsampling=0)
                mime = "image/jpeg"
    except Exception:
        return None

    return "data:%s;base64,%s" % (
        mime, base64.b64encode(buffer.getvalue()).decode("ascii"))


def _load_thumbs(manifest: str, config: dict, log=print
                 ) -> tuple[dict[str, str], dict[tuple[int, str], str],
                            dict[tuple[int, str], tuple[float, str]],
                            dict[tuple[int, str], str]]:
    """Read the macro's manifest.  Returns (images, previews by row, prices by
    row, files by row).

    The manifest is the tab-separated file the VBA writes beside the bitmaps:

        sheet <TAB> item <TAB> part <TAB> configuration <TAB> bitmap
              <TAB> cost <TAB> basis <TAB> file <TAB> version
              <TAB> vault part number

    `file` (optional) is the model the row resolved to, by name.  It is
    what lets a silhouette on the sheet find its row when the BOM's part number
    is a purchased number and the file is called something else.

    `version` (optional, newest) is that file's PDM version as a string -
    whatever the macro chose to write, shown as-is.  Blank on a row whose file
    is not in the vault or that resolved to no file at all.

    `sheet` is 0-based and is the same index the hotspots carry; `item` is the
    text of the BOM's leftmost column, which is what a hotspot is keyed on, so
    the balloon and the BOM row for one item both find the same picture and the
    same price.  `part` and `configuration` are read for the log only - the
    configuration is routinely empty, since the bitmap SolidWorks stores is the
    same one whichever configuration is asked for.

    `cost` is the SHOP cost of one of that part, written by modPolyCost through
    dl_RowCost, and `basis` is how it was arrived at (LENGTH, SHEET, STOCK#,
    ASSY...).  BOTH ARE OPTIONAL, and so is `bitmap`: a line carries whichever
    of the two it has.  A purchased part with no model has a price and no
    picture; a model saved without a preview graphic has a picture's worth of
    nothing and a price.  Either is a reason to write the line.

    Tab-separated rather than JSON because the writer is VBA: hand-rolled JSON
    escaping in a macro is a bug waiting for the first part number with a
    backslash in it.

    Images come back as their own table, keyed by bitmap name, and the rows
    only carry the key into it.  One part called out on three sheets is one
    picture - copying the base64 onto every hotspot would put the same image
    into the page again and again.

    Pillow is needed for the pictures and NOT for the prices, so a machine
    without it still publishes a priced page - which is why the import guard
    below sets a flag instead of returning early, as it used to.
    """
    images: dict[str, str] = {}
    by_row: dict[tuple[int, str], str] = {}
    prices: dict[tuple[int, str], tuple[float, str]] = {}
    files: dict[tuple[int, str], str] = {}
    configs: dict[tuple[int, str], str] = {}   # (sheet, item) -> the row's configuration
    versions: dict[tuple[int, str], str] = {}  # (sheet, item) -> the file's PDM version
    pns: dict[tuple[int, str], str] = {}       # (sheet, item) -> the card's StockedPartNumber
    cells: dict[tuple[int, str], str] = {}     # (sheet, item) -> the drawing's Part # cell

    if not manifest:
        return images, by_row, prices, files, configs, versions, pns, cells
    if not os.path.isfile(manifest):
        log(f"  WARNING: no preview manifest at {manifest} - no previews, no prices.")
        return images, by_row, prices, files, configs, versions, pns, cells

    try:
        import PIL  # noqa: F401
        can_shrink = True
    except ImportError:
        log("  WARNING: model previews need Pillow (pip install pillow) - "
            "prices still published.")
        can_shrink = False

    folder = os.path.dirname(os.path.abspath(manifest))
    size = int(config.get("thumb_px", 260) or 260)
    quality = int(config.get("thumb_quality", 82) or 82)
    unreadable: set[str] = set()

    # UTF-8 or the Windows code page, whichever the macro wrote - see
    # _read_manifest.
    if True:
        for line in _read_manifest(manifest):
            fields = line.split("\t")
            if len(fields) < 5 or not fields[0].strip().lstrip("-").isdigit():
                continue                       # blank line, or a header someone added
            sheet = int(fields[0])
            item = fields[1].strip()
            if not item:
                continue

            # The price, if this line carries one.  Str$ on the VBA side always
            # writes a "." decimal point whatever the workstation's regional
            # settings are, so a plain float() is safe; anything unparseable is
            # dropped rather than guessed at.
            cost = fields[5].strip() if len(fields) > 5 else ""
            if cost:
                try:
                    value = float(cost)
                except ValueError:
                    log(f"  WARNING: sheet {sheet} item {item} - unreadable "
                        f"price {cost!r}, dropped.")
                else:
                    if value > 0:
                        basis = fields[6].strip() if len(fields) > 6 else ""
                        prices[(sheet, item)] = (value, basis)

            if len(fields) > 7 and fields[7].strip():
                files[(sheet, item)] = fields[7].strip().lower()
                configs[(sheet, item)] = fields[3].strip()

            # column 3 is the drawing's own Part # cell (the SolidWorks
            # property, read off the hidden column); column 10 (2026-10-13)
            # the vault card's StockedPartNumber, "" when the card has none
            if fields[2].strip():
                cells[(sheet, item)] = fields[2].strip()
            if len(fields) > 9 and fields[9].strip():
                pns[(sheet, item)] = fields[9].strip()

            if len(fields) > 8 and fields[8].strip():
                versions[(sheet, item)] = fields[8].strip()

            bitmap = fields[4].strip()
            if not bitmap or not can_shrink:
                continue
            key = os.path.splitext(bitmap)[0]

            if key not in images:
                encoded = _encode_thumb(os.path.join(folder, bitmap), size, quality)
                if encoded is None:
                    unreadable.add(bitmap)
                    continue
                images[key] = encoded
            by_row[(sheet, item)] = key

    if unreadable:
        log(f"  WARNING: {len(unreadable)} preview bitmap(s) named in the "
            f"manifest could not be read, or were blank: "
            f"{', '.join(sorted(unreadable)[:8])}{' ...' if len(unreadable) > 8 else ''}")
    if images:
        weight = sum(len(v) for v in images.values()) * 3 / 4
        log(f"  {len(images)} model preview(s) over {len(by_row)} BOM row(s)"
            f"  ({weight / 1e6:.1f} MB)")
    if prices:
        log(f"  {len(prices)} priced BOM row(s), "
            f"{sum(v for v, _ in prices.values()):,.2f} shop cost in total")
    if versions:
        log(f"  {len(versions)} BOM row(s) carry a PDM version")
    return images, by_row, prices, files, configs, versions, pns, cells


def build_section(src_pdf: str, pdf_href: str, sheets: list[dict], drawing: str,
                  config: dict, embed_pdf: bool, viewer_mode: str, dpi: int,
                  log=print, thumbs: dict | None = None,
                  pdf_file: str = "") -> tuple[dict, str, str]:
    """Build the illustrated-parts SECTION.  Returns (section, viewer, format).

    Everything that used to be baked straight into a page is gathered here
    instead, so the same bytes can either be poured into a standalone page
    (build_html, below) or written out for the Technical Manual to absorb
    (--emit section).  Nothing about the drawing is decided differently in the
    two cases - only where the result ends up.
    """
    payload_sheets = []
    used = "raster"
    used_format = "-"

    image_format = config.get("image_format", "auto")
    quality = int(config.get("image_quality", 85))

    for spec in sheets:
        image = None
        if viewer_mode in ("auto", "raster"):
            image = _rasterise(src_pdf, spec["page"], dpi, image_format, quality)
            if image is None and viewer_mode == "raster":
                log("  WARNING: --viewer raster needs PyMuPDF "
                    "(pip install pymupdf); embedding the PDF instead.")
                viewer_mode = "pdf"
        if image is None:
            used = "pdf"
        else:
            used_format = image[0].split("/")[-1]
        payload_sheets.append({
            "label": spec["label"],
            "page": spec["page"],
            "w": round(spec["w"], 4),
            "h": round(spec["h"], 4),
            "columns": spec["columns"],
            "hotspots": spec["hotspots"],
            "image_mime": image[0] if image else "",
            "image": image[1] if image else "",
            # Zones - see build_sections / build_silhouettes.  Absent without
            # --zones, and the page does not miss them.
            "sections": spec.get("sections", []),
            "silhouettes": spec.get("silhouettes", []),
            "ref_overlay": spec.get("ref_overlay", ""),
            "ref_names": spec.get("ref_names", []),
            # The reference assembly's parts - see build_allparts.  Absent
            # from an older manifest, and the panel stays hidden.
            "allparts": spec.get("allparts", []),
        })

    payload = {
        "drawing": drawing,
        "serial": _serial_of(drawing),
        "page_zoom": float(config.get("page_zoom", 0.95) or 0.95),
        "html_action": config.get("html_action", "cart"),
        "length_units": config.get("length_units") or {},
        # Keyed data URIs, shared by every hotspot that names one - see
        # _load_thumbs.  Empty when the macro exported no previews, and the
        # card simply shows none.
        "thumbs": thumbs or {},
        # What the two panels beside the sheet are called.
        "labels": {"parts": config.get("parts_label") or "Critical parts",
                   "allparts": config.get("allparts_label") or "All parts"},
        # How to LABEL a price, not what one is: the multiplier was already
        # applied to each hotspot upstream, so there is nothing here a reader
        # could work a cost back out of.
        "price": {
            "label": config.get("price_label", "Est. Price"),
            "currency": config.get("price_currency", "$"),
            "in_bom": bool(config.get("price_in_bom", True)),
        },
        "pdm_version_label": config.get("pdm_version_label") or "",
        "sheets": payload_sheets,
        "pdf_href": pdf_href,
        "pdf_base64": "",
        "pdf_file": pdf_file,    # the PDF beside the page, for the download button
        "viewer": used,          # the section tints itself pdf-mode off this
        "quote": {
            "email": config.get("quote_email", ""),
            "subject": config.get("quote_subject", "{SERIAL}-PARTS RFQ"),
            "intro": config.get("quote_intro", ""),
        },
    }

    # Only the "pdf" viewer needs the file itself, and then only when the page
    # has to stand alone away from it.
    if embed_pdf:
        with open(src_pdf, "rb") as handle:
            payload["pdf_base64"] = base64.b64encode(handle.read()).decode("ascii")

    # The sheets ARE this section's rail rows.  Declaring them here means the
    # manual can draw the rail before the section's script has run - and the
    # script then claims those same rows to hang its counts on.
    nav = [{"id": "%s-page-%d" % (SECTION_ID, i), "label": s["label"]}
           for i, s in enumerate(payload_sheets)] if len(payload_sheets) > 1 else []

    sec = shell.section(
        SECTION_ID,
        config.get("section_label") or "Illustrated Parts",
        SECTION_HTML,
        icon=config.get("section_icon", "\U0001f6d2"),
        css=SECTION_CSS,
        js=SECTION_JS,
        payload=payload,
        nav=nav,
        slots={"header": ["quote", "align"]},
    )
    return sec, used, used_format


def build_html(out_html: str, src_pdf: str, pdf_href: str, sheets: list[dict],
               drawing: str, config: dict, embed_pdf: bool, viewer_mode: str,
               dpi: int, log=print) -> tuple[str, str]:
    """Write the standalone quote page.  Returns (viewer mode, image format).

    One section, poured into the shell - which is what this page always was,
    now said out loud.
    """
    sec, used, used_format = build_section(
        src_pdf, pdf_href, sheets, drawing, config, embed_pdf, viewer_mode,
        dpi, log)

    html = shell.render_page(
        drawing, [sec],
        subtitle=config.get("page_subtitle", "Parts & service"),
        contact=config.get("contact") or {},
        key="drawing:" + drawing,
    )
    with open(out_html, "w", encoding="utf-8") as handle:
        handle.write(html)

    return used, used_format


def build_section_file(out_json: str, *args, **kw) -> tuple[str, str]:
    """Write <base>_section.json for the Technical Manual to absorb."""
    sec, used, used_format = build_section(*args, **kw)
    shell.write_section(out_json, sec)
    return used, used_format


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------

def _pillow() -> bool:
    try:
        import PIL  # noqa: F401
        return True
    except ImportError:
        return False


def _serial_of(drawing: str) -> str:
    """The machine serial inside a drawing number: the four characters after
    the leading letter, so A9987-0000-M -> 9987."""
    found = re.match(r"^[A-Za-z](\w{4})", drawing)
    return found.group(1) if found else drawing


def _sheet_label(sheet: Sheet, index: int, total: int,
                 override: str | None = None) -> str:
    """What the tab says.

    The macro passes SolidWorks' own sheet names, which is what anyone
    actually calls them, so those win.  Failing that a paperspace layout keeps
    its layout name, and a per-sheet file falls back to its number.
    """
    if override and override.strip():
        return override.strip()
    name = (sheet.name or "").strip()
    if name and not name.lower().startswith("layout") and total > 1:
        # "A9987-0000-M_S2" -> "Sheet 2"; anything else is a real name
        tail = name.rsplit("_", 1)[-1]
        if len(tail) <= 4 and tail.upper().startswith("S") and tail[1:].isdigit():
            return f"Sheet {int(tail[1:])}"
        return name
    return f"Sheet {index + 1}" if total > 1 else "Sheet 1"


def _pair_sheets_to_pages(sheets: list[Sheet], n_pages: int, page_index: int,
                          dxf_paths: list[str]) -> list[int]:
    """Decide which PDF page each parsed sheet belongs to."""
    if len(sheets) == 1:
        if page_index >= n_pages:
            raise IndexError(
                f"PDF has {n_pages} page(s); --page {page_index} is past the end.")
        return [page_index]

    if len(sheets) == n_pages:
        return list(range(n_pages))

    # The usual cause: the DXF export is set to "all sheets", so every
    # per-sheet file came out holding all of them.
    per_file: dict[str, int] = {}
    for sheet in sheets:
        per_file[sheet.source] = per_file.get(sheet.source, 0) + 1
    detail = ", ".join(f"{os.path.basename(p)}: {n} sheet(s)"
                       for p, n in per_file.items())
    raise ValueError(
        f"Found {len(sheets)} drawing sheet(s) across {len(dxf_paths)} DXF file(s) "
        f"but the PDF has {n_pages} page(s), so they cannot be paired up.\n"
        f"  {detail}\n"
        f"  Give one DXF per sheet, in sheet order, and a PDF with one page per "
        f"sheet.  The SW_DrawingLinks macro does this for you; by hand, set "
        f"Tools > Options > Export > DXF/DWG to export the active sheet only "
        f"and save each sheet separately.")


def process(dxf_paths, pdf_path: str | None = None, out_dir: str | None = None,
            config: dict | None = None, embed_pdf: bool = False,
            viewer_mode: str = "auto", dpi: int = 200,
            draw_boxes: bool = True, page_index: int = 0,
            sheet_names: list[str] | None = None, work_dir: str | None = None,
            linked_pdf: bool = False, emit: str = "page",
            thumbs_path: str | None = None, zones_path: str | None = None,
            log=print) -> dict:
    """Generate the linked PDF / HTML / JSON.  Returns a summary dict.

    `dxf_paths` is one DXF or a list of them, in sheet order.
    `sheet_names` are SolidWorks' own sheet names, matched positionally.
    `work_dir` takes everything nobody opens - the hotspot map, and the
    annotated PDF when `linked_pdf` asks for one.  `out_dir` gets the single
    deliverable, the page.
    `emit` is "page" (the standalone quote page - the default, and what the
    SolidWorks macro asks for), "section" (a manifest for the Technical Manual
    to absorb), or "both".
    `thumbs_path` is the macro's model-preview manifest; see _load_thumbs.
    `zones_path` is the macro's zones manifest; see _load_zones.
    """
    cfg = dict(DEFAULT_CONFIG)
    if config:
        cfg.update(config)

    if isinstance(dxf_paths, str):
        dxf_paths = [dxf_paths]
    dxf_paths = [os.path.abspath(p) for p in dxf_paths]
    for path in dxf_paths:
        if not os.path.isfile(path):
            raise FileNotFoundError(path)
    if not dxf_paths:
        raise ValueError("No DXF given.")

    # Name the outputs after the first DXF, minus any _S1 / _Sheet1 suffix the
    # per-sheet export added.
    base = os.path.splitext(os.path.basename(dxf_paths[0]))[0]
    base = re.sub(r"[_-](S|SHEET)0*1$", "", base, flags=re.IGNORECASE)

    if pdf_path is None:
        pdf_path = os.path.join(os.path.dirname(dxf_paths[0]), base + ".pdf")
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(
            f"No PDF beside the DXF - looked for {pdf_path}. Pass --pdf.")

    out_dir = os.path.abspath(out_dir or os.path.dirname(pdf_path))
    os.makedirs(out_dir, exist_ok=True)

    # The hotspot map is a debugging / integration artefact - the page does not
    # read it - so it goes to the scratch folder unless asked for otherwise.
    work_dir = os.path.abspath(
        work_dir or os.path.join(tempfile.gettempdir(), "DRAWING_LINKS"))
    os.makedirs(work_dir, exist_ok=True)

    sheets: list[Sheet] = []
    for path in dxf_paths:
        found = parse_dxf(path)
        if not found:
            log(f"  WARNING: {os.path.basename(path)} has no usable extents - skipped.")
        elif not any(s.balloons or s.table for s in found):
            log(f"  WARNING: no balloons or BOM in {os.path.basename(path)} - "
                f"kept as a page with no hotspots.")
        sheets.extend(found)
    if not any(s.balloons or s.table for s in sheets):
        raise ValueError("No balloons or BOM rows found in the DXF(s).")

    reader = PdfReader(pdf_path)
    pages = _pair_sheets_to_pages(sheets, len(reader.pages), page_index, dxf_paths)

    out_html = os.path.join(out_dir, base + ".html")
    out_json = os.path.join(work_dir, base + "_links.json")
    out_section = os.path.join(out_dir, base + "_section.json")

    # The annotated PDF is a by-product now, not a deliverable, so it goes to
    # the scratch folder and the published folder holds only the page.  It
    # keeps the _linked suffix THERE on purpose: the PDF it is built from is
    # already <base>.pdf in that same folder, and without the suffix the
    # annotated copy would overwrite its own source.
    out_pdf = os.path.join(work_dir, base + "_linked.pdf")

    # The PDF's hotspots link back to the HTML page, so its URL has to be
    # settled before the annotations are written.
    link_base = (cfg.get("html_link_base") or "").strip()
    if link_base:
        html_url = link_base.rstrip("/") + "/" + os.path.basename(out_html)
    else:
        html_url = pathlib.Path(out_html).as_uri()

    thumb_images, thumb_rows, row_prices, row_files, row_configs, row_versions, \
        row_pns, row_cells = \
        _load_thumbs(thumbs_path or "", cfg, log)
    show_version = bool((cfg.get("pdm_version_label") or "").strip())

    # The margin is applied ONCE, here, on the way into the page - so nothing
    # downstream has to remember to, and a page built at 0 has no cost figure in
    # its source at all rather than a hidden one.
    price_mult = float(cfg.get("price_multiplier") or 0)
    materials = _load_materials(cfg, log) if price_mult > 0 else {}

    zones = _load_zones(zones_path or "", log)
    main_index = zones["main"]["sheet"] if zones and zones.get("main") else -1

    specs, by_page, json_sheets = [], {}, []
    for index, (sheet, page) in enumerate(zip(sheets, pages)):
        given = sheet_names[index] if sheet_names and index < len(sheet_names) else None
        label = _sheet_label(sheet, index, len(sheets), given)
        pdf_page = reader.pages[page]
        box = pdf_page.mediabox
        rotate = int(pdf_page.get("/Rotate", 0) or 0) % 360
        page_ratio = (float(box.height) / float(box.width) if rotate in (90, 270)
                      else float(box.width) / float(box.height))
        _choose_extents(sheet, page_ratio, label, log)

        transform = Transform(sheet, pdf_page, cfg)
        _vault_parts(sheet, index, row_files, cfg, log, row_pns, row_cells)
        hotspots = build_hotspots(sheet, transform, base, cfg, index)

        # The balloon and the BOM row for one item are two hotspots with the
        # same item number, so both pick up the same preview and the same price
        # here.  That is what makes the price a property of the ITEM rather than
        # of the balloon: a drawing that never got a balloon for item 12 still
        # has a row for it, and the row is priced.
        for hotspot in hotspots:
            key = (index, str(hotspot["item"]).strip())
            found = thumb_rows.get(key)
            if found:
                hotspot["thumb"] = found
            priced = row_prices.get(key)
            # a length material is costed HERE, on the length this row uses -
            # the macro's figure for it is one purchase unit
            by_len = None
            if materials:
                row = next((r for r in sheet.table.rows if str(r.item).strip() == key[1]), None) if sheet.table else None
                by_len = _length_price(hotspot, materials, list(row.values.values()) if row else None)
            if by_len:
                if hotspot["kind"] == "row":
                    log(f"    item {key[1]}: {by_len[1]} = ${by_len[0]:.2f} cost"
                        f"{f' (macro said ${priced[0]:.2f})' if priced else ''}")
                priced = by_len
            if priced and price_mult > 0:
                hotspot["price"] = round(priced[0] * price_mult, 2)
                if cfg.get("price_show_basis", True) and priced[1]:
                    hotspot["price_basis"] = priced[1]
            version = row_versions.get(key)
            if version and show_version:
                hotspot["pdm_ver"] = version
            if row_files.get(key):
                hotspot["file"] = row_files[key]

        columns = sheet.table.columns if sheet.table else []
        by_page[page] = hotspots
        specs.append({"label": label, "page": page, "w": transform.vis_w,
                      "h": transform.vis_h, "columns": columns,
                      "hotspots": hotspots})
        json_sheets.append({
            "label": label, "page_index": page, "source_dxf": sheet.source,
            "layout": sheet.name, "extents_from": sheet.ext_source,
            "page_size_pt": [transform.vis_w, transform.vis_h],
            "scale_pt_per_dxf_unit": transform.scale_x,
            "columns": columns, "hotspots": hotspots,
        })

        n_b = sum(1 for h in hotspots if h["kind"] == "balloon")
        n_r = sum(1 for h in hotspots if h["kind"] == "row")
        log(f"  {label} (page {page + 1}): {n_b} balloon(s), {n_r} BOM row(s), "
            f"scale {transform.scale_x:.6f} pt/unit from {sheet.ext_source} "
            f"(axis mismatch {transform.residual * 100:.3f}%)")
        # A section sheet whose DXF holds no view block at all was exported
        # before SolidWorks had generated its views: no silhouette can be
        # drawn for anything on it, critical or reference, and every part is
        # a box.  Said here by name, once per sheet, so it is not mistaken
        # for the layers failing.
        if zones is not None and index != main_index and not sheet.blocks:
            log(f"    WARNING: {os.path.basename(sheet.source)} carries NO drawing-view geometry "
                f"(no DRAWING_VIEW block) - the sheet was exported before its views were "
                f"generated.  Open it in SolidWorks and publish again; nothing on it can be outlined.")

        # The silhouettes.  Drawn from the DXF, named from the manifest, so a
        # sheet the manifest says nothing about still gets them - unnamed.
        if zones is not None and index != main_index:
            sheet_files = {(f, (row_configs.get((s, item)) or "").upper()): item
                           for (s, item), f in row_files.items() if s == index}
            sils, ref_overlay, ref_names = build_silhouettes(
                sheet, transform, zones["top"].get(index, []), cfg, log,
                row_files=sheet_files, images=thumb_images,
                view=(zones.get("views") or {}).get(index),
                bitmap_dir=os.path.join(os.path.dirname(os.path.abspath(zones_path)), "thumbs"))
            if sils:
                specs[-1]["silhouettes"] = sils
                specs[-1]["ref_overlay"] = ref_overlay
                specs[-1]["ref_names"] = ref_names
                # A row the BOM scan could not picture borrows its component's
                # preview: same file, seen from the other side.
                borrowed = 0
                for z in sils:
                    if not (z["item"] and z["thumb"]):
                        continue
                    for hotspot in hotspots:
                        if str(hotspot["item"]).strip() == z["item"] and not hotspot.get("thumb"):
                            hotspot["thumb"] = z["thumb"]
                            borrowed += 1
                named = sum(1 for z in sils if z["n_names"])
                in_bom = sum(1 for z in sils if z["item"])
                pictured = sum(1 for z in sils if z["thumb"])
                log(f"    {len(sils)} silhouette(s), {named} named, {in_bom} on the parts list, "
                    f"{pictured} with a preview"
                    + (f", {borrowed} row(s) pictured from them" if borrowed else "")
                    + (f", !REF hidden: {', '.join(ref_names)}" if ref_names else "")
                    + (f", REF layer: {sum(len(v) for v in sheet.ref_blocks.values())} "
                       f"polyline(s) in {len(sheet.ref_blocks)} view block(s) from "
                       f"{os.path.basename(sheet.source)[:-4]}{REF_LAYER_SUFFIX}"
                       if sheet.ref_blocks else ""))
                # What the REF-only export is made of - the one line that says
                # whether the macro's per-part layers survived the export, and
                # whether the blocks still name the parts (both are how an
                # outline gets drawn; neither, and every part is a box).
                if sheet.ref_blocks or sheet.ref_layers:
                    names = sorted(sheet.ref_blocks)
                    log(f"    REF export: {len(sheet.ref_blocks)} view block(s)"
                        + (" - " + ", ".join(n[len(VIEW_BLOCK_PREFIX):] if n.startswith(VIEW_BLOCK_PREFIX) else n
                                             for n in names[:6]) + (" ..." if len(names) > 6 else "")
                           if names else "")
                        + f"; {len(sheet.ref_layers)} part layer(s) carrying "
                        f"{sum(len(v) for v in sheet.ref_layers.values())} polyline(s)")
            elif sheet.blocks and not _pillow():
                log("    silhouettes need Pillow (pip install pillow) - none drawn.")
            # The reference assembly's parts - the ALL PARTS panel - when the
            # macro walked them (REFPART lines; a 2026-10-09 macro).
            refparts = (zones.get("refparts") or {}).get(index, [])
            if refparts:
                aps = build_allparts(
                    sheet, transform, refparts, cfg, log,
                    row_files=sheet_files, images=thumb_images,
                    view=(zones.get("views") or {}).get(index),
                    bitmap_dir=os.path.join(os.path.dirname(os.path.abspath(zones_path)), "thumbs"),
                    show_version=show_version, pdmpn=zones.get("pdmpn"))
                if aps:
                    specs[-1]["allparts"] = aps

    # MAIN's click zones, one per section sheet, placed from the manifest.
    if zones is not None and 0 <= main_index < len(specs):
        specs[main_index]["sections"] = build_sections(
            zones, [s["label"] for s in specs], log)

    # Build the section ONCE - rasterising every sheet is the expensive part of
    # this whole tool - then write it out in whichever shapes were asked for.
    # What the page falls back to when it cannot rasterise has to actually
    # resolve.  The PDF is no longer beside the HTML, so a bare basename would
    # point at nothing; name the file itself.  Without --linked-pdf there is no
    # annotated copy to point at, so the plain exported PDF stands in - which is
    # also what the old code should have done, having handed the page a
    # _linked.pdf that was never written unless the flag was set.
    viewer_pdf = out_pdf if linked_pdf else pdf_path

    wants_page = emit in ("page", "both")
    wants_section = emit in ("section", "both")

    # The drawing as one PDF, beside the page, for the page's download button.
    # The only file that joins the HTML in the publish folder: every sheet in
    # one document is what a reader without the page wants, and too large to
    # ride inside it (tens of MB against a 3 MB page).  Named relatively so
    # the link holds from the share and from a copied folder alike; a copy
    # that fails costs the button, nothing more.
    pdf_file = ""
    if wants_page:
        pdf_beside = os.path.join(out_dir, base + ".pdf")
        try:
            shutil.copyfile(pdf_path, pdf_beside)
            pdf_file = base + ".pdf"
        except OSError as exc:
            log(f"  WARNING: could not copy the PDF beside the page ({exc}) - no download button.")

    sec, used_viewer, used_format = build_section(
        pdf_path, pathlib.Path(viewer_pdf).as_uri(), specs, base, cfg,
        embed_pdf, viewer_mode, dpi, log, thumbs=thumb_images, pdf_file=pdf_file)

    if wants_page:
        html = shell.render_page(
            base, [sec], subtitle=cfg.get("page_subtitle", "Parts & service"),
            contact=cfg.get("contact") or {}, key="drawing:" + base)
        with open(out_html, "w", encoding="utf-8") as handle:
            handle.write(html)
    if wants_section:
        shell.write_section(out_section, sec,
                            hidden=bool(cfg.get("hide_section_json", True)))

    # The annotated PDF is opt-in and lands in the scratch folder: the quote
    # page is the deliverable, and a second copy of the drawing beside it in
    # the publish folder was just something to keep in step.
    if linked_pdf:
        build_pdf(pdf_path, out_pdf, by_page, cfg, draw_boxes, html_url)

    with open(out_json, "w", encoding="utf-8") as handle:
        json.dump({
            "drawing": base,
            "source_dxf": dxf_paths,
            "source_pdf": pdf_path,
            "sheets": json_sheets,
        }, handle, indent=2)

    all_hotspots = [h for spec in specs for h in spec["hotspots"]]
    n_balloons = sum(1 for h in all_hotspots if h["kind"] == "balloon")
    n_rows = sum(1 for h in all_hotspots if h["kind"] == "row")
    log(f"  {len(specs)} sheet(s), {n_balloons} balloon(s), {n_rows} BOM row(s)")
    log(f"  HTML drawing: {used_viewer}"
        + (f"  ({used_format}, {dpi} dpi, exact alignment)"
           if used_viewer == "raster"
           else "  (browser PDF viewer - use the Align sliders if it is off)"))
    if wants_page:
        log(f"  -> {out_html}  ({os.path.getsize(out_html) / 1e6:.1f} MB)")
        if pdf_file:
            log(f"  -> {os.path.join(out_dir, pdf_file)}  "
                f"({os.path.getsize(os.path.join(out_dir, pdf_file)) / 1e6:.1f} MB)  the download button's PDF")
    if wants_section:
        log(f"  -> {out_section}  ({os.path.getsize(out_section) / 1e6:.1f} MB)"
            f"  section '{sec['id']}' for the Technical Manual"
            + ("  [hidden]" if cfg.get("hide_section_json", True) else ""))
    if linked_pdf:
        log(f"  (working) {out_pdf}")
    log(f"  (working) {out_json}")

    return {"pdf": out_pdf if linked_pdf else "",
            "html": out_html if wants_page else "",
            "section": out_section if wants_section else "",
            "json": out_json, "sheets": len(specs),
            "balloons": n_balloons, "rows": n_rows}


def run_gui() -> None:
    """Drag-and-drop window, same idea as DXF2HTML.py."""
    from tkinter import Label, filedialog

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        root = TkinterDnD.Tk()
        has_dnd = True
    except ImportError:
        from tkinter import Tk
        root = Tk()
        has_dnd = False

    root.title("SolidWorks Drawing -> Clickable PDF/HTML")
    root.geometry("580x300")
    root.configure(bg="#f5f5f5")

    prompt = ("Drag & drop the drawing .DXF here"
              if has_dnd else "Click to choose the drawing .DXF")
    drop = Label(root,
                 text=f"[+]\n\n{prompt}\n(the .pdf of the same name is used)\n"
                      f"multi-sheet: drop one DXF per sheet together",
                 font=("Segoe UI", 13, "bold"), bg="white", fg="#555",
                 bd=2, relief="solid", cursor="hand2")
    drop.pack(fill="both", expand=True, padx=22, pady=22)

    status = Label(root, text="Waiting for a file...", font=("Segoe UI", 9, "italic"),
                   bg="#f5f5f5", fg="#777", wraplength=540, justify="center")
    status.pack(pady=(0, 14))

    def handle(paths) -> None:
        paths = [p for p in paths if p]
        if not paths or not all(p.lower().endswith(".dxf") for p in paths):
            status.config(text="Drop .DXF files only.", fg="#c62828")
            return
        # sheet order is the file order, and per-sheet exports sort naturally
        paths = sorted(paths)
        status.config(text="Working...", fg="#1565c0")
        root.update()
        try:
            result = process(paths, log=lambda m: None)
            status.config(
                text=(f"Done - {result['sheets']} sheet(s), {result['balloons']} "
                      f"balloons, {result['rows']} rows.\n"
                      f"{os.path.basename(result['html'])}"), fg="#2e7d32")
            webbrowser.open("file://" + result["html"])
        except Exception as exc:  # surfaced in the window, not a traceback
            status.config(text=f"{type(exc).__name__}: {exc}", fg="#c62828")

    if has_dnd:
        drop.drop_target_register(DND_FILES)
        # tkinterdnd2 hands back a Tcl list: paths with spaces come in {braces}
        drop.dnd_bind("<<Drop>>", lambda e: handle(
            [p.strip("{}") for p in root.tk.splitlist(e.data)]))
    drop.bind("<Button-1>", lambda e: handle(list(
        filedialog.askopenfilenames(filetypes=[("DXF", "*.dxf"), ("All", "*.*")]) or [])))

    root.mainloop()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Overlay clickable balloon / BOM hotspots onto a drawing PDF.")
    parser.add_argument("dxf", nargs="*",
                        help="drawing DXF(s) exported from SolidWorks - one per "
                             "sheet, in sheet order")
    parser.add_argument("--pdf", help="source PDF (default: same name as the DXF)")
    parser.add_argument("--out", help="output folder (default: beside the PDF)")
    parser.add_argument("--config", help="JSON config file")
    parser.add_argument("--sheet-names", nargs="+", metavar="NAME",
                        help="SolidWorks' own sheet names, in the same order as "
                             "the DXFs - these become the tab labels")
    parser.add_argument("--work", metavar="DIR",
                        help="where the working files go (default: "
                             "%%TEMP%%\\DRAWING_LINKS).  Only the HTML and the "
                             "linked PDF land in --out.")
    parser.add_argument("--thumbs", metavar="PATH",
                        help="the BOM manifest written by the SolidWorks macro "
                             "(thumbs.txt, bitmaps beside it).  Puts each "
                             "item's saved model preview AND its price in the "
                             "part card.  Pillow is needed for the previews "
                             "only.")
    parser.add_argument("--zones", metavar="PATH",
                        help="the zones manifest written by the SolidWorks "
                             "macro (zones.txt).  Click zones on MAIN for "
                             "every section sheet, and a silhouette for each "
                             "top-level component on its sheet.  Pillow needed.")
    parser.add_argument("--page", type=int, default=0, help="0-based PDF page index")
    parser.add_argument("--open", action="store_true", help="open the HTML when done")
    parser.add_argument("--viewer", choices=("auto", "raster", "pdf"), default="auto",
                        help="how the drawing is shown in the HTML (default: auto)")
    parser.add_argument("--dpi", type=int, default=200,
                        help="raster resolution (default: 200)")
    parser.add_argument("--image-format", choices=("auto", "webp", "jpeg", "png"),
                        help="how each sheet is encoded (default: auto = WebP)")
    parser.add_argument("--quality", type=int,
                        help="webp/jpeg quality, 1-100 (default: 85)")
    parser.add_argument("--embed-pdf", action="store_true",
                        help="inline the PDF into the HTML so the page stands "
                             "alone (roughly quadruples its size)")
    parser.add_argument("--linked-pdf", action="store_true",
                        help="also write <base>_linked.pdf into --work - the "
                             "drawing with clickable hotspots that open this "
                             "page.  A by-product, not a deliverable; --out "
                             "gets the page only.")
    parser.add_argument("--no-boxes", action="store_true",
                        help="with --linked-pdf, do not draw the hotspot outlines")
    parser.add_argument("--emit", choices=("page", "section", "both"), default="page",
                        help="page: the standalone quote page (default).  "
                             "section: <base>_section.json for the Technical "
                             "Manual to absorb.  both: one of each.")
    parser.add_argument("--section-label", metavar="TEXT",
                        help="the manual's rail label for this section "
                             "(default: Illustrated Parts)")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    if not args.dxf:
        run_gui()
        return 0

    config = {}
    if args.config:
        # utf-8-sig: Notepad and PowerShell's Set-Content both write a BOM.
        with open(args.config, encoding="utf-8-sig") as handle:
            config = json.load(handle)
    # explicit flags win over the config file
    if args.image_format:
        config["image_format"] = args.image_format
    if args.quality:
        config["image_quality"] = args.quality
    if args.section_label:
        config["section_label"] = args.section_label

    log = (lambda *a: None) if args.quiet else print
    log("DRAWING_LINKS: " + " + ".join(os.path.basename(p) for p in args.dxf))
    try:
        result = process(args.dxf, args.pdf, args.out, config,
                         embed_pdf=args.embed_pdf, viewer_mode=args.viewer,
                         dpi=args.dpi, draw_boxes=not args.no_boxes,
                         page_index=args.page, sheet_names=args.sheet_names,
                         work_dir=args.work, linked_pdf=args.linked_pdf,
                         emit=args.emit, thumbs_path=args.thumbs,
                         zones_path=args.zones, log=log)
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    if args.open and result["html"]:
        webbrowser.open("file://" + result["html"])
    elif args.open:
        print("--open needs a page; --emit section only wrote a manifest.",
              file=sys.stderr)
    return 0



# ==========================================================================
# POLYPACK_SHELL - embedded copy
# ==========================================================================
# What follows is backend/tools/tech-manual-publish/POLYPACK_SHELL.py, less its
# imports and its CLI, pasted in whole so that DRAWING_LINKS.py stands alone.
# The header, the contact drawer, the rail, the toast and section()/render_page()
# all live here.  If the shell changes in the repo, paste it in again between
# this line and the __main__ guard.
# ==========================================================================


SCHEMA_VERSION = 1
FORMAT = "polypack-section"

# --------------------------------------------------------------------------
# The chrome
# --------------------------------------------------------------------------

# Polypack's own tokens, straight off the app: the red, the condensed display
# face, the white header under a red rule.  The font stacks lead with faces
# Windows already has, so the page still looks right with no network - it gets
# emailed around and opened off a share.
SHELL_CSS = r"""
  :root {
    --red: #e10600; --red2: #ff2a1f; --red-soft: #fdeceb;
    --bg: #e9ebef; --panel: #fff; --panel2: #f3f4f7; --shade: #f7f8fa;
    --line: #d4d8e0; --ink: #1d2026; --muted: #6c7280; --dim: #9aa0ab;
    --ok: #1a7f45; --hot: #e58a00;
    --display: 'Bahnschrift', 'Bahnschrift Condensed', 'Bebas Neue',
               'Arial Narrow', sans-serif;
    --body: 'Oswald', 'Segoe UI', Arial, sans-serif;
    /* Two stick points, because two things pin to the top and one sits under
       the other.  --headh is the header alone: the contact drawer parks there.
       --stickh is everything currently pinned (header + the drawer when it is
       open): the rail parks there, so opening the drawer pushes the rail down
       instead of the drawer covering its top rows.  Both are re-measured by
       headerHeight() at runtime - a header that wraps onto a second line on a
       narrow window is taller than this default. */
    --headh: 54px;
    --stickh: 54px;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #14171c; --panel: #1c2027; --panel2: #232830; --shade: #1a1e24;
      --line: #2c323b; --ink: #e6e8ec; --muted: #9aa3af; --dim: #79818d;
      --red-soft: #3a1512; --ok: #4ac47f; --hot: #ffab2e;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: var(--body); font-size: 13px;
         background: var(--bg); color: var(--ink); }

  /* ===== Header ===== */
  header { display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
           background: var(--panel); border-bottom: 2px solid var(--red);
           padding: 0 16px; min-height: 54px; position: sticky; top: 0; z-index: 6; }
  .brand { font-family: var(--display); font-size: 27px; letter-spacing: 2px;
           color: var(--red); line-height: 1; display: flex; align-items: baseline;
           gap: 9px; white-space: nowrap; }
  .brand small { color: var(--dim); font-size: 11px; letter-spacing: 1.5px; }
  .dwg { font-family: var(--display); font-size: 17px; letter-spacing: 1.2px;
         padding-left: 16px; border-left: 1px solid var(--line); line-height: 1.15; }
  .dwg span { display: block; font-family: var(--body); font-size: 10.5px;
              letter-spacing: .5px; color: var(--muted); text-transform: uppercase; }
  .spacer { flex: 1; }
  .phone { font-family: var(--display); font-size: 19px; letter-spacing: 1px;
           color: var(--ink); text-decoration: none; white-space: nowrap; }
  .phone:hover { color: var(--red); }
  .phone .lbl { display: block; font-family: var(--body); font-size: 10px;
                letter-spacing: .5px; color: var(--muted); text-transform: uppercase; }

  button, .btn { font: inherit; font-family: var(--body); padding: 6px 12px;
                 border: 1px solid var(--line); background: var(--panel);
                 color: var(--ink); border-radius: 4px; cursor: pointer; }
  button:hover { border-color: var(--red); color: var(--red); }
  button.on { background: var(--red); border-color: var(--red); color: #fff; }
  button.primary { background: var(--red); border-color: var(--red); color: #fff;
                   font-family: var(--display); letter-spacing: 1.2px;
                   text-transform: uppercase; font-size: 15px; padding: 7px 14px; }
  button.primary:hover { background: var(--red2); border-color: var(--red2);
                         color: #fff; }
  button.primary:disabled { opacity: .4; cursor: not-allowed; background: var(--red);
                            border-color: var(--red); }
  button.icon { padding: 1px 7px; line-height: 1.4; }
  /* a section's own header controls, tagged with its root class so its CSS
     still reaches them out here - see SHELL.headerSlot */
  .slotwrap { display: flex; align-items: center; gap: 10px; }
  .count { display: inline-block; min-width: 19px; padding: 0 5px; margin-left: 7px;
           border-radius: 10px; background: #fff; color: var(--red);
           font-family: var(--body); font-size: 11px; font-weight: 600;
           letter-spacing: 0; text-align: center; }
  [hidden] { display: none !important; }

  /* ===== Contact drawer =====
     STICKY, directly under the header.  It is a sibling of <header>, so without
     this it opens at its document position - roughly 54px down the page.  Click
     Contact after scrolling into section 7 of a long manual and the drawer
     opened somewhere thousands of pixels above the viewport: nothing appeared
     to happen.  z-index below the header's 6 so it slides out from under it. */
  .contactbar { background: var(--shade); border-bottom: 1px solid var(--line);
                padding: 12px 16px; display: flex; flex-wrap: wrap; gap: 26px;
                position: sticky; top: var(--headh); z-index: 5;
                box-shadow: 0 6px 14px rgba(0,0,0,.09); }
  .cblock { min-width: 150px; }
  .cblock .cl { font-size: 10px; letter-spacing: .8px; text-transform: uppercase;
                color: var(--muted); margin-bottom: 2px; }
  .cblock a { display: block; color: var(--ink); text-decoration: none;
              font-family: var(--display); font-size: 16px; letter-spacing: .6px; }
  .cblock a:hover { color: var(--red); }
  /* pre-line so a postal address keeps the line breaks it was written with -
     it is set as text, and HTML would otherwise run the three lines together. */
  .cblock .addr { font-size: 12px; color: var(--muted); max-width: 230px;
                  line-height: 1.35; margin-bottom: 5px; white-space: pre-line; }

  /* ===== Layout ===== */
  .layout { display: flex; gap: 14px; padding: 14px; align-items: flex-start; }
  @media (max-width: 1100px) {
    .layout { flex-direction: column; }
    /* Stacked, the rail is a horizontal strip above the stage - sticking it
       would pin a growing band of buttons over the content it navigates. */
    .sheetbar { width: auto; flex-direction: row; flex-wrap: wrap;
                position: static; max-height: none; overflow: visible; }
    .sheetbar .barhead { display: none; }
  }

  /* ===== The rail =====
     One group per section: the section's own button, then its nav rows.  With a
     single section the button is hidden and only its rows show - which is
     exactly the sheet switcher the drawing page has always had. */
  /* STICKY, and scrolls within itself.  A manual is long - the rail is how you
     get around it, so scrolling it off the top is the one thing it must not do.
     `align-items: flex-start` on .layout is what makes this work at all: a
     stretched flex item is already as tall as the container and has nothing to
     stick to.  max-height keeps a twenty-row rail inside the viewport rather
     than letting its tail hang below the fold where sticky can never reach it. */
  .sheetbar { flex: 0 0 auto; width: 176px; display: flex; flex-direction: column;
              gap: 5px; position: sticky; top: calc(var(--stickh) + 8px);
              max-height: calc(100vh - var(--stickh) - 22px);
              overflow-y: auto; overscroll-behavior: contain;
              padding-right: 2px; transition: top .12s; }
  .sheetbar:empty { display: none; }
  .sheetbar .barhead { font-family: var(--display); font-size: 12px;
                       letter-spacing: 1px; color: var(--muted);
                       text-transform: uppercase; padding: 2px 4px 3px; }
  .sheetbar button { display: flex; align-items: center; gap: 8px; text-align: left;
                     padding: 7px 10px; font-size: 13px; }
  .sheetbar button .lbl { flex: 1; min-width: 0; overflow-wrap: anywhere; }
  .sheetbar button .n { flex: 0 0 auto; min-width: 16px; padding: 0 5px;
                        border-radius: 9px; background: var(--ok); color: #fff;
                        font-size: 11px; font-weight: 600; text-align: center; }
  .sheetbar button.on .n { background: rgba(255,255,255,.3); }
  .navgroup { display: flex; flex-direction: column; gap: 5px; }
  .navsub { display: none; flex-direction: column; gap: 3px; }
  .navgroup.on .navsub { display: flex; }
  /* NESTED under a section, the rows are secondary: indented off a hairline,
     quiet until hovered.  SOLO - one section, so the rail is only ever those
     rows - they are the whole navigation and keep the full button treatment
     above, which is exactly what the standalone drawing page has always shown
     for its sheet switcher.  Hence :not(.solo): the quiet styling is the
     exception, not the rule. */
  .sheetbar:not(.solo) .navsub { margin: 0 0 4px 8px; padding-left: 7px;
                                 border-left: 1px solid var(--line); }
  .sheetbar:not(.solo) .navsub button { padding: 4px 8px; font-size: 12px;
                                        background: transparent;
                                        border-color: transparent;
                                        color: var(--muted); }
  .sheetbar:not(.solo) .navsub button:hover { color: var(--red); }
  .sheetbar:not(.solo) .navsub button.on { color: var(--red);
                                           border-color: var(--line);
                                           background: var(--panel); }
  .sheetbar:not(.solo) .navsub button.deep { padding-left: 18px;
                                             font-size: 11.5px; }
  .sheetbar.solo .navsub { gap: 5px; }

  /* A GROUP - several sections under one rail row (Electrical over its
     schematics, Procedures over each Word export).  Its children stay folded
     away until one of them is the open section, so a twenty-page manual still
     opens as a short list. */
  .navband { display: flex; flex-direction: column; gap: 5px; }
  .navband .bandhead .chev { color: var(--muted); transition: transform .15s; }
  .navband.on .bandhead .chev { transform: rotate(90deg); }
  .navkids { display: none; flex-direction: column; gap: 5px;
             margin-left: 8px; padding-left: 7px;
             border-left: 1px solid var(--line); }
  .navband.on .navkids { display: flex; }

  /* ===== The stage - one section shown at a time ===== */
  .stage { position: relative; flex: 1 1 auto; min-width: 0; display: flex;
           gap: 14px; align-items: flex-start; }
  .pp-sec { display: none; flex: 1 1 auto; min-width: 0; }
  .pp-sec.on { display: flex; gap: 14px; align-items: flex-start; width: 100%; }

  /* ===== Panels - generic, every section may use them ===== */
  aside { flex: 0 0 440px; max-width: 100%; display: flex; flex-direction: column;
          gap: 14px; }
  .panel { background: var(--panel); border: 1px solid var(--line);
           border-radius: 4px; overflow: hidden; }
  .panel h2 { margin: 0; padding: 9px 12px; font-family: var(--display);
              font-size: 14px; letter-spacing: 1.1px; text-transform: uppercase;
              color: var(--ink); border-bottom: 1px solid var(--line);
              display: flex; align-items: center; gap: 8px; background: var(--shade); }
  .panel h2 .spacer { flex: 1; }
  .panel h2 .sub { font-family: var(--body); font-size: 11px; letter-spacing: .4px;
                   color: var(--muted); text-transform: none; }
  .panel h2 .chev { transition: transform .15s; color: var(--muted); }
  .panel.collapsed .chev { transform: rotate(-90deg); }
  .panel.collapsed .panel-body { display: none; }
  .panel h2.clickable { cursor: pointer; }
  .panel h2.clickable:hover { background: var(--red-soft); }

  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th, td { padding: 6px 9px; text-align: left; border-bottom: 1px solid var(--line);
           vertical-align: middle; }
  th { font-size: 10.5px; letter-spacing: .6px; text-transform: uppercase;
       color: var(--muted); }
  td.num, th.num { text-align: right; white-space: nowrap; }
  td.ctr, th.ctr { text-align: center; }

  .toast { position: fixed; left: 50%; bottom: 22px; transform: translateX(-50%);
           background: var(--ink); color: var(--panel); padding: 8px 16px;
           border-radius: 3px; font-size: 13px; opacity: 0; pointer-events: none;
           transition: opacity .18s, transform .18s; z-index: 50; }
  .toast.show { opacity: .96; transform: translateX(-50%) translateY(-4px); }
"""

SHELL_HTML = r"""
<header>
  <div class="brand" id="brand"></div>
  <div class="dwg">__TITLE__<span>__SUBTITLE__</span></div>
  <div class="spacer"></div>
  <a class="phone" id="mainphone" href="#"><span class="lbl"></span><span class="num"></span></a>
  <button id="toggle-contact">Contact</button>
  <div id="header-slots" style="display:flex;align-items:center;gap:10px"></div>
</header>

<div class="contactbar" id="contactbar" hidden></div>

<div class="layout">
  <div class="sheetbar" id="rail"></div>
  <div class="stage" id="stage"></div>
</div>

<div class="toast" id="toast"></div>
"""

# The shell's own behaviour, plus the registry every section hands itself to.
# A section's JS never sees `document` for its own elements: it is called with
# (ROOT, DATA, SHELL) and uses $/$$ off SHELL, which resolve inside its root
# against its namespace.  That is the whole of the collision fix.
SHELL_JS = r"""
(function () {
  "use strict";

  var CONFIG = JSON.parse(document.getElementById("shell-config").textContent);
  var CONTACT = CONFIG.contact || {};
  var toastEl = document.getElementById("toast");
  var KEY = CONFIG.key || "polypack";

  function toast(msg) {
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(toastEl._t);
    toastEl._t = setTimeout(function () { toastEl.classList.remove("show"); }, 1700);
  }
  function store(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) {}
  }
  function load(key, fallback) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (e) { return fallback; }
  }
  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined) n.textContent = text;
    return n;
  }

  // ---- masthead + contact ------------------------------------------------
  var brand = document.getElementById("brand");
  brand.appendChild(document.createTextNode(CONTACT.company || "POLYPACK"));
  if (CONTACT.tagline) brand.appendChild(el("small", "", CONTACT.tagline));

  /* One phone or many, one email or many - normalised to a list so everything
     below reads one shape.  See the note at the drawer. */
  function many(one, list) {
    if (list && list.length) return list;
    return one ? [one] : [];
  }

  /* WHICH office the masthead advertises is a choice, not a side effect of the
     drawer's order.  The quote page puts Parts up there on purpose - it is a
     parts RFQ - while the drawer reads headquarters-first, and before `masthead`
     existed reordering the drawer silently retitled the masthead.  Falls back to
     the first line, which is what every config written before this did. */
  var lines = CONTACT.lines || [];
  var primary = lines.filter(function (c) { return c.masthead; })[0]
                || lines[0] || {};
  var primaryPhones = many(primary.phone, primary.phones);
  var mainPhone = document.getElementById("mainphone");
  if (primaryPhones.length) {
    /* The FIRST number of the first block - a switchboard, not a fax. */
    var pn = primaryPhones[0].number || primaryPhones[0];
    mainPhone.href = primaryPhones[0].tel || primary.tel
                     || ("tel:" + String(pn).replace(/[^\d+]/g, ""));
    mainPhone.querySelector(".lbl").textContent = primary.label || "Call";
    mainPhone.querySelector(".num").textContent = pn;
  } else {
    mainPhone.hidden = true;
  }

  /* A block is a PLACE you can reach - headquarters, service, parts - so it
     carries however many numbers and addresses that place has.  One phone and
     one email per block could not say "main line AND fax", and the address was
     stranded in a block of its own captioned "Address" rather than sitting
     under the office it belongs to.  Both shapes are accepted: `phone`/`email`
     singular is normalised by many(), so a config written for the old shape
     (and the app's own, which this file's JS is synced into) still renders. */
  var bar = document.getElementById("contactbar");
  lines.forEach(function (c) {
    var block = el("div", "cblock");
    block.appendChild(el("div", "cl", c.label || ""));
    if (c.address) block.appendChild(el("div", "addr", c.address));
    many(c.phone, c.phones).forEach(function (p, i) {
      var num = p.number || p;
      var a = el("a", "", p.note ? num + "  " + p.note : num);
      a.href = p.tel || (i === 0 && c.tel ? c.tel
                                          : "tel:" + String(num).replace(/[^\d+]/g, ""));
      block.appendChild(a);
    });
    many(c.email, c.emails).forEach(function (e) {
      var m = el("a", "", e);
      m.href = "mailto:" + e;
      block.appendChild(m);
    });
    if (c.website) {
      var lw = el("a", "", c.website.replace(/^https?:\/\//, "").replace(/\/$/, ""));
      lw.href = c.website;
      lw.target = "_blank";
      block.appendChild(lw);
    }
    bar.appendChild(block);
  });
  /* A top-level address is the OLD shape - no line claimed it, so it still gets
     its own block rather than being dropped. */
  if (CONTACT.address && !lines.some(function (c) { return c.address; })) {
    var ab = el("div", "cblock");
    ab.appendChild(el("div", "cl", "Address"));
    ab.appendChild(el("div", "addr", CONTACT.address));
    bar.appendChild(ab);
  }
  if (CONTACT.website && !lines.some(function (c) { return c.website; })) {
    var wb = el("div", "cblock");
    wb.appendChild(el("div", "cl", "Online"));
    var w = el("a", "", CONTACT.website.replace(/^https?:\/\//, "").replace(/\/$/, ""));
    w.href = CONTACT.website;
    w.target = "_blank";
    wb.appendChild(w);
    bar.appendChild(wb);
  }
  /* Measure what is pinned to the top, rather than trusting the CSS defaults:
     the header wraps onto a second line on a narrow window, and the drawer's
     height depends on how many offices are configured.  --headh is the header
     alone (the drawer parks there); --stickh adds the drawer when it is open
     (the rail parks there, so it is never hidden behind the drawer). */
  function headerHeight() {
    var h = document.querySelector("header");
    if (!h) return;
    var head = h.offsetHeight;
    var root = document.documentElement.style;
    root.setProperty("--headh", head + "px");
    root.setProperty("--stickh",
                     (head + (bar.hidden ? 0 : bar.offsetHeight)) + "px");
  }
  headerHeight();
  window.addEventListener("resize", headerHeight);

  var toggleContact = document.getElementById("toggle-contact");
  toggleContact.onclick = function () {
    bar.hidden = !bar.hidden;
    toggleContact.classList.toggle("on", !bar.hidden);
    headerHeight();   /* the drawer just changed how much is pinned */
  };

  /* ---- the rail + the stage ---------------------------------------------
     One group per section.  With ONE section the group button is hidden and
     only its nav rows show, so a single-section deliverable (the drawing page
     on its own) looks exactly as it always did: a bare list in the left rail. */
  var SECTIONS = CONFIG.sections || [];
  var rail = document.getElementById("rail");
  var stage = document.getElementById("stage");
  var solo = SECTIONS.length < 2;
  var active = null;
  var groups = {};

  rail.classList.toggle("solo", solo);
  if (!solo && CONFIG.rail_head) {
    rail.appendChild(el("div", "barhead", CONFIG.rail_head));
  }

  /* The section roots are written at BODY level - the page template plants the
     section-HTML block after the whole layout block, and the manual appends its
     own the same way.  (Spelled out rather than written as the placeholder
     token: SHELL_JS is substituted into the page AFTER that token has already
     been consumed, so a literal one here survives into the built page and every
     build looks half-rendered.)  The stage is where they belong.  Left where they land,
     `.stage` stays empty and every section renders underneath the shell
     instead of beside the rail, which reads as the drawing being way down the
     bottom of the page.  Claim them here, before anything is shown.
     Idempotent, so a caller that already nests them correctly is unaffected. */
  if (stage) {
    SECTIONS.forEach(function (s) {
      var root = document.getElementById("sec-" + s.id);
      if (root && root.parentNode !== stage) stage.appendChild(root);
    });
  }

  /* A one-section page whose section has nothing to navigate has no rail at
     all - the old page collapsed its empty sheet switcher the same way, and a
     bare 176px gutter beside the drawing is just lost width.  Re-checked after
     each section registers, because a section may fill its own rows then. */
  function refreshRail() {
    if (!solo) return;
    var any = false;
    Object.keys(groups).forEach(function (k) {
      if (groups[k].subs.children.length) any = true;
    });
    rail.hidden = !any;
  }

  /* Sections carrying the same `group` share one rail row - Electrical holding
     its schematics, Procedures holding each Word export - so a manual with
     twenty pages still opens as a short list.  A group row is a row, not a
     caption: clicking it opens the first section under it. */
  var bands = {};

  function railHost(s) {
    var g = s.group;
    if (!g || !g.length) return rail;
    var key = String(g[1] || g[0]);
    if (!bands[key]) {
      var band = el("div", "navband");
      var head = el("button", "bandhead");
      head.appendChild(el("span", "lbl", (g[0] ? g[0] + "  " : "") + (g[1] || "")));
      head.appendChild(el("span", "chev", "▸"));
      head.onclick = function () { show(bands[key].first); };
      band.appendChild(head);
      var kids = el("div", "navkids");
      band.appendChild(kids);
      rail.appendChild(band);
      bands[key] = { band: band, head: head, kids: kids, first: s.id, ids: [] };
    }
    bands[key].ids.push(s.id);
    return bands[key].kids;
  }

  SECTIONS.forEach(function (s) {
    var group = el("div", "navgroup");
    group.dataset.sec = s.id;

    var btn = el("button", "");
    btn.appendChild(el("span", "lbl", (s.icon ? s.icon + "  " : "") + s.label));
    btn.onclick = function () { show(s.id); };
    if (solo) btn.hidden = true;
    group.appendChild(btn);

    var subs = el("div", "navsub");
    subs.dataset.subsFor = s.id;
    (s.nav || []).forEach(function (row) {
      var b = el("button", row.deep ? "deep" : "");
      b.textContent = row.label;
      b.onclick = function () { show(s.id, row.id); };
      subs.appendChild(b);
    });
    group.appendChild(subs);
    railHost(s).appendChild(group);
    groups[s.id] = { group: group, btn: btn, subs: subs, band: s.group ?
      String(s.group[1] || s.group[0]) : "" };
  });

  /* A section owns everything inside its root, so its markup is planted here
     and its init is called with that root - never with `document`. */
  var API = {};
  var pending = {};

  function show(id, anchor) {
    if (!groups[id]) return;
    active = id;
    SECTIONS.forEach(function (s) {
      var on = s.id === id;
      var g = groups[s.id];
      g.group.classList.toggle("on", on || solo);
      g.btn.classList.toggle("on", on);
      var root = document.getElementById("sec-" + s.id);
      if (root) root.classList.toggle("on", on);
      // the section's header controls come and go with it
      var slot = document.querySelector('[data-slot-for="' + s.id + '"]');
      if (slot) slot.hidden = !on;
    });
    // a group opens only while one of its own sections is the open one
    var band = groups[id].band;
    Object.keys(bands).forEach(function (k) {
      bands[k].band.classList.toggle("on", k === band);
      bands[k].head.classList.toggle("on", k === band);
    });
    if (anchor) {
      var target = document.getElementById(anchor);
      if (target && target.scrollIntoView) {
        target.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }
    markRail();
  }

  function markRail() {
    var g = groups[active];
    if (!g) return;
    if (API[active] && API[active].railMark) API[active].railMark(g.subs);
  }

  window.POLYPACK = {
    version: CONFIG.version,
    toast: toast, store: store, load: load, el: el,
    key: KEY,
    show: show,
    markRail: markRail,
    /* A section registers by id.  `ns` is its id-prefix; $ and $$ are the only
       way its code should ever reach the DOM - both stay inside its own root,
       so two sections carrying the same id never see each other. */
    register: function (id, init) {
      var root = document.getElementById("sec-" + id);
      if (!root) return;
      var ns = id + "-";
      var shell = {
        root: root, ns: ns, id: id,
        toast: toast, store: store, load: load, el: el,
        key: KEY + ":" + id,
        $: function (rest) { return document.getElementById(ns + rest); },
        $$: function (sel) { return root.querySelectorAll(sel); },
        sel: function (rest) { return "#" + ns + rest; },
        /* Chrome the section wants in the header (the drawing's Request Quote
           button and its Align sliders) is planted here, not in the section's
           own markup, so it sits with the rest of the chrome and disappears
           with the section.

           The wrapper carries the section's ROOT CLASS - `pp-sec-ipl` but not
           `pp-sec` - which is the whole trick: the section's CSS is scoped to
           that class, so a control lifted out of its column is still reached by
           its own rules, while the shell's `.pp-sec { display: none }` does not
           blank it.  Without this, moving a button into the header silently
           stripped its styling. */
        headerSlot: function (node) {
          var host = document.getElementById("header-slots");
          var wrap = host.querySelector('[data-slot-for="' + id + '"]');
          if (!wrap) {
            wrap = el("span", "slotwrap pp-sec-" + id);
            wrap.dataset.slotFor = id;
            // sections register AFTER the first show(), so a wrapper born now
            // has to work out for itself whether its section is the open one
            wrap.hidden = id !== active;
            host.appendChild(wrap);
          }
          wrap.appendChild(node);
          return node;
        },
        railSubs: function () {
          return document.querySelector('[data-subs-for="' + id + '"]');
        },
        solo: solo
      };
      var payload = JSON.parse(
        document.getElementById("payload-" + id).textContent);
      API[id] = init(root, payload, shell) || {};
      refreshRail();
      if (pending[id]) { pending[id](); delete pending[id]; }
    },
    api: function (id) { return API[id]; }
  };

  // open the first section (and honour a #sec/anchor deep link)
  var hash = (location.hash || "").replace(/^#/, "");
  var opened = false;
  if (hash) {
    var target = document.getElementById(hash);
    var host = target && target.closest ? target.closest(".pp-sec") : null;
    if (host) { show(host.dataset.sec, hash); opened = true; }
  }
  if (!opened && SECTIONS.length) show(SECTIONS[0].id);
  refreshRail();
})();
"""

PAGE_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
__SHELL_CSS__
__SECTION_CSS__
</style>
</head>
<body class="__BODY_CLASS__">
__SHELL_HTML__
__SECTION_HTML__
<script id="shell-config" type="application/json">__SHELL_CONFIG__</script>
__PAYLOADS__
<script>
__SHELL_JS__
</script>
__SECTION_JS__
</body>
</html>
"""


# --------------------------------------------------------------------------
# Namespacing
# --------------------------------------------------------------------------

_ID_ATTR = re.compile(r'\b(id|for|aria-labelledby|aria-controls)="([^"]*)"')
_HREF_HASH = re.compile(r'\bhref="#([^"]*)"')


def namespace_html(html: str, ns: str) -> str:
    """Prefix every id/for/href="#..." in a section's markup with its namespace.

    The section's own JS never hardcodes these - it goes through SHELL.$ , which
    applies the same prefix - so the two stay in step by construction.  Only
    attributes are touched: text content and class names are left exactly as
    written, because CSS is scoped separately (scope_css) and class collisions
    between sections are harmless once the selectors are rooted.
    """
    if not ns:
        return html
    html = _ID_ATTR.sub(
        lambda m: '%s="%s"' % (m.group(1),
                               " ".join(ns + p for p in m.group(2).split() if p)
                               if m.group(1).startswith("aria")
                               else ns + m.group(2)),
        html)
    return _HREF_HASH.sub(lambda m: 'href="#%s%s"' % (ns, m.group(1)), html)


# `body` / `html` only count as the element when what follows them cannot be
# part of an identifier - so `body.show-hot` matches and `bodyguard` does not.
_LEADING_BODY = re.compile(r"^body(?![\w-])")
_LEADING_HTML = re.compile(r"^html(?![\w-])")


# Selector-list splitting has to respect commas inside :is(), [attr=","] and
# the like, so it is done by hand rather than with str.split(",").
def _split_selectors(chunk: str) -> list[str]:
    out, depth, quote, buf = [], 0, "", []
    for ch in chunk:
        if quote:
            if ch == quote:
                quote = ""
            buf.append(ch)
            continue
        if ch in "\"'":
            quote = ch
        elif ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        elif ch == "," and depth == 0:
            out.append("".join(buf))
            buf = []
            continue
        buf.append(ch)
    out.append("".join(buf))
    return [s for s in (p.strip() for p in out) if s]


def _prefix_one(selector: str, root: str) -> str:
    """`.hot:hover` -> `.pp-sec-ipl .hot:hover`, and the root itself stays put.

    A selector that already starts at the section root (`&`, or the root class
    written out) is anchored rather than nested - otherwise `.pp-sec-ipl.pdf-mode`
    would become `.pp-sec-ipl .pp-sec-ipl.pdf-mode` and match nothing.
    """
    selector = selector.strip()
    if selector.startswith("&"):
        return root + selector[1:].strip()
    if _LEADING_HTML.match(selector):
        return selector                       # nothing above the root to scope
    # `body`, `body.show-hot`, `body:not(.x) .y` - the section root IS the body
    # as far as its own CSS is concerned, so the leading `body` is replaced
    # rather than nested under.
    lead = _LEADING_BODY.match(selector)
    if lead:
        return root + selector[lead.end():]
    if selector.startswith(root):
        return selector
    return root + " " + selector


def scope_css(css: str, root_class: str) -> str:
    """Root every rule in a section's CSS under `.pp-sec-<id>`.

    Not a CSS parser - a brace walker.  It only has to survive the CSS this repo
    writes: rules, comments, and @media/@supports blocks whose contents are
    themselves rules.  @font-face / @keyframes carry no selectors to scope, so
    they are copied through whole.
    """
    root = "." + root_class
    out, i, n = [], 0, len(css)
    buf = []

    def flush_rule(selector_text: str, body: str) -> None:
        sels = _split_selectors(selector_text)
        if not sels:
            return
        out.append(", ".join(_prefix_one(s, root) for s in sels))
        out.append(" {" + body + "}\n")

    while i < n:
        ch = css[i]
        if ch == "/" and i + 1 < n and css[i + 1] == "*":
            end = css.find("*/", i + 2)
            end = n if end < 0 else end + 2
            out.append(css[i:end])
            i = end
            continue
        if ch == "{":
            selector_text = "".join(buf).strip()
            buf = []
            depth, j = 1, i + 1
            while j < n and depth:
                if css[j] == "{":
                    depth += 1
                elif css[j] == "}":
                    depth -= 1
                j += 1
            body = css[i + 1:j - 1]
            if selector_text.startswith("@"):
                head = selector_text.split()[0].lower()
                if head in ("@media", "@supports", "@container", "@layer"):
                    out.append(selector_text + " {\n")
                    out.append(scope_css(body, root_class))
                    out.append("}\n")
                else:
                    out.append(selector_text + " {" + body + "}\n")   # @keyframes &c
            else:
                flush_rule(selector_text, body)
            i = j
            continue
        buf.append(ch)
        i += 1
    tail = "".join(buf).strip()
    if tail:
        out.append(tail)
    return "".join(out)


# --------------------------------------------------------------------------
# Sections
# --------------------------------------------------------------------------

_ID_OK = re.compile(r"^[a-z][a-z0-9-]*$")


def section(id: str, label: str, html: str, *, icon: str = "", css: str = "",
            js: str = "", payload: dict | None = None,
            nav: list | None = None, slots: dict | None = None,
            group: list | None = None) -> dict:
    """Build one section manifest.  `id` is the namespace, so it is checked.

    `group` is an optional [icon, label] pair: sections sharing one fold
    together under a single rail row.
    """
    if not _ID_OK.match(id or ""):
        raise ValueError(
            "section id must be lower-case letters, digits and dashes, "
            "starting with a letter (got %r)" % (id,))
    return {
        "format": FORMAT,
        "version": SCHEMA_VERSION,
        "id": id,
        "label": label,
        "icon": icon,
        "css": css,
        "html": html,
        "js": js,
        "payload": payload if payload is not None else {},
        "nav": nav or [],
        "slots": slots or {},
        "group": list(group) if group else [],
    }


_WIN_HIDDEN = 0x02
_WIN_NORMAL = 0x80


def set_hidden(path: str, on: bool = True) -> bool:
    """Set or clear Windows' hidden attribute.  True if it took.

    A manifest is written BESIDE the deliverable it belongs to - the drawing's
    page, the manual's .docx - because that is where the thing that reads it
    looks for it, and it cannot be moved somewhere tidier without breaking that.
    Hidden is the compromise: the publish folder shows only the files a person
    opens, while os.listdir - which is how the API enumerates manifests - still
    returns it.  A no-op off Windows, and never fatal: a manifest that is merely
    visible has lost nothing.
    """
    if os.name != "nt":
        return False
    try:
        import ctypes
        attr = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if attr == -1:
            return False
        want = (attr | _WIN_HIDDEN) if on else (attr & ~_WIN_HIDDEN)
        if want == attr:
            return True
        return bool(ctypes.windll.kernel32.SetFileAttributesW(
            str(path), want or _WIN_NORMAL))
    except Exception:
        return False


def write_section(path: str, sec: dict, hidden: bool = False) -> str:
    """Write one manifest.  `hidden` keeps it out of the publish folder's view.

    The attribute is cleared BEFORE the write whatever `hidden` says: on Windows
    a truncating open() of an existing HIDDEN file fails outright with access
    denied (CreateFile CREATE_ALWAYS refuses to hand a hidden file back as a
    normal one), so a second run over the same drawing would die on a file the
    first run hid.  Clear, write, re-hide.
    """
    if os.path.exists(path):
        set_hidden(path, False)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(sec, handle, ensure_ascii=False)
    if hidden:
        set_hidden(path, True)
    return path


def load_section(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as handle:
        sec = json.load(handle)
    if sec.get("format") != FORMAT:
        raise ValueError("%s is not a %s manifest" % (path, FORMAT))
    if int(sec.get("version", 0)) > SCHEMA_VERSION:
        raise ValueError(
            "%s is schema v%s; this build understands up to v%s - update "
            "POLYPACK_SHELL.py" % (path, sec.get("version"), SCHEMA_VERSION))
    return sec


def tel_href(display: str) -> str:
    """(727) 578-5000 x245  ->  tel:+17275785000;ext=245"""
    found = re.match(
        r"\D*(\d{3})\D*(\d{3})\D*(\d{4})(?:\s*(?:x|ext\.?)\s*(\d+))?",
        display or "", re.IGNORECASE)
    if not found:
        return ""
    number = "tel:+1" + found.group(1) + found.group(2) + found.group(3)
    return number + (";ext=" + found.group(4) if found.group(4) else "")


def prepare_contact(contact: dict) -> dict:
    """Fill in every tel: from its printed phone number.

    Handles both shapes - a line's single `phone`, and a `phones` list whose
    entries are either bare strings or {number, note} - because the extension
    numbers only dial correctly if tel_href has parsed the printed form.  A
    hand-written config should never have to spell out `tel:+17275785000;ext=245`.
    """
    out = dict(contact or {})
    out["lines"] = [dict(line) for line in out.get("lines", [])]
    for line in out["lines"]:
        if line.get("phone") and not line.get("tel"):
            line["tel"] = tel_href(line["phone"])
        if line.get("phones"):
            phones = []
            for entry in line["phones"]:
                p = {"number": entry} if isinstance(entry, str) else dict(entry)
                if p.get("number") and not p.get("tel"):
                    p["tel"] = tel_href(p["number"])
                phones.append(p)
            line["phones"] = phones
    return out


def _json_for_html(value) -> str:
    """JSON safe to sit inside a <script> block."""
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")


def render_page(title: str, sections: list, *, subtitle: str = "",
                contact: dict | None = None, key: str = "polypack",
                rail_head: str = "", body_class: str = "") -> str:
    """Pour any number of sections into one self-contained page."""
    if not sections:
        raise ValueError("a page needs at least one section")
    seen = set()
    for sec in sections:
        if sec["id"] in seen:
            raise ValueError(
                "two sections share the id %r - ids are namespaces, so they "
                "must be unique within one page" % sec["id"])
        seen.add(sec["id"])

    css_parts, html_parts, payload_parts, js_parts = [], [], [], []
    for sec in sections:
        sid = sec["id"]
        ns = sid + "-"
        root_class = "pp-sec-" + sid
        if sec.get("css"):
            css_parts.append("\n/* ===== %s ===== */\n" % sid)
            css_parts.append(scope_css(sec["css"], root_class))
        html_parts.append(
            '\n<div class="pp-sec %s" id="sec-%s" data-sec="%s">\n%s\n</div>'
            % (root_class, sid, sid, namespace_html(sec.get("html", ""), ns)))
        payload_parts.append(
            '<script id="payload-%s" type="application/json">%s</script>'
            % (sid, _json_for_html(sec.get("payload", {}))))
        if sec.get("js"):
            # registered, not run: the shell hands it its root and its payload
            js_parts.append(
                '<script>\nPOLYPACK.register("%s", function (ROOT, DATA, SHELL) {\n'
                'var $ = SHELL.$, $$ = SHELL.$$, NS = SHELL.ns;\n%s\n});\n</script>'
                % (sid, sec["js"]))

    config = {
        "version": SCHEMA_VERSION,
        "key": key,
        "rail_head": rail_head,
        "contact": prepare_contact(contact or {}),
        "sections": [
            {"id": s["id"], "label": s["label"], "icon": s.get("icon", ""),
             "nav": s.get("nav", []), "slots": s.get("slots", {}),
             "group": s.get("group", [])}
            for s in sections
        ],
    }

    return (PAGE_TEMPLATE
            .replace("__SHELL_CSS__", SHELL_CSS)
            .replace("__SECTION_CSS__", "".join(css_parts))
            .replace("__SHELL_HTML__",
                     SHELL_HTML.replace("__TITLE__", _esc(title))
                               .replace("__SUBTITLE__", _esc(subtitle)))
            .replace("__SECTION_HTML__", "".join(html_parts))
            .replace("__SHELL_CONFIG__", _json_for_html(config))
            .replace("__PAYLOADS__", "\n".join(payload_parts))
            .replace("__SHELL_JS__", SHELL_JS)
            .replace("__SECTION_JS__", "\n".join(js_parts))
            .replace("__BODY_CLASS__", body_class)
            .replace("__TITLE__", _esc(title)))


def _esc(text: str) -> str:
    return (str(text or "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def render_from_files(title: str, paths: list, **kw) -> str:
    """Merge N manifests off disk, in the order given."""
    return render_page(title, [load_section(p) for p in paths], **kw)

if __name__ == "__main__":
    sys.exit(main())
