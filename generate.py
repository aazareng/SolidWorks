"""
generate.py — Polypack Manual Generation Script
================================================
Usage:
    python generate.py --config templates/machine_config_template.xlsx

Requirements:
    pip install python-pptx python-docx openpyxl Pillow
    LibreOffice must be installed for PPTX → image conversion:
        Linux:   sudo apt install libreoffice
        macOS:   brew install --cask libreoffice
        Windows: https://www.libreoffice.org/download/

Folder layout expected:
    /
    ├── generate.py
    ├── template.docx                  ← master Word template
    ├── templates/
    │   ├── machine_config_template.xlsx
    │   └── module_template.pptx
    ├── modules/
    │   ├── infeed_conveyor/
    │   │   └── slides.pptx            ← filled-in copy of module_template.pptx
    │   ├── seal_bar/
    │   │   └── slides.pptx
    │   └── ...
    └── output/                        ← generated manuals land here

PPTX slide conventions (read from slide notes):
    SKIP=TRUE           → divider / instructions slide, ignored
    TYPE=OVERVIEW       → exported as Section 5 image
    TYPE=CHANGEOVER     → exported as Section 7 image; notes body = step text
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DInches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

TEMPLATE_DOCX = Path("template.docx")
MODULES_DIR   = Path("modules")
OUTPUT_DIR    = Path("output")

MODULE_PLACEHOLDER = "xXxMODULExXx"

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG LOADING  (reads machine_config_template.xlsx)
# ─────────────────────────────────────────────────────────────────────────────

def load_config(xlsx_path: Path) -> dict:
    """
    Returns a dict with keys:
        machine   : {CUSTOMER, MODEL, SERIAL NUMBER, MANUAL DATE, YEAR (YY), ...film specs...}
        modules   : list of {order, folder, display_name, include, co_include, description}
        bom_sheets: {folder_name: [{ITEM #, PART NUMBER, DESCRIPTION, QTY, MATERIAL, VENDOR, NOTES}]}
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # ── MACHINE INFO sheet ────────────────────────────────────────────────────
    ws = wb["MACHINE INFO"]
    machine = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[0] and row[1] is not None:
            machine[str(row[0]).strip()] = str(row[1]).strip() if row[1] else ""

    # ── MODULES sheet ─────────────────────────────────────────────────────────
    wm = wb["MODULES"]
    modules = []
    for row in wm.iter_rows(min_row=4, values_only=True):
        if not row[0]:          # empty order cell → end of list
            continue
        order, folder, display, include, co_inc, desc = (row + (None,) * 6)[:6]
        if folder:
            modules.append({
                "order":        int(order) if order else 99,
                "folder":       str(folder).strip(),
                "display_name": str(display).strip() if display else str(folder).strip(),
                "include":      str(include).strip().upper() == "TRUE",
                "co_include":   str(co_inc).strip().upper() == "TRUE" if co_inc else False,
                "description":  str(desc).strip() if desc else "",
            })
    modules.sort(key=lambda m: m["order"])

    # ── BOM sheets (any sheet whose name starts with "BOM") ───────────────────
    bom_data = {}
    for sheet_name in wb.sheetnames:
        if not sheet_name.upper().startswith("BOM"):
            continue
        ws_bom = wb[sheet_name]
        # Row 3, col 2 = module folder name
        folder_key = ws_bom.cell(row=3, column=2).value
        if not folder_key:
            continue
        folder_key = str(folder_key).strip()
        rows = []
        for row in ws_bom.iter_rows(min_row=6, values_only=True):
            if not any(row):
                continue
            rows.append({
                "item":        row[0] or "",
                "part_number": row[1] or "",
                "description": row[2] or "",
                "qty":         row[3] or "",
                "material":    row[4] or "",
                "vendor":      row[5] or "",
                "notes":       row[6] or "",
            })
        bom_data[folder_key] = rows

    return {"machine": machine, "modules": modules, "bom": bom_data}


# ─────────────────────────────────────────────────────────────────────────────
# PPTX → IMAGES  (via LibreOffice headless)
# ─────────────────────────────────────────────────────────────────────────────

def _libreoffice_export_slides(pptx_path: Path, out_dir: Path) -> list[Path]:
    """
    Calls LibreOffice headless to export each slide of pptx_path as a PNG.
    Returns list of PNG paths sorted by slide number.
    """
    cmd = [
        "libreoffice", "--headless", "--convert-to", "png",
        "--outdir", str(out_dir), str(pptx_path)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice failed on {pptx_path}:\n{result.stderr}"
        )
    # LibreOffice names them  <stem>1.png, <stem>2.png, ...
    stem = pptx_path.stem
    pngs = sorted(out_dir.glob(f"{stem}*.png"),
                  key=lambda p: int(re.search(r"(\d+)$", p.stem).group(1)))
    return pngs


def _get_slide_meta(pptx_path: Path) -> list[dict]:
    """
    Parse each slide's notes for TYPE= and SKIP= directives.
    Returns list of {index, type, tag, notes_body, skip}.
    """
    prs = Presentation(str(pptx_path))
    meta = []
    for i, slide in enumerate(prs.slides):
        notes_text = ""
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text or ""

        skip = "SKIP=TRUE" in notes_text.upper()
        slide_type = ""
        tag = ""
        notes_body = []

        for line in notes_text.splitlines():
            line = line.strip()
            if line.upper().startswith("TYPE="):
                slide_type = line.split("=", 1)[1].strip().upper()
            elif line.upper().startswith("TAG="):
                tag = line.split("=", 1)[1].strip()
            elif not line.upper().startswith("SKIP="):
                notes_body.append(line)

        # Remove leading blank lines and the template prompt text
        while notes_body and (not notes_body[0] or notes_body[0].startswith("---")):
            notes_body.pop(0)

        meta.append({
            "index":      i,
            "type":       slide_type,
            "tag":        tag,
            "notes_body": "\n".join(notes_body).strip(),
            "skip":       skip,
        })
    return meta


def extract_module_images(module_folder: Path, tmp_root: Path) -> dict:
    """
    Given a module folder containing slides.pptx, returns:
        {
          "overview":    [Path, Path, ...],   # up to 4, in slide order
          "changeover":  [(Path, str), ...],  # (image_path, step_text)
        }
    Raises FileNotFoundError if slides.pptx is missing.
    """
    pptx_path = module_folder / "slides.pptx"
    if not pptx_path.exists():
        raise FileNotFoundError(f"Missing slides.pptx in {module_folder}")

    tmp_dir = tmp_root / module_folder.name
    tmp_dir.mkdir(parents=True, exist_ok=True)

    png_paths = _libreoffice_export_slides(pptx_path, tmp_dir)
    meta      = _get_slide_meta(pptx_path)

    if len(png_paths) != len(meta):
        raise RuntimeError(
            f"Slide count mismatch for {pptx_path}: "
            f"{len(meta)} slides in PPTX but {len(png_paths)} PNGs exported."
        )

    overview   = []
    changeover = []

    for m, png in zip(meta, png_paths):
        if m["skip"]:
            continue
        if m["type"] == "OVERVIEW":
            overview.append(png)
        elif m["type"] == "CHANGEOVER":
            changeover.append((png, m["notes_body"]))

    return {"overview": overview[:4], "changeover": changeover}


# ─────────────────────────────────────────────────────────────────────────────
# WORD DOCUMENT ASSEMBLY
# ─────────────────────────────────────────────────────────────────────────────

def replace_text_in_doc(doc: Document, old: str, new: str):
    """Replace all occurrences of old with new across paragraphs and table cells."""
    for para in doc.paragraphs:
        if old in para.text:
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if old in para.text:
                        for run in para.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, new)


def insert_2x2_grid(doc: Document, images: list[Path], caption: str = ""):
    """
    Insert a 2×2 table of images (comic-book layout) into doc at current position.
    images: list of 1–4 image paths. Missing slots are left blank.
    """
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"

    cells = [table.cell(r, c) for r in range(2) for c in range(2)]
    for i, cell in enumerate(cells):
        if i < len(images):
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(str(images[i]), width=DInches(3.0))

    if caption:
        cap_para = doc.add_paragraph(caption)
        cap_para.style = "Caption1"


def insert_changeover_steps(doc: Document, steps: list[tuple]):
    """
    steps: list of (image_path, step_text)
    Inserts numbered heading + image + step text for each changeover step.
    """
    for i, (img_path, step_text) in enumerate(steps, 1):
        h = doc.add_paragraph(f"Step {i}", style="Heading2")
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(str(img_path), width=DInches(5.5))
        if step_text:
            doc.add_paragraph(step_text, style="NoSpacing")


def insert_bom_table(doc: Document, bom_rows: list[dict]):
    """Insert a Bill of Materials table."""
    headers = ["ITEM #", "PART NUMBER", "DESCRIPTION", "QTY",
               "MATERIAL", "VENDOR", "NOTES"]
    table = doc.add_table(rows=1 + len(bom_rows), cols=len(headers))
    table.style = "Table Grid"

    hdr_row = table.rows[0]
    for ci, h in enumerate(headers):
        cell = hdr_row.cells[ci]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True

    for ri, bom in enumerate(bom_rows):
        data_row = table.rows[ri + 1]
        vals = [bom["item"], bom["part_number"], bom["description"],
                bom["qty"], bom["material"], bom["vendor"], bom["notes"]]
        for ci, val in enumerate(vals):
            data_row.cells[ci].text = str(val)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN GENERATION PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def generate(config_path: Path):
    config = load_config(config_path)
    machine = config["machine"]
    modules = [m for m in config["modules"] if m["include"]]

    # ── Determine output filename ─────────────────────────────────────────────
    yy     = machine.get("YEAR (YY)", "XX")
    serial = machine.get("SERIAL NUMBER", "0000").replace("-", "")
    sn_short = serial[-4:] if len(serial) >= 4 else serial
    customer = machine.get("CUSTOMER", "CUSTOMER").upper()
    model    = machine.get("MODEL", "MODEL").upper()
    out_name = f"{yy}-{sn_short} - {customer} - {model}.docx"
    out_path = OUTPUT_DIR / out_name
    OUTPUT_DIR.mkdir(exist_ok=True)

    print(f"\n  Output : {out_path}")
    print(f"  Modules: {[m['folder'] for m in modules]}\n")

    # ── Clone template ────────────────────────────────────────────────────────
    if not TEMPLATE_DOCX.exists():
        sys.exit(f"ERROR: {TEMPLATE_DOCX} not found.")
    shutil.copy(TEMPLATE_DOCX, out_path)
    doc = Document(str(out_path))

    # ── Replace cover + header placeholders ───────────────────────────────────
    replacements = {
        "CUSTOMER":           machine.get("CUSTOMER", "CUSTOMER"),
        "MODEL: MODEL":       f"MODEL: {machine.get('MODEL', 'MODEL')}",
        "SERIAL: 2X-4XXX":   f"SERIAL: {machine.get('SERIAL NUMBER', '??-????')}",
        "MMMM-YYYY":          machine.get("MANUAL DATE", ""),
    }
    for old, new in replacements.items():
        replace_text_in_doc(doc, old, new)

    # ── Populate ComboBox content controls (film specs) ───────────────────────
    # TODO: iterate doc.element.body to find w:sdt tags and set w:t values
    # for the 6 film spec controls. (Requires direct XML manipulation.)

    # ── Process modules — find xXxMODULExXx paragraphs and expand ─────────────
    with tempfile.TemporaryDirectory() as tmp_root:
        tmp_root = Path(tmp_root)

        # Collect all overview images + changeover steps per module
        module_data = {}
        for m in modules:
            folder = MODULES_DIR / m["folder"]
            if not folder.exists():
                print(f"  WARNING: module folder not found: {folder} — skipping.")
                continue
            try:
                imgs = extract_module_images(folder, tmp_root)
                module_data[m["folder"]] = {**m, **imgs}
                print(f"  [{m['folder']}] overview={len(imgs['overview'])}  "
                      f"changeover={len(imgs['changeover'])}")
            except FileNotFoundError as e:
                print(f"  WARNING: {e} — skipping.")

        # Walk paragraphs and replace MODULE placeholders
        # Each xXxMODULExXx occurrence marks one module slot (in order).
        module_iter = iter([d for d in module_data.values()])
        para_list   = list(doc.paragraphs)

        for para in para_list:
            if MODULE_PLACEHOLDER not in para.text:
                continue
            try:
                m = next(module_iter)
            except StopIteration:
                # More placeholders than modules — clear remaining
                for run in para.runs:
                    run.text = ""
                continue

            # Clear the placeholder paragraph (we'll insert after it)
            for run in para.runs:
                run.text = ""

            # Insert section heading
            # (In practice we'd insert before/after para using docx XML tricks;
            #  for now we append to end of doc as a structural stub.)
            doc.add_heading(m["display_name"], level=1)

            # 2×2 image grid
            if m["overview"]:
                insert_2x2_grid(doc, m["overview"], caption=m.get("description", ""))
            elif m.get("description"):
                doc.add_paragraph(m["description"], style="NoSpacing")

            # Changeover
            if m.get("co_include") and m.get("changeover"):
                doc.add_heading("Changeover Procedure", level=2)
                insert_changeover_steps(doc, m["changeover"])

            # BOM
            bom_rows = config["bom"].get(m["folder"], [])
            if bom_rows:
                doc.add_heading("Bill of Materials", level=2)
                insert_bom_table(doc, bom_rows)

    doc.save(str(out_path))
    print(f"\n  Saved: {out_path}\n")


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Polypack Manual Generator")
    parser.add_argument(
        "--config", "-c",
        required=True,
        type=Path,
        help="Path to machine_config_template.xlsx (or a filled copy)"
    )
    args = parser.parse_args()
    generate(args.config)
