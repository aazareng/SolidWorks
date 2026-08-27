"""
CleanSalesDwg.py — a sales DXF with the small holes taken out
=============================================================
Usage:
    python CleanSalesDwg.py "C:\\...\\A9987.dxf" --out "C:\\publish"
    python CleanSalesDwg.py "C:\\...\\A9987.dxf" --min-dia 50

    Normally run by CleanSalesDwg.bas beside this file, which exports the
    active drawing sheet to DXF and PDF and then calls this over the DXF.
    It also runs by hand over any DXF, which is the thing to do the day a
    result looks wrong.

    --out DIR       where the cleaned DXF goes, under the same file name
                    (default: beside the source as <name>_CLEAN.dxf)
    --min-dia DIA   remove full circles under this diameter (default: 100)
    --blocks / --no-blocks
                    whether block definitions are cleaned too (default: yes)
    --quiet         only print errors

Requirements:
    pip install ezdxf

WHAT GOES.  Full circles only, under --min-dia: a CIRCLE, an ARC that comes
back round to its start, an ELLIPSE whose ratio is 1.  Everything else stays,
including anything that merely POINTS AT a hole — centre marks, centrelines,
the dimension calling it out are none of them circles.  Strip those in the
drawing if the customer's copy should not carry them.

THE DIAMETER IS IN DRAWING UNITS AND IS MEASURED ON THE SHEET.  A DXF carries
whatever units the drawing was in, so 100 is 100 mm on a metric sheet and 100
inches on an inch one; and a view at 1:2 puts a 200 hole into the file as a
100 circle.  Export the views 1:1 — which is what this is for — and sheet size
is part size.

Nothing is edited in place.  A cleaned copy is written and the export it was
read from is left alone, so a surprising result can always be compared against
what SolidWorks actually wrote.
"""

from __future__ import annotations

import argparse
import math
import os
import sys

import ezdxf

# What the macro's input box starts on, and what this uses when nobody says.
DEFAULT_MIN_DIA = 100.0

# A circle drawn at exactly the cut-off is KEPT, rather than left to floating
# point to decide.
HOLE_TOL = 1e-6


def _full_circle_diameter(entity) -> float | None:
    """The diameter of `entity` if it draws a COMPLETE circle, else None.

    CIRCLE is the ordinary case.  An ARC that comes back round to its start is
    the same thing spelled differently, and some exports write one.  An ELLIPSE
    counts only where its ratio is 1 and it closes — anything less is a hole
    seen at an angle, which is an ellipse and stays.
    """
    kind = entity.dxftype()

    if kind == "CIRCLE":
        return 2.0 * float(entity.dxf.radius)

    if kind == "ARC":
        sweep = (float(entity.dxf.end_angle) - float(entity.dxf.start_angle)) % 360.0
        if sweep < 1e-4 or sweep > 360.0 - 1e-4:
            return 2.0 * float(entity.dxf.radius)
        return None

    if kind == "ELLIPSE":
        if abs(float(entity.dxf.ratio) - 1.0) > 1e-9:
            return None
        span = abs(float(entity.dxf.end_param) - float(entity.dxf.start_param))
        if abs(span - 2.0 * math.pi) > 1e-6:
            return None
        axis = entity.dxf.major_axis
        return 2.0 * math.sqrt(axis[0] ** 2 + axis[1] ** 2 + axis[2] ** 2)

    return None


def strip_small_circles(dxf_path: str, min_diameter: float, out_path: str,
                        blocks: bool = True, log=print) -> dict:
    """Copy one DXF to `out_path` without the full circles under `min_diameter`.

    Block definitions are cleaned too unless `blocks` is False, because a hole
    drawn once and INSERTed twenty times has to go twenty times.  That same
    reach is why a balloon RING under the cut-off comes out with the holes
    while its item number does not: the ring is a circle living in a block, the
    number is text.  SolidWorks' own marker blocks — _ORIGIN, _SMALL,
    _DOTBLANK — carry small circles too and go the same way.

    None of that is hidden: the count is logged PER SPACE rather than summed,
    so a line reading `block SW_NOTE1` is exactly that happening, visible in
    the macro's report rather than discovered by a customer.

    Returns what came out, for a caller that wants to report it.
    """
    if os.path.abspath(dxf_path) == os.path.abspath(out_path):
        raise ValueError("the cleaned DXF would overwrite the file it was read "
                         "from — send it somewhere else with --out")

    doc = ezdxf.readfile(dxf_path)

    spaces = [("modelspace", doc.modelspace())]
    for name in doc.layouts.names_in_taborder():
        if name.lower() != "model":
            spaces.append((f"layout {name}", doc.layout(name)))
    if blocks:
        for block in doc.blocks:
            # *Model_Space and *Paper_Space are the layouts above wearing their
            # block names; walking them again would count every circle twice.
            if block.name.startswith("*"):
                continue
            spaces.append((f"block {block.name}", block))

    removed: list[float] = []
    kept: list[float] = []
    by_space: dict[str, int] = {}

    for label, space in spaces:
        # Collected first, deleted after — deleting inside the walk is how you
        # skip the entity that follows the one you just removed.
        doomed = []
        for entity in space:
            diameter = _full_circle_diameter(entity)
            if diameter is None:
                continue
            if diameter < min_diameter - HOLE_TOL:
                doomed.append(entity)
                removed.append(diameter)
            else:
                kept.append(diameter)
        for entity in doomed:
            space.delete_entity(entity)
        if doomed:
            by_space[label] = len(doomed)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.saveas(out_path)

    log(f"holes: {len(removed)} circles under {min_diameter:g} removed, "
        f"{len(kept)} kept -> {os.path.basename(out_path)}")
    if removed:
        biggest = f"largest removed {max(removed):g}"
        smallest = f"smallest kept {min(kept):g}" if kept else "none left above the line"
        log(f"       {biggest}, {smallest}")
    for label, count in by_space.items():
        log(f"       {count:>4}  {label}")

    return {"out": out_path,
            "removed": len(removed),
            "kept": len(kept),
            "largest_removed": max(removed) if removed else None,
            "smallest_kept": min(kept) if kept else None,
            "by_space": by_space}


def _target_for(dxf_path: str, out_dir: str | None) -> str:
    """Where the cleaned copy of `dxf_path` goes.

    With --out it keeps its name in another folder, which is what the macro
    wants: the raw export sits in %TEMP% and the clean one lands beside the
    PDF.  Without it, the name changes instead of the folder, so running this
    by hand never quietly replaces the file you pointed it at.
    """
    if out_dir:
        return os.path.join(out_dir, os.path.basename(dxf_path))
    stem, ext = os.path.splitext(os.path.abspath(dxf_path))
    return f"{stem}_CLEAN{ext}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Write a copy of a drawing DXF with the small holes removed.")
    parser.add_argument("dxf", nargs="+", help="the DXF(s) to clean")
    parser.add_argument("--out", metavar="DIR",
                        help="where the cleaned DXF goes, under the same file "
                             "name (default: beside the source, _CLEAN.dxf)")
    parser.add_argument("--min-dia", type=float, default=DEFAULT_MIN_DIA,
                        metavar="DIA",
                        help=f"remove full circles under this diameter, in "
                             f"drawing units (default: {DEFAULT_MIN_DIA:g})")
    parser.add_argument("--blocks", dest="blocks", action="store_true",
                        default=True, help="clean block definitions too (default)")
    parser.add_argument("--no-blocks", dest="blocks", action="store_false",
                        help="leave block definitions alone — keeps balloon "
                             "rings and SolidWorks' marker blocks intact")
    parser.add_argument("--quiet", action="store_true", help="only print errors")
    args = parser.parse_args(argv)

    if args.min_dia <= 0:
        print("ERROR: --min-dia has to be a positive diameter", file=sys.stderr)
        return 2

    log = (lambda *a: None) if args.quiet else print

    status = 0
    for path in args.dxf:
        try:
            strip_small_circles(path, args.min_dia, _target_for(path, args.out),
                                blocks=args.blocks, log=log)
        except Exception as exc:
            print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
            status = 1
    return status


if __name__ == "__main__":
    sys.exit(main())
