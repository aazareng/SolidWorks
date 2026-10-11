# Structural Sheet Metal — a SOLIDWORKS add-in

**Structural Member for sheet metal.** Pick a path (sketch lines or straight edges), pick a profile
from Standard / Type / Size dropdowns, pick a gauge from your gauge table, choose the corner
treatment, press OK. Every member is a real **Base Flange** from an **open** profile (an L is two
lines, a C is three), so the result is ordinary multibody sheet metal: it flattens, it uses your
gauge tables, and nothing proprietary is left in the file.

Why not Structural Member? It sweeps a *closed* region from a library feature file per size and has
no way to take an open contour, so every size needed its own thickened `.sldlfp`, then Convert to
Sheet Metal on each body. This tool skips all of that.

Status, 2026-10-11, SOLIDWORKS 2026 SP2.1:

| Piece | State |
|---|---|
| Engine (`Core\`) — planes, open profiles, base flanges, gauge table, inside/outside, miter / butt 1 / butt 2, gap, rotation, pierce point, mirror, closed loops, folders | **Verified live** from the terminal: L, U, closed rectangle (4 miters), butt joints, outside material, two groups hugging, 16 ga from the sample steel table |
| CLI (`Cli\ssm.exe`) | Works; the way to test engine changes without touching the add-in |
| Add-in (`AddIn\`) — command tab, PropertyManager page, live preview | Loads, the tab, page, preview and OK build work (user-tested 2026-10-11). Rotation box fixed (radians through the API). |

## Layout

```
StructuralSheetMetal\
  Core\            shared engine, compiled into BOTH projects (C# 7-compatible, no records)
    Vec3.cs            3-vector maths (metres)
    ProfileCatalog.cs  open profile shapes (flat, angle, channel, zee, hat, jay, polyline) + the JSON catalogue
    GaugeTable.cs      reads SOLIDWORKS .xlsx gauge tables without Excel (zip + XML)
    MemberSpec.cs      the spec: profile, thickness/gauge, groups (side, corner, gap, rotation, pierce, mirror)
    PathChain.cs       selected sketch lines / edges → segments in model space → ordered chains
    MemberBuilder.cs   the SOLIDWORKS work (see "How it builds")
  AddIn\           net48 COM add-in: SwAddin.cs (ISwAddin, commands, registration), MemberPage.cs (PropertyManager page), Icons\
  Cli\             net10 console: ssm.exe (attach | catalog | gauges | build | lab | load-addin)
  Catalog\profiles.json   the Standard → Type → Size catalogue (copy to %APPDATA%\StructuralSheetMetal\profiles.json to own it)
  install\register.ps1    build + deploy to %LOCALAPPDATA%\StructuralSheetMetal\AddIn + register (per-user by default, -Machine for HKLM)
  install\unregister.ps1
```

## Build, register, test

```powershell
# engine from the terminal, against the SOLIDWORKS already running as you
dotnet build StructuralSheetMetal\Cli\ssm.csproj -c Release
$ssm = "StructuralSheetMetal\Cli\bin\Release\net10.0-windows\ssm.exe"
& $ssm lab rect --std="Generic (inch)" --type=Angle --size="1.5 x 1.5" --table="sample table - steel - english units" --gauge="16 Gauge" --png=rect.png
& $ssm lab hug                       # inner channel INSIDE + outer angle OUTSIDE on one rectangle: they touch along the path
& $ssm lab l --corner=butt1 --side=outside --shape=channel --W=50 --F=25 --pierce="Web Center"
& $ssm build --type=Channel --size="2 x 1" --gauge="14 Gauge"   # on YOUR selection in the active part

# the add-in
.\StructuralSheetMetal\install\register.ps1      # builds, deploys, registers per-user (no admin)
# restart SOLIDWORKS → Tools > Add-Ins > Structural Sheet Metal → the "Structural Sheet Metal" tab in parts
```

`lab` makes a new part with a "Path" sketch on the Top plane (`rect`, `l`, `u`, `z`, `two`, `hug`;
`--pw/--ph` path size) and builds on it; `--png=` saves an isometric picture; earlier unsaved lab parts
are closed first. Options shared by `lab` and `build`: `--std/--type/--size` or `--shape=angle --A=38.1 --B=38.1`,
`--t --r --k`, `--table --gauge`, `--side=inside|outside`, `--corner=miter|butt1|butt2|none`, `--gap`,
`--rot`, `--pierce="Corner"`, `--mirror=h|v|hv`, `--name`.

## The page (mirrors Structural Member)

* **Profile** — Standard, Type, Size (from `profiles.json`); gauge table (from Tools > Options > File
  Locations > Sheet Metal Gauge Table) and gauge, or a typed thickness; inside bend radius (defaults to
  the table's smallest usable radius, else the thickness).
* **Groups** — several groups, each with its own path and settings; *New Group* copies the current
  settings so "same path, Outside" is two clicks.
* **Path segments** — sketch lines (2D or 3D) or straight edges. Connected segments get a corner
  treatment: *End miter*, *End butt 1* (first member through), *End butt 2* (second through), *None*;
  optional *Gap*.
* **Locate profile** — **Material side: Inside / Outside** (the hug toggle), rotation about the path,
  pierce point (every vertex and every segment midpoint of the shape, so a channel offers 7: Web
  Center, Corner 1/2, Flange 1/2 Mid, Flange 1/2 End; a hat also has Base Center), mirror
  horizontal / vertical, optional alignment line for the profile's "up".

Inside = the sheet goes to the concave side of the profile (the L's inner corner). Two members on the
same path with the same profile, one Inside and one Outside, share the profile line: the outer one
hugs the inner one. For a flat, Inside = +Y of the profile.

* **Show preview** (on by default) — every member as temporary bodies, one thin box per profile leg
  (no bends, no corner cuts), redrawn on every change; the group being edited is blue, the others
  yellow. Nothing is created: the bodies vanish when the page closes. To rotate or flip only some
  segments, give them their own group. `MemberPreview` builds each leg as a planar sheet trimmed by
  four lines (`CreateTrimmedSheet4`, curves passed as `DispatchWrapper[]`, the same marshalling rule as
  the cut bodies), extruded along the path with `CreateExtrudedBody` and shown with `IBody2.Display3`.

## How it builds (`MemberBuilder`)

1. **Plan** — chain the segments; per member set D (along the path), Q ("up": alignment line, else the
   path sketch's normal, else model Z) and P = D × Q; rotate; thicken the profile into a 2D point cloud;
   work out each member's start/end shift and the cut planes for its corners. Pure maths.
2. **Planes** — ONE hidden 3D construction sketch with a 10 mm line per plane, then `InsertRefPlane`
   perpendicular to each line at its start (origin on the point). Member planes and cut planes alike.
3. **Members** — the open profile is sketched on its plane through `ISketch.ModelToSketchTransform`
   (so the plane's own axes never matter), then `InsertSheetMetalBaseFlange2` with the API-help argument
   pattern. Mass properties then say where the material went: extrude direction and thickness side are
   corrected through `IBaseFlangeFeatureData` (`ReverseDirection`, `ReverseThickness`) and the gauge
   table is attached (`UseGaugeTable`, `GaugeTablePath`, `ThicknessTableName`).
4. **Corners** — `InsertCutSurface` on each miter/butt plane, limited to the one body; the kept side
   is checked by centre of mass and flipped via `ISurfaceCutFeatureData.Flip` if needed.
5. **Tidy** — construction hidden; two feature-tree folders (members + corners, construction).

Corner maths: miter plane normal = `D1 + D2`, each member extended by `R / tan(φ/2)` (R = radial
extent of the thickened section); gap = half on each side along the normal. Butt: the through member
ends at the other section's extent along its own axis; the butting member is trimmed by the through
member's outer face plane (+ gap).

### SOLIDWORKS API findings (the expensive ones)

* `FeatureByPositionReverse(0)` is **not** the sketch you just exited in a sheet-metal part: the
  Flat-Pattern folder is always last. Cast `ISketchManager.ActiveSketch` to `IFeature` *while the sketch
  is active*. This is why "the second base flange always failed".
* `InsertCutSurface`'s `Bodies` must be a `DispatchWrapper[]`; a plain `object[]` is marshalled as
  VARIANT[] and SOLIDWORKS answers error 1 (`swSurfaceCutFeatureError_BodiesNotSpecified`).
* Base Flange from an open profile: only the API-help argument pattern is accepted (default relief,
  `DirToUse = 1`, Merge/FeatScope/AutoSelect false). Do not pass `FlipExtruDir = true`; fix the
  direction afterwards via the definition.
* The gauge row name must match the cell text exactly, trailing space included (`"16 Gauge "` in the
  shipped tables). `GaugeRow.RawName` keeps it.
* SOLIDWORKS inserts its own `Sheet-Metal<n>` features between your planes and your first flange;
  they cannot be foldered, hence two folders.
* SOLIDWORKS keeps a loaded add-in DLL locked and the CLR cannot reload a same-version assembly:
  after rebuilding the add-in, **restart SOLIDWORKS**. `register.ps1` deploys to a versioned folder
  when the current one is locked.
* The COM class can be registered per-user, but SOLIDWORKS 2026 lists add-ins from
  `HKLM\SOFTWARE\SolidWorks\AddIns` only, so `register.ps1` writes that one key through an elevated
  helper (`hklm-addin-key.ps1`, one UAC prompt).
* The add-in class needs a dispatch interface (`ClassInterfaceType.AutoDispatch`): SOLIDWORKS hands
  the callback object over through IDispatch and calls toolbar callbacks by name. With
  `ClassInterfaceType.None`, `SetAddinCallbackInfo2` throws InvalidCastException and "Active" never sticks.
* Running the command from inside the path sketch is fine: the builder leaves the active sketch first,
  otherwise `Insert3DSketch` would exit that sketch instead of opening the construction sketch.

## Test plan for the add-in UI (not yet run)

1. `install\register.ps1`, restart SOLIDWORKS, Tools > Add-Ins: *Structural Sheet Metal* listed and
   ticked; a *Structural Sheet Metal* tab with one button in part documents (menu: Tools).
2. New part, sketch a 300 × 200 rectangle on Top, exit, select the four lines, press the button: the
   page opens with the four segments listed; OK → four mitered members, folders `SSM1 …`.
3. Rotation box: type 90 and check the result is a quarter turn. (Confirmed: the angle numberbox
   works in radians through the API while displaying degrees; the page converts both ways.)
4. New Group → Outside → same four lines → OK: the hug.
5. Gauge table dropdown lists the files from File Locations; picking a gauge greys the thickness box
   and sets the radius; the resulting flange shows the gauge in its PropertyManager.
6. Butt 1 / Butt 2 / Gap on an L path; mirror H/V; pierce point "Leg A End".

## Next

* **Profiles as SOLIDWORKS files.** The JSON catalogue is one parametric shape per type with sizes as
  rows, which replaces a `.sldlfp` per size. If you would rather author profiles as sketch files with
  configurations (L, C, Z, OMEGA…), a loader that reads an open sketch + its configuration list from a
  part file drops into `ProfileCatalog` without touching the builder.
* Arcs in the path (swept flange per tangent run), a preview before OK, a macro feature so a path edit
  rebuilds the members, cut-list properties (profile name, gauge, length) on each body.
