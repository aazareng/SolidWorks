// ============================================================================
//  ssm — Structural Sheet Metal from the terminal, against the running SolidWorks.
//
//    ssm attach                         SolidWorks version + active document
//    ssm catalog                        the Standard / Type / Size catalogue in use
//    ssm gauges [table.xlsx|name]       gauge rows read from a gauge table (default: first table SolidWorks knows)
//    ssm build [options]                build members along the CURRENT SELECTION (sketch lines / linear edges) of the active part
//    ssm lab <rect|l|u|hug|z> [options] new part, draw a test path on the Top plane, build, report
//    ssm load-addin <dll>               ask SolidWorks to load the registered add-in DLL now (no restart)
//
//  Options (all optional):
//    --std=.. --type=.. --size=..       catalogue row            --shape=angle --A=38.1 --B=38.1 (direct shape + params)
//    --t=1.5 --r=1.5 --k=0.5            thickness / inside radius / K-factor, mm
//    --table=<path|name> --gauge="16 Gauge"   gauge table + row (thickness and radius then come from the table)
//    --side=inside|outside --corner=miter|butt1|butt2|none --gap=0 --rot=0 --pierce="Corner" --mirror=h|v|hv
//    --name="..."                       feature folder name
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using StructuralSheetMetal;

static class Program
{
    const double Mm = 0.001;

    [STAThread]
    static int Main(string[] args)
    {
        var opts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var rest = new List<string>();
        foreach (var a in args)
        {
            if (a.StartsWith("--")) { var i = a.IndexOf('='); if (i > 0) opts[a.Substring(2, i - 2)] = a.Substring(i + 1); else opts[a.Substring(2)] = "true"; }
            else rest.Add(a);
        }
        var cmd = rest.Count > 0 ? rest[0].ToLowerInvariant() : "help";
        try
        {
            switch (cmd)
            {
                case "attach":
                    Console.WriteLine(SwSession.Describe(SwSession.Connect(opts.ContainsKey("start"))));
                    return 0;
                case "catalog":
                {
                    string src;
                    var cat = ProfileCatalog.LoadOrDefault(AppContext.BaseDirectory, out src);
                    Console.WriteLine("catalogue: " + src);
                    foreach (var s in cat.Sizes) Console.WriteLine($"  {s.Standard} | {s.Type} | {s.Size}  [{s.Shape}: {string.Join(", ", s.Params.Select(kv => kv.Key + "=" + kv.Value))}]");
                    return 0;
                }
                case "gauges":
                {
                    var sw = SwSession.Connect(false);
                    var path = ResolveTable(sw, rest.Count > 1 ? rest[1] : null);
                    var t = GaugeTable.Load(path);
                    Console.WriteLine($"{t.Path}\n  title '{t.Title}', unit {t.Unit}, K={(t.KFactor.HasValue ? t.KFactor.Value.ToString() : "-")}, {t.Rows.Count} rows");
                    foreach (var r in t.Rows) Console.WriteLine("  " + r);
                    return 0;
                }
                case "build":
                {
                    var sw = SwSession.Connect(false);
                    var model = sw.ActiveDoc as IModelDoc2 ?? throw new InvalidOperationException("no active document");
                    var rejected = new List<string>();
                    var segs = PathChain.FromSelection(sw, model, -1, rejected);
                    foreach (var r in rejected) Console.WriteLine("skipped " + r);
                    if (segs.Count == 0) throw new InvalidOperationException("select sketch lines or straight edges first");
                    var spec = SpecFrom(sw, opts, segs);
                    var rep = MemberBuilder.Build(sw, model, spec);
                    Console.WriteLine(rep);
                    return rep.Ok ? 0 : 1;
                }
                case "lab":
                    return Lab(rest.Count > 1 ? rest[1].ToLowerInvariant() : "rect", opts);
                case "load-addin":
                {
                    var sw = SwSession.Connect(false);
                    var dll = Path.GetFullPath(rest.Count > 1 ? rest[1] : throw new ArgumentException("load-addin <dll>"));
                    int r = sw.LoadAddIn(dll);
                    Console.WriteLine(r == 0 ? "LoadAddIn ok: " + dll : $"LoadAddIn returned {r} for {dll} (0 = ok, 1 = failed/not registered, 2 = already loaded)");
                    object? obj = null;
                    try { obj = sw.GetAddInObject("StructuralSheetMetal.AddIn"); } catch (Exception e) { Console.WriteLine("GetAddInObject threw " + e.Message); }
                    Console.WriteLine(obj != null ? "add-in object is live (ConnectToSW ran)" : "add-in object NOT found: SolidWorks did not instantiate it (registration or load failure; see %TEMP%\\StructuralSheetMetal.log)");
                    return r == 0 && obj != null ? 0 : 1;
                }
                case "unload-addin":
                {
                    var sw = SwSession.Connect(false);
                    var dll = Path.GetFullPath(rest.Count > 1 ? rest[1] : throw new ArgumentException("unload-addin <dll>"));
                    int r = sw.UnloadAddIn(dll);
                    Console.WriteLine($"UnloadAddIn returned {r}");
                    return r == 0 ? 0 : 1;
                }
                default:
                    Console.WriteLine("ssm attach | catalog | gauges [table] | build [options] | lab <rect|l|u|hug|z> [options] | load-addin <dll> | unload-addin <dll>");
                    Console.WriteLine("options: --std= --type= --size= | --shape=angle --A= --B= ; --t= --r= --k= ; --table= --gauge= ; --side=inside|outside --corner=miter|butt1|butt2|none --gap= --rot= --pierce= --mirror=h|v|hv ; --name= ; lab: --pw= --ph= --png= --save=");
                    return cmd == "help" ? 0 : 2;
            }
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("ERROR: " + e.Message);
            if (opts.ContainsKey("trace")) Console.Error.WriteLine(e);
            return 1;
        }
    }

    // ------------------------------------------------------------------ spec from options
    static string? Opt(Dictionary<string, string> o, string k) { string v; return o.TryGetValue(k, out v!) ? v : null; }
    static double Dbl(Dictionary<string, string> o, string k, double d)
    {
        var v = Opt(o, k); double x;
        return v != null && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out x) ? x : d;
    }

    static string ResolveTable(ISldWorks sw, string? nameOrPath)
    {
        var folders = GaugeTable.SplitFolders(sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetMetalGaugeTable));
        var files = GaugeTable.ListFiles(folders);
        if (string.IsNullOrEmpty(nameOrPath))
        {
            if (files.Count == 0) throw new InvalidOperationException("SolidWorks has no gauge table folder with .xlsx files (Tools > Options > File Locations > Sheet Metal Gauge Table)");
            return files[0];
        }
        if (File.Exists(nameOrPath)) return Path.GetFullPath(nameOrPath);
        var hit = files.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).Equals(nameOrPath, StringComparison.OrdinalIgnoreCase))
               ?? files.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).IndexOf(nameOrPath, StringComparison.OrdinalIgnoreCase) >= 0);
        if (hit == null) throw new FileNotFoundException("no gauge table matches '" + nameOrPath + "'; known: " + string.Join("; ", files.Select(Path.GetFileName)));
        return hit;
    }

    static MemberSpec SpecFrom(ISldWorks sw, Dictionary<string, string> o, List<PathSegment> segs)
    {
        var spec = new MemberSpec();
        string src;
        var cat = ProfileCatalog.LoadOrDefault(AppContext.BaseDirectory, out src);
        var shape = Opt(o, "shape");
        if (shape != null)
        {
            var prm = new Dictionary<string, double>();
            foreach (var kv in o) if (kv.Key.Length <= 2 && kv.Key.All(char.IsLetter) && kv.Key.ToUpperInvariant() == kv.Key) prm[kv.Key] = double.Parse(kv.Value, CultureInfo.InvariantCulture);
            spec.Profile = ProfileShapes.Build(shape, prm, null);
            spec.Profile.Name = shape + " " + string.Join(" x ", prm.Values.Select(v => v.ToString("0.##")));
        }
        else
        {
            var std = Opt(o, "std") ?? cat.Standards().First();
            var type = Opt(o, "type") ?? cat.Types(std).First();
            var size = Opt(o, "size") ?? cat.SizesOf(std, type).First().Size;
            var row = cat.Find(std, type, size) ?? throw new ArgumentException($"no catalogue row {std} | {type} | {size} (catalogue {src})");
            spec.Profile = row.Build();
        }
        spec.ThicknessMm = Dbl(o, "t", 1.5);
        spec.KFactor = Dbl(o, "k", 0.5);
        if (Opt(o, "r") != null) spec.BendRadiusMm = Dbl(o, "r", spec.ThicknessMm);

        var table = Opt(o, "table"); var gauge = Opt(o, "gauge");
        if (table != null || gauge != null)
        {
            var path = ResolveTable(sw, table);
            var gt = GaugeTable.Load(path);
            var row = gauge != null ? gt.Find(gauge) : gt.Rows.FirstOrDefault(r => Math.Abs(r.ThicknessMm - spec.ThicknessMm) < 0.02) ?? gt.Rows[0];
            if (row == null) throw new ArgumentException($"gauge '{gauge}' not in {Path.GetFileName(path)}; rows: {string.Join(", ", gt.Rows.Select(r => r.Name))}");
            spec.GaugeTablePath = path; spec.GaugeName = row.RawName;
            spec.ThicknessMm = row.ThicknessMm;
            if (Opt(o, "r") == null) spec.BendRadiusMm = GaugeTable.DefaultRadiusMm(row);
            if (gt.KFactor.HasValue && Opt(o, "k") == null) spec.KFactor = gt.KFactor.Value;
        }
        spec.Name = Opt(o, "name") ?? "Structural Sheet Metal";

        var g = new MemberGroup { Segments = segs, Name = "Group1" };
        ApplyGroupOpts(g, o);
        spec.Groups.Add(g);
        return spec;
    }

    static void ApplyGroupOpts(MemberGroup g, Dictionary<string, string> o)
    {
        var side = Opt(o, "side"); if (side != null) g.Side = side.StartsWith("out", StringComparison.OrdinalIgnoreCase) ? MaterialSide.Outside : MaterialSide.Inside;
        var corner = Opt(o, "corner");
        if (corner != null)
            g.Corner = corner.ToLowerInvariant() switch { "none" => CornerTreatment.None, "butt1" => CornerTreatment.Butt1, "butt2" => CornerTreatment.Butt2, _ => CornerTreatment.Miter };
        g.GapMm = Dbl(o, "gap", 0);
        g.RotationDeg = Dbl(o, "rot", 0);
        g.PiercePoint = Opt(o, "pierce");
        var mirror = Opt(o, "mirror") ?? "";
        g.MirrorHorizontal = mirror.IndexOf('h') >= 0; g.MirrorVertical = mirror.IndexOf('v') >= 0;
    }

    // ------------------------------------------------------------------ lab
    /// <summary>Close earlier lab parts: unsaved "PartN" documents that carry our "Path" sketch. Nothing else is touched.</summary>
    static void CloseOldLabParts(ISldWorks sw)
    {
        var titles = new List<string>();
        for (var d = sw.GetFirstDocument() as IModelDoc2; d != null; d = d.GetNext() as IModelDoc2)
        {
            if (d.GetType() != (int)swDocumentTypes_e.swDocPART) continue;
            var title = d.GetTitle();
            if (!System.Text.RegularExpressions.Regex.IsMatch(title, @"^Part\d+(\.SLDPRT)?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) continue;
            if (!string.IsNullOrEmpty(d.GetPathName())) continue;         // saved somewhere: not ours to close
            bool ours = false;
            for (var f = d.FirstFeature() as IFeature; f != null; f = f.GetNextFeature() as IFeature)
                if (f.Name == "Path" && f.GetTypeName2() == "ProfileFeature") { ours = true; break; }
            if (ours) titles.Add(title);
        }
        foreach (var t in titles) { sw.CloseDoc(t); Console.WriteLine("closed earlier lab part " + t); }
    }

    static int Lab(string kind, Dictionary<string, string> o)
    {
        var sw = SwSession.Connect(o.ContainsKey("start"));
        try { Console.OutputEncoding = System.Text.Encoding.UTF8; } catch { }
        Console.WriteLine(SwSession.Describe(sw));
        CloseOldLabParts(sw);
        var template = sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        if (string.IsNullOrWhiteSpace(template) || !File.Exists(template)) throw new InvalidOperationException("no default part template: '" + template + "'");
        var model = sw.NewDocument(template, 0, 0, 0) as IModelDoc2 ?? throw new InvalidOperationException("NewDocument(part) failed");

        // path on the Top plane (second reference plane of the template)
        var planes = new List<IFeature>();
        for (var f = model.FirstFeature() as IFeature; f != null; f = f.GetNextFeature() as IFeature)
            if (f.GetTypeName2() == "RefPlane") planes.Add(f);
        planes[1].Select2(false, 0);
        var skm = model.SketchManager;
        skm.InsertSketch(true);
        var segs = new List<ISketchSegment>();
        double w = Dbl(o, "pw", 300) * Mm, h = Dbl(o, "ph", 200) * Mm;   // path size (not the profile params A, B, W...)
        ISketchSegment L(double x1, double y1, double x2, double y2) { var s = skm.CreateLine(x1, y1, 0, x2, y2, 0); segs.Add(s); return s; }
        switch (kind)
        {
            case "l": L(0, 0, w, 0); L(w, 0, w, h); break;
            case "u": L(0, h, 0, 0); L(0, 0, w, 0); L(w, 0, w, h); break;
            case "z": L(-w / 2, h, 0, h); L(0, h, 0, 0); L(0, 0, w / 2, 0); break;
            case "two": L(0, 0, w, 0); L(0, 2 * h, w, 2 * h); break;   // two disjoint members, no corners
            default:  L(0, 0, w, 0); L(w, 0, w, h); L(w, h, 0, h); L(0, h, 0, 0); break;   // rect, hug: counter-clockwise loop
        }
        var pathSketch = (IFeature)skm.ActiveSketch;
        model.ClearSelection2(true);
        skm.InsertSketch(true);
        pathSketch.Name = "Path";
        Console.WriteLine($"path '{kind}': {segs.Count} line(s), {w * 1000:0} x {h * 1000:0} mm on {planes[1].Name}");

        var rejected = new List<string>();
        var path = PathChain.FromEntities(sw, segs.Cast<object>(), rejected);
        foreach (var r in rejected) Console.WriteLine("rejected " + r);

        int rc = 0;
        if (o.ContainsKey("preview"))
        {
            // temporary-body preview, as the page shows it, held for a few seconds (and pictured if --png)
            var spec = SpecFrom(sw, o, path);
            using (var pv = new MemberPreview(sw, model))
            {
                pv.Show(spec, 0);
                Console.WriteLine($"preview: {pv.Members} member(s) as temp bodies" + (pv.Problems.Count > 0 ? "; problems: " + string.Join("; ", pv.Problems) : ""));
                var ppng = Opt(o, "png");
                if (ppng != null)
                {
                    model.ShowNamedView2("*Isometric", 7); model.ViewZoomtofit2();
                    int e1 = 0, w1 = 0;
                    var p = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(ppng))!, Path.GetFileNameWithoutExtension(ppng) + "-preview.png");
                    Console.WriteLine(model.Extension.SaveAs3(p, 0, 1, null, null, ref e1, ref w1) ? "preview picture " + p : "preview picture failed");
                }
                System.Threading.Thread.Sleep((int)(Dbl(o, "seconds", 4) * 1000));
            }
            Console.WriteLine("preview cleared");
        }
        if (kind == "hug")
        {
            // 1: a channel hugging the INSIDE of the loop; 2: an angle on the OUTSIDE, touching it
            var o1 = new Dictionary<string, string>(o, StringComparer.OrdinalIgnoreCase);
            if (!o1.ContainsKey("shape") && !o1.ContainsKey("type")) { o1["shape"] = "channel"; o1["W"] = "50"; o1["F"] = "25"; o1["pierce"] = "Web Center"; }
            o1["side"] = "inside"; o1["name"] = "Inner channel";
            var rep1 = MemberBuilder.Build(sw, model, SpecFrom(sw, o1, path));
            Console.WriteLine(rep1); rc |= rep1.Ok ? 0 : 1;

            var o2 = new Dictionary<string, string>(o, StringComparer.OrdinalIgnoreCase);
            o2["shape"] = "angle"; o2["A"] = Opt(o, "A") ?? "25"; o2["B"] = Opt(o, "B") ?? "25"; o2["pierce"] = "Corner";
            o2["side"] = "outside"; o2["name"] = "Outer angle"; o2["mirror"] = "h";
            var rep2 = MemberBuilder.Build(sw, model, SpecFrom(sw, o2, path));
            Console.WriteLine(rep2); rc |= rep2.Ok ? 0 : 1;
        }
        else
        {
            var rep = MemberBuilder.Build(sw, model, SpecFrom(sw, o, path));
            Console.WriteLine(rep);
            rc = rep.Ok ? 0 : 1;
        }

        var png = Opt(o, "png");
        if (png != null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(png))!);
            model.ShowNamedView2("*Isometric", 7);
            model.ViewZoomtofit2();
            int err = 0, warn = 0;
            bool ok = model.Extension.SaveAs3(Path.GetFullPath(png), 0, 1, null, null, ref err, ref warn);
            Console.WriteLine(ok ? "picture " + png : $"picture failed (err {err})");
        }
        var save = Opt(o, "save");
        if (save != null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(save))!);
            int err = 0, warn = 0;
            bool ok = model.Extension.SaveAs3(Path.GetFullPath(save), 0, 1, null, null, ref err, ref warn);
            Console.WriteLine(ok ? "saved " + save : $"save failed (err {err})");
        }
        return rc;
    }
}
