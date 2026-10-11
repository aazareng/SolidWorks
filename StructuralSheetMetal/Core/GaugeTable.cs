using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace StructuralSheetMetal
{
    public sealed class GaugeRow
    {
        public string Name = "";
        /// <summary>The cell text exactly as written (SolidWorks ships tables with trailing spaces, and matches the exact text).</summary>
        public string RawName = "";
        public double ThicknessMm;
        public List<double> RadiiMm = new List<double>();
        public override string ToString()
        {
            return string.Format("{0}: t={1:0.###} mm, radii [{2}]", Name, ThicknessMm, string.Join(", ", RadiiMm.Select(r => r.ToString("0.###"))));
        }
    }

    /// <summary>
    /// A SolidWorks sheet-metal gauge table (.xlsx), read without Excel: the workbook is a zip of XML.
    /// Layout (as SolidWorks ships them): a few "Type:/Process/K-Factor/Unit:" rows, then a header row
    /// whose first cell is "Gauge No.", then one row per gauge: name | thickness | "r1; r2; r3".
    /// </summary>
    public sealed class GaugeTable
    {
        public string Path = "";
        public string Title = "";
        public string Unit = "mm";
        public double UnitToMm = 1;
        public double? KFactor;
        public List<GaugeRow> Rows = new List<GaugeRow>();

        public string Name { get { return System.IO.Path.GetFileNameWithoutExtension(Path); } }

        public GaugeRow? Find(string gaugeName)
        {
            if (string.IsNullOrEmpty(gaugeName)) return null;
            var want = Norm(gaugeName);
            return Rows.FirstOrDefault(r => Norm(r.Name) == want)
                ?? Rows.FirstOrDefault(r => Norm(r.Name).StartsWith(want))
                ?? Rows.FirstOrDefault(r => want.StartsWith(Norm(r.Name)));
        }
        static string Norm(string s) { return s.Trim().ToLowerInvariant().Replace("  ", " "); }

        /// <summary>Default inside bend radius for a row: the smallest listed radius not below the thickness, else the thickness.</summary>
        public static double DefaultRadiusMm(GaugeRow row)
        {
            var ok = row.RadiiMm.Where(r => r >= row.ThicknessMm - 1e-9).OrderBy(r => r).ToList();
            if (ok.Count > 0) return ok[0];
            if (row.RadiiMm.Count > 0) return row.RadiiMm.Max();
            return row.ThicknessMm;
        }

        /// <summary>The folders SolidWorks searches for gauge tables (Tools > Options > File Locations), ';' separated.</summary>
        public static List<string> SplitFolders(string? pref)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(pref)) return list;
            foreach (var p in pref!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var d = p.Trim();
                if (d.Length > 0 && Directory.Exists(d)) list.Add(d);
            }
            return list;
        }

        public static List<string> ListFiles(IEnumerable<string> folders)
        {
            var files = new List<string>();
            foreach (var f in folders)
            {
                try { files.AddRange(Directory.GetFiles(f, "*.xlsx").Where(x => !System.IO.Path.GetFileName(x).StartsWith("~$"))); }
                catch { /* unreadable folder */ }
            }
            return files.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
        }

        public static GaugeTable Load(string path)
        {
            var t = new GaugeTable { Path = path };
            using (var zip = ZipFile.OpenRead(path))
            {
                var shared = ReadSharedStrings(zip);
                var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml")
                    ?? zip.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));
                if (sheetEntry == null) throw new InvalidOperationException("no worksheet in " + path);
                var rows = ReadRows(sheetEntry, shared);
                Parse(t, rows);
            }
            if (t.Rows.Count == 0) throw new InvalidOperationException("no gauge rows recognised in " + path + " (expected a 'Gauge No.' header row)");
            return t;
        }

        static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var e = zip.GetEntry("xl/sharedStrings.xml");
            if (e == null) return list;
            using (var s = e.Open())
            {
                var doc = XDocument.Load(s);
                XNamespace ns = doc.Root!.Name.Namespace;
                foreach (var si in doc.Root.Elements(ns + "si"))
                    list.Add(string.Concat(si.Descendants(ns + "t").Select(x => x.Value)));
            }
            return list;
        }

        /// <summary>Row → (column letter → text).</summary>
        static List<Dictionary<string, string>> ReadRows(ZipArchiveEntry entry, List<string> shared)
        {
            var rows = new List<Dictionary<string, string>>();
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                XNamespace ns = doc.Root!.Name.Namespace;
                var data = doc.Root.Element(ns + "sheetData");
                if (data == null) return rows;
                foreach (var row in data.Elements(ns + "row"))
                {
                    var cells = new Dictionary<string, string>();
                    foreach (var c in row.Elements(ns + "c"))
                    {
                        var r = (string?)c.Attribute("r") ?? "";
                        var col = new string(r.TakeWhile(char.IsLetter).ToArray());
                        var type = (string?)c.Attribute("t") ?? "";
                        string text;
                        if (type == "s")
                        {
                            int idx;
                            var v = c.Element(ns + "v")?.Value ?? "";
                            text = int.TryParse(v, out idx) && idx >= 0 && idx < shared.Count ? shared[idx] : "";
                        }
                        else if (type == "inlineStr")
                            text = string.Concat(c.Descendants(ns + "t").Select(x => x.Value));
                        else
                            text = c.Element(ns + "v")?.Value ?? "";
                        if (text.Length > 0) cells[col] = text;
                    }
                    if (cells.Count > 0) rows.Add(cells);
                }
            }
            return rows;
        }

        static string Cell(Dictionary<string, string> row, string col)
        {
            string? v; return row.TryGetValue(col, out v) && v != null ? v.Trim() : "";
        }

        static bool Num(string s, out double v)
        {
            return double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }

        static void Parse(GaugeTable t, List<Dictionary<string, string>> rows)
        {
            bool inData = false;
            foreach (var row in rows)
            {
                var a = Cell(row, "A"); var b = Cell(row, "B"); var c = Cell(row, "C");
                var al = a.ToLowerInvariant();
                if (!inData)
                {
                    if (al.StartsWith("type")) { t.Title = b; continue; }
                    if (al.StartsWith("unit")) { SetUnit(t, b); continue; }
                    if (al.StartsWith("k-factor") || al.StartsWith("k factor") || al.StartsWith("kfactor"))
                    {
                        double k; if (Num(b, out k)) t.KFactor = k; continue;
                    }
                    if (al.StartsWith("gauge no") || al == "gauge" || al.StartsWith("gauge number")) { inData = true; continue; }
                    continue;
                }
                double th;
                if (a.Length == 0 || !Num(b, out th))
                {
                    if (t.Rows.Count > 0 && a.Length == 0) break;   // blank after the data = end of table
                    continue;
                }
                string rawA; if (!row.TryGetValue("A", out rawA!) || rawA == null) rawA = a;
                var g = new GaugeRow { Name = a, RawName = rawA, ThicknessMm = th * t.UnitToMm };
                foreach (var part in c.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    double r; if (Num(part, out r)) g.RadiiMm.Add(r * t.UnitToMm);
                }
                t.Rows.Add(g);
            }
        }

        static void SetUnit(GaugeTable t, string unit)
        {
            var u = unit.Trim().ToLowerInvariant();
            if (u.StartsWith("inch") || u == "in" || u == "\"") { t.Unit = "in"; t.UnitToMm = 25.4; }
            else if (u.StartsWith("millimet") || u == "mm") { t.Unit = "mm"; t.UnitToMm = 1; }
            else if (u.StartsWith("centimet") || u == "cm") { t.Unit = "cm"; t.UnitToMm = 10; }
            else if (u.StartsWith("meter") || u.StartsWith("metre") || u == "m") { t.Unit = "m"; t.UnitToMm = 1000; }
            else if (u.StartsWith("feet") || u.StartsWith("foot") || u == "ft") { t.Unit = "ft"; t.UnitToMm = 304.8; }
            else { t.Unit = unit; t.UnitToMm = 1; }
        }
    }
}
