using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace StructuralSheetMetal
{
    /// <summary>A point of a profile, in profile-local millimetres (X right, Y up).</summary>
    public struct P2
    {
        public double X, Y;
        public P2(double x, double y) { X = x; Y = y; }
        public static P2 operator -(P2 a, P2 b) { return new P2(a.X - b.X, a.Y - b.Y); }
        public static P2 operator +(P2 a, P2 b) { return new P2(a.X + b.X, a.Y + b.Y); }
        public static P2 operator *(P2 a, double s) { return new P2(a.X * s, a.Y * s); }
        public double Length { get { return Math.Sqrt(X * X + Y * Y); } }
        public override string ToString() { return string.Format("({0:0.###}, {1:0.###})", X, Y); }
    }

    /// <summary>
    /// An OPEN sheet-metal profile: a polyline of straight segments (the sheet's mid-surface,
    /// before thickness), plus named points the user can choose as the pierce point.
    /// An L is two lines; a C is three. This is the whole point of the tool.
    /// </summary>
    public sealed class Profile
    {
        public string Name = "";
        public List<P2> Points = new List<P2>();
        /// <summary>Named points in display order; the first one is the default pierce point.</summary>
        public List<KeyValuePair<string, P2>> NamedPoints = new List<KeyValuePair<string, P2>>();

        public string DefaultPierce { get { return NamedPoints.Count > 0 ? NamedPoints[0].Key : ""; } }

        public P2 Pierce(string name)
        {
            foreach (var kv in NamedPoints)
                if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase)) return kv.Value;
            return NamedPoints.Count > 0 ? NamedPoints[0].Value : new P2(0, 0);
        }

        /// <summary>Shoelace area of the polyline closed back to its start. Sign tells which side is "inside".</summary>
        public double SignedArea()
        {
            double a = 0;
            for (int i = 0; i < Points.Count; i++)
            {
                var p = Points[i]; var q = Points[(i + 1) % Points.Count];
                a += p.X * q.Y - q.X * p.Y;
            }
            return a / 2;
        }

        public Profile Mirrored(bool horizontal, bool vertical)
        {
            if (!horizontal && !vertical) return this;
            Func<P2, P2> f = p => new P2(horizontal ? -p.X : p.X, vertical ? -p.Y : p.Y);
            var m = new Profile { Name = Name };
            m.Points = Points.Select(f).ToList();
            m.NamedPoints = NamedPoints.Select(kv => new KeyValuePair<string, P2>(kv.Key, f(kv.Value))).ToList();
            return m;
        }

        public double TotalLength()
        {
            double l = 0;
            for (int i = 0; i + 1 < Points.Count; i++) l += (Points[i + 1] - Points[i]).Length;
            return l;
        }
    }

    /// <summary>One catalogue row: Standard / Type / Size → a parametric shape.</summary>
    public sealed class ProfileSize
    {
        [JsonPropertyName("standard")] public string Standard { get; set; } = "";
        [JsonPropertyName("type")] public string Type { get; set; } = "";
        [JsonPropertyName("size")] public string Size { get; set; } = "";
        /// <summary>flat | angle | channel | zee | hat | jay | polyline</summary>
        [JsonPropertyName("shape")] public string Shape { get; set; } = "angle";
        [JsonPropertyName("params")] public Dictionary<string, double> Params { get; set; } = new Dictionary<string, double>();
        /// <summary>For shape "polyline": [[x,y],...] in mm.</summary>
        [JsonPropertyName("points")] public List<double[]>? Points { get; set; }

        public Profile Build()
        {
            var p = ProfileShapes.Build(Shape, Params, Points);
            p.Name = string.Format("{0} {1} {2}", Standard, Type, Size).Trim();
            return p;
        }
    }

    /// <summary>The parametric shapes. All in mm, X right, Y up. Points run in one direction so the
    /// shoelace sign is meaningful; "inside" is the concave side of the open shape.</summary>
    public static class ProfileShapes
    {
        public static readonly string[] Names = { "flat", "angle", "channel", "zee", "hat", "jay", "polyline" };

        static double Get(Dictionary<string, double> d, string key, double dflt)
        {
            double v;
            foreach (var kv in d) if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase)) { v = kv.Value; return v; }
            return dflt;
        }

        public static Profile Build(string shape, Dictionary<string, double> prm, List<double[]>? pts)
        {
            prm = prm ?? new Dictionary<string, double>();
            var p = new Profile();
            switch ((shape ?? "").ToLowerInvariant())
            {
                case "flat":
                {
                    double w = Get(prm, "W", 50);
                    p.Points.Add(new P2(-w / 2, 0)); p.Points.Add(new P2(w / 2, 0));
                    Named(p, "Center", 0, 0); Named(p, "Left End", -w / 2, 0); Named(p, "Right End", w / 2, 0);
                    break;
                }
                case "angle":
                {
                    // Leg A goes UP from the corner, leg B goes RIGHT. Inside = the quadrant between them.
                    double a = Get(prm, "A", 38.1), b = Get(prm, "B", 38.1);
                    p.Points.Add(new P2(0, a)); p.Points.Add(new P2(0, 0)); p.Points.Add(new P2(b, 0));
                    Named(p, "Corner", 0, 0); Named(p, "Leg A End", 0, a); Named(p, "Leg B End", b, 0);
                    Named(p, "Leg A Mid", 0, a / 2); Named(p, "Leg B Mid", b / 2, 0);
                    break;
                }
                case "channel":
                {
                    // Web W along X at the bottom, two flanges F going UP. Inside = above the web.
                    double w = Get(prm, "W", 50), f = Get(prm, "F", 25);
                    p.Points.Add(new P2(0, f)); p.Points.Add(new P2(0, 0)); p.Points.Add(new P2(w, 0)); p.Points.Add(new P2(w, f));
                    Named(p, "Web Center", w / 2, 0); Named(p, "Corner 1", 0, 0); Named(p, "Corner 2", w, 0);
                    Named(p, "Flange 1 End", 0, f); Named(p, "Flange 2 End", w, f);
                    break;
                }
                case "zee":
                {
                    // Bottom flange A to the LEFT, web W up, top flange B to the RIGHT.
                    double a = Get(prm, "A", 25), w = Get(prm, "W", 50), b = Get(prm, "B", 25);
                    p.Points.Add(new P2(-a, 0)); p.Points.Add(new P2(0, 0)); p.Points.Add(new P2(0, w)); p.Points.Add(new P2(b, w));
                    Named(p, "Bottom Corner", 0, 0); Named(p, "Top Corner", 0, w); Named(p, "Web Center", 0, w / 2);
                    Named(p, "Flange A End", -a, 0); Named(p, "Flange B End", b, w);
                    break;
                }
                case "hat":
                {
                    // Feet F on the base line, legs H up, crown C across the top. Inside = under the crown.
                    double c = Get(prm, "C", 50), h = Get(prm, "H", 25), f = Get(prm, "F", 20);
                    p.Points.Add(new P2(-(c / 2 + f), 0)); p.Points.Add(new P2(-c / 2, 0)); p.Points.Add(new P2(-c / 2, h));
                    p.Points.Add(new P2(c / 2, h)); p.Points.Add(new P2(c / 2, 0)); p.Points.Add(new P2(c / 2 + f, 0));
                    Named(p, "Base Center", 0, 0); Named(p, "Crown Center", 0, h);
                    Named(p, "Foot 1 End", -(c / 2 + f), 0); Named(p, "Foot 2 End", c / 2 + f, 0);
                    break;
                }
                case "jay":
                {
                    // A channel with one short flange: long flange A, web W, short lip B.
                    double w = Get(prm, "W", 50), a = Get(prm, "A", 25), b = Get(prm, "B", 10);
                    p.Points.Add(new P2(0, a)); p.Points.Add(new P2(0, 0)); p.Points.Add(new P2(w, 0)); p.Points.Add(new P2(w, b));
                    Named(p, "Corner 1", 0, 0); Named(p, "Corner 2", w, 0); Named(p, "Web Center", w / 2, 0);
                    Named(p, "Long End", 0, a); Named(p, "Short End", w, b);
                    break;
                }
                case "polyline":
                {
                    if (pts == null || pts.Count < 2) throw new ArgumentException("polyline profile needs at least two points");
                    int i = 0;
                    foreach (var xy in pts)
                    {
                        if (xy == null || xy.Length < 2) throw new ArgumentException("polyline point needs [x, y]");
                        p.Points.Add(new P2(xy[0], xy[1]));
                        Named(p, "Point " + (++i), xy[0], xy[1]);
                    }
                    break;
                }
                default:
                    throw new ArgumentException("unknown profile shape '" + shape + "' (flat, angle, channel, zee, hat, jay, polyline)");
            }
            return p;
        }

        static void Named(Profile p, string name, double x, double y)
        {
            p.NamedPoints.Add(new KeyValuePair<string, P2>(name, new P2(x, y)));
        }
    }

    /// <summary>The Standard → Type → Size catalogue, loaded from profiles.json.</summary>
    public sealed class ProfileCatalog
    {
        [JsonPropertyName("sizes")] public List<ProfileSize> Sizes { get; set; } = new List<ProfileSize>();

        static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true, WriteIndented = true, ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true
        };

        public static ProfileCatalog Load(string path)
        {
            var cat = JsonSerializer.Deserialize<ProfileCatalog>(File.ReadAllText(path), JsonOpts);
            if (cat == null || cat.Sizes.Count == 0) throw new InvalidOperationException("profile catalogue is empty: " + path);
            return cat;
        }

        public void Save(string path) { File.WriteAllText(path, JsonSerializer.Serialize(this, JsonOpts)); }

        /// <summary>
        /// Where the catalogue lives: %APPDATA%\StructuralSheetMetal\profiles.json if present (the user's own),
        /// else the Catalog\profiles.json shipped next to the assembly, else a built-in default.
        /// </summary>
        public static string ResolvePath(string? baseDir)
        {
            var user = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "StructuralSheetMetal", "profiles.json");
            if (File.Exists(user)) return user;
            foreach (var dir in CandidateDirs(baseDir))
            {
                var p = Path.Combine(dir, "profiles.json");
                if (File.Exists(p)) return p;
            }
            return user; // not there; caller falls back to Default()
        }

        static IEnumerable<string> CandidateDirs(string? baseDir)
        {
            if (string.IsNullOrEmpty(baseDir)) yield break;
            var d = baseDir!;
            yield return Path.Combine(d, "Catalog");
            yield return d;
            for (int up = 0; up < 4; up++)
            {
                var parent = Path.GetDirectoryName(d);
                if (string.IsNullOrEmpty(parent)) break;
                d = parent!;
                yield return Path.Combine(d, "Catalog");
            }
        }

        public static ProfileCatalog LoadOrDefault(string? baseDir, out string source)
        {
            var path = ResolvePath(baseDir);
            if (File.Exists(path)) { source = path; return Load(path); }
            source = "(built-in default)";
            return Default();
        }

        public IEnumerable<string> Standards() { return Sizes.Select(s => s.Standard).Distinct(); }
        public IEnumerable<string> Types(string standard) { return Sizes.Where(s => s.Standard == standard).Select(s => s.Type).Distinct(); }
        public IEnumerable<ProfileSize> SizesOf(string standard, string type) { return Sizes.Where(s => s.Standard == standard && s.Type == type); }
        public ProfileSize? Find(string standard, string type, string size)
        {
            return Sizes.FirstOrDefault(s => Eq(s.Standard, standard) && Eq(s.Type, type) && Eq(s.Size, size));
        }
        static bool Eq(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

        /// <summary>A sensible starter catalogue (inch sizes written in mm).</summary>
        public static ProfileCatalog Default()
        {
            var c = new ProfileCatalog();
            const double I = 25.4;
            foreach (var leg in new[] { 0.75, 1, 1.25, 1.5, 2, 2.5, 3 })
                c.Sizes.Add(Row("Generic (inch)", "Angle", string.Format("{0} x {0}", leg), "angle", "A", leg * I, "B", leg * I));
            c.Sizes.Add(Row("Generic (inch)", "Angle", "2 x 3", "angle", "A", 3 * I, "B", 2 * I));
            c.Sizes.Add(Row("Generic (inch)", "Angle", "3 x 2", "angle", "A", 2 * I, "B", 3 * I));
            foreach (var w in new[] { 1, 1.5, 2, 3, 4 })
                c.Sizes.Add(Row("Generic (inch)", "Channel", string.Format("{0} x 1", w), "channel", "W", w * I, "F", 1 * I));
            c.Sizes.Add(Row("Generic (inch)", "Channel", "2 x 1.5", "channel", "W", 2 * I, "F", 1.5 * I));
            c.Sizes.Add(Row("Generic (inch)", "Channel", "3 x 1.5", "channel", "W", 3 * I, "F", 1.5 * I));
            c.Sizes.Add(Row("Generic (inch)", "Zee", "1 x 2 x 1", "zee", "A", 1 * I, "W", 2 * I, "B", 1 * I));
            c.Sizes.Add(Row("Generic (inch)", "Zee", "1.5 x 3 x 1.5", "zee", "A", 1.5 * I, "W", 3 * I, "B", 1.5 * I));
            c.Sizes.Add(Row("Generic (inch)", "Hat", "2 x 1 x 1", "hat", "C", 2 * I, "H", 1 * I, "F", 1 * I));
            c.Sizes.Add(Row("Generic (inch)", "Hat", "3 x 1.5 x 1", "hat", "C", 3 * I, "H", 1.5 * I, "F", 1 * I));
            c.Sizes.Add(Row("Generic (inch)", "Jay", "2 x 1 x 0.5", "jay", "W", 2 * I, "A", 1 * I, "B", 0.5 * I));
            foreach (var w in new[] { 1, 2, 3, 4 })
                c.Sizes.Add(Row("Generic (inch)", "Flat", string.Format("{0}", w), "flat", "W", w * I));
            foreach (var leg in new[] { 20, 25, 30, 40, 50 })
                c.Sizes.Add(Row("Generic (mm)", "Angle", string.Format("{0} x {0}", leg), "angle", "A", leg, "B", leg));
            foreach (var w in new[] { 30, 40, 50, 80 })
                c.Sizes.Add(Row("Generic (mm)", "Channel", string.Format("{0} x 25", w), "channel", "W", w, "F", 25));
            c.Sizes.Add(Row("Generic (mm)", "Zee", "25 x 50 x 25", "zee", "A", 25, "W", 50, "B", 25));
            c.Sizes.Add(Row("Generic (mm)", "Hat", "50 x 25 x 20", "hat", "C", 50, "H", 25, "F", 20));
            return c;
        }

        static ProfileSize Row(string std, string type, string size, string shape, params object[] kv)
        {
            var r = new ProfileSize { Standard = std, Type = type, Size = size, Shape = shape };
            for (int i = 0; i + 1 < kv.Length; i += 2) r.Params[(string)kv[i]] = Convert.ToDouble(kv[i + 1]);
            return r;
        }
    }
}
