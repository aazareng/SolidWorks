using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace StructuralSheetMetal
{
    /// <summary>
    /// Builds "structural sheet metal" members in the active PART:
    ///
    ///   A. plan   — chain the path segments, set each member's frame (D along the path, Q "up", P = D × Q),
    ///               the thickened profile cloud, start/end shifts and the corner cut planes; all pure maths
    ///   B. planes — ONE hidden 3D construction sketch with a short line per plane, then a reference plane
    ///               perpendicular to each line at its start (origin on the point)
    ///   C. members— per member: the open profile sketched on its plane (via ModelToSketchTransform, so the
    ///               plane's own axis orientation does not matter), a Base Flange, then a mass-property check
    ///               that fixes the extrude direction and puts the thickness on the asked side (Inside /
    ///               Outside) through the feature definition, and attaches the gauge table
    ///   D. corners— Cut With Surface on each miter/butt plane, kept side verified by centre of mass
    ///   E. tidy   — hide construction, fold everything into one feature-tree folder
    ///
    /// Everything is a plain SolidWorks feature afterwards (planes, sketches, Base-Flange, Surface-Cut);
    /// nothing proprietary is left in the file.
    /// </summary>
    public sealed class MemberBuilder
    {
        const double Mm = 0.001;
        const double Margin = 0.0002;         // 0.2 mm extra where a member must overrun a cut
        const double HelperLineLen = 0.01;    // 10 mm construction lines

        readonly ISldWorks sw;
        readonly IModelDoc2 model;
        readonly IPartDoc part;
        readonly IFeatureManager fm;
        readonly ISketchManager skm;
        readonly ISelectionMgr selMgr;
        readonly IMathUtility mu;
        readonly MemberSpec spec;
        readonly BuildReport rep = new BuildReport();
        readonly List<IFeature> created = new List<IFeature>();

        sealed class Plan
        {
            public int Index;
            public MemberGroup Group = null!;
            public PathSegment Seg = null!;
            public Vec3 D, P, Q;
            public Profile Prof = null!;
            public P2 Pierce;
            public int Side;                     // +1 inside, -1 outside
            public double StartShift, EndShift;  // metres; positive grows the member beyond the path point
            public List<P2> Cloud = new List<P2>();   // thickened profile, mm, relative to the pierce point
            public List<P2> Mid = new List<P2>();     // mid-surface points relative to pierce, mm
            public List<P2> InNormal = new List<P2>(); // per profile segment, unit normal toward the inside
            public double RadialM;                     // max cloud radius, metres
            public IFeature? Plane, Sketch, Flange;
            public string Name = "";
            public Vec3 Start { get { return Seg.A - D * StartShift; } }
            public Vec3 End { get { return Seg.B + D * EndShift; } }
            public double Length { get { return Vec3.Dot(End - Start, D); } }
        }

        sealed class PlaneCut
        {
            public Plan Member = null!;
            public Vec3 Point, Normal;
            public int KeepSign;
            public IFeature? Plane;
            public string Label = "";
        }

        sealed class PlaneReq
        {
            public Vec3 Point, Normal;
            public string Name = "";
            public ISketchSegment? Line;
            public IFeature? Feature;
            public Vec3 ActualNormal;
        }

        MemberBuilder(ISldWorks sw, IModelDoc2 model, MemberSpec spec)
        {
            this.sw = sw; this.model = model; this.spec = spec;
            part = (IPartDoc)model;
            fm = model.FeatureManager;
            skm = model.SketchManager;
            selMgr = (ISelectionMgr)model.SelectionManager;
            mu = (IMathUtility)sw.GetMathUtility();
        }

        public static BuildReport Build(ISldWorks sw, IModelDoc2 model, MemberSpec spec)
        {
            if (model == null) { var r = new BuildReport { Ok = false, Error = new InvalidOperationException("no active document") }; r.Log("ERROR: no active document"); return r; }
            var b = new MemberBuilder(sw, model, spec);
            try { b.Run(); b.rep.Ok = true; }
            catch (Exception e) { b.rep.Ok = false; b.rep.Error = e; b.rep.Log("ERROR: " + e.Message); }
            return b.rep;
        }

        // ------------------------------------------------------------------ run
        void Run()
        {
            spec.Validate();
            if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                throw new InvalidOperationException("the active document must be a part");
            rep.Prefix = NextPrefix();
            rep.Log(string.Format("{0}: profile '{1}' ({2} points), t={3:0.###} mm, r={4:0.###} mm, K={5}{6}",
                rep.Prefix, spec.Profile.Name, spec.Profile.Points.Count, spec.ThicknessMm, spec.RadiusMm, spec.KFactor,
                spec.GaugeTablePath != null ? ", gauge '" + spec.GaugeName + "' from " + System.IO.Path.GetFileName(spec.GaugeTablePath) : ""));

            // A. plan
            var plans = new List<Plan>();
            var cuts = new List<PlaneCut>();
            int gi = 0;
            foreach (var group in spec.Groups)
            {
                gi++;
                var chains = PathChain.Chain(group.Segments);
                rep.Log(string.Format("group {0} '{1}': {2} segment(s) in {3} chain(s); side={4}, corner={5}, rot={6}°, pierce={7}",
                    gi, group.Name, group.Segments.Count, chains.Count, group.Side, group.Corner, group.RotationDeg, group.PiercePoint ?? "(default)"));
                foreach (var chain in chains)
                {
                    var members = chain.Segments.Select(s => MakePlan(group, s, plans.Count + chain.Segments.IndexOf(s) + 1)).ToList();
                    for (int i = 0; i + 1 < members.Count; i++) Corner(members[i], members[i + 1], cuts);
                    if (chain.Closed && members.Count >= 3) Corner(members[members.Count - 1], members[0], cuts);
                    plans.AddRange(members);
                }
            }
            foreach (var p in plans)
            {
                if (p.Length <= Tol())
                    throw new InvalidOperationException(string.Format("{0} has no length left after corner treatment (segment {1:0.##} mm, start shift {2:0.##}, end shift {3:0.##})",
                        p.Name, p.Seg.Length * 1000, p.StartShift * 1000, p.EndShift * 1000));
            }

            // B. planes
            var reqs = new List<PlaneReq>();
            foreach (var p in plans)
                reqs.Add(new PlaneReq { Point = p.Start, Normal = p.D, Name = rep.Prefix + " Plane" + p.Index });
            int ci = 0;
            foreach (var c in cuts)
            {
                ci++;
                c.Label = rep.Prefix + " Corner" + ci;
                reqs.Add(new PlaneReq { Point = c.Point, Normal = c.Normal, Name = rep.Prefix + " CutPlane" + ci });
            }
            CreatePlanes(reqs);
            for (int i = 0; i < plans.Count; i++) plans[i].Plane = reqs[i].Feature;
            for (int i = 0; i < cuts.Count; i++) cuts[i].Plane = reqs[plans.Count + i].Feature;

            // C. members
            for (int i = 0; i < plans.Count; i++) BuildMember(plans[i], reqs[i].ActualNormal);
            rep.Members = plans.Count;

            // D. corners
            foreach (var c in cuts) ApplyCut(c);

            // E. tidy
            Tidy(reqs);
            model.EditRebuild3();
            model.ViewZoomtofit2();
            var bodies = Bodies();
            rep.Log(string.Format("done: {0} solid bodies in the part, {1} feature(s) created", bodies.Count, created.Count));
        }

        static double Tol() { return 1e-6; }
        static string Dir(Vec3 v) { return string.Format("[{0:0.###},{1:0.###},{2:0.###}]", v.X, v.Y, v.Z); }

        // ------------------------------------------------------------------ A. plan
        Plan MakePlan(MemberGroup group, PathSegment seg, int index)
        {
            var plan = new Plan { Index = index, Group = group, Seg = seg, Name = rep.Prefix + " Member" + index };
            plan.Prof = spec.Profile.Mirrored(group.MirrorHorizontal, group.MirrorVertical);
            plan.Pierce = plan.Prof.Pierce(group.PiercePoint ?? plan.Prof.DefaultPierce);
            plan.Side = group.Side == MaterialSide.Inside ? 1 : -1;

            var d = seg.Dir;
            var up = group.AlignUp ?? seg.SketchNormal ?? Vec3.UnitZ;
            if (Vec3.Cross(up, d).Length < 1e-6) up = Vec3.Cross(Vec3.UnitY, d).Length > 1e-6 ? Vec3.UnitY : Vec3.UnitX;
            var q = (up - d * Vec3.Dot(up, d)).Normalized();
            var p = Vec3.Cross(d, q).Normalized();
            if (Math.Abs(group.RotationDeg) > 1e-9)
            {
                double a = group.RotationDeg * Math.PI / 180, ca = Math.Cos(a), sa = Math.Sin(a);
                var p2 = p * ca + q * sa;
                var q2 = q * ca - p * sa;
                p = p2; q = q2;
            }
            plan.D = d; plan.P = p; plan.Q = q;

            // thickened cloud and inside normals, mm, relative to the pierce point
            double t = spec.ThicknessMm;
            int sgn = plan.Prof.SignedArea() >= 0 ? 1 : -1;
            var pts = plan.Prof.Points.Select(x => x - plan.Pierce).ToList();
            plan.Mid = pts;
            for (int k = 0; k + 1 < pts.Count; k++)
            {
                var a = pts[k]; var b = pts[k + 1];
                var dir = b - a; var len = dir.Length;
                if (len < 1e-9) continue;
                var n = new P2(-dir.Y / len, dir.X / len) * sgn;   // toward the inside
                plan.InNormal.Add(n);
                var off = n * (t * plan.Side);
                plan.Cloud.Add(a); plan.Cloud.Add(b); plan.Cloud.Add(a + off); plan.Cloud.Add(b + off);
            }
            plan.RadialM = plan.Cloud.Count > 0 ? plan.Cloud.Max(c => c.Length) * Mm : 0;
            return plan;
        }

        /// <summary>Largest reach of the member's thickened section from its path point, along a model direction (metres).</summary>
        static double Extent(Plan m, Vec3 dir)
        {
            double best = double.NegativeInfinity;
            foreach (var c in m.Cloud)
            {
                var v = (m.P * c.X + m.Q * c.Y) * Mm;
                best = Math.Max(best, Vec3.Dot(v, dir));
            }
            return best == double.NegativeInfinity ? 0 : best;
        }

        void Corner(Plan i, Plan j, List<PlaneCut> cuts)
        {
            var c = i.Seg.B;
            if (!Vec3.Near(c, j.Seg.A, PathChain.Tol)) { rep.Warn(i.Name + " and " + j.Name + " do not share an endpoint; no corner treatment"); return; }
            var di = i.D; var dj = j.D;
            if (Vec3.Cross(di, dj).Length < 1e-6) { rep.Log(i.Name + "/" + j.Name + ": collinear, no corner treatment"); return; }
            var treat = i.Group.Corner;
            if (treat == CornerTreatment.None) return;
            double cosPhi = Math.Max(-1, Math.Min(1, Vec3.Dot(-di, dj)));
            double phi = Math.Acos(cosPhi);                       // angle between the two member axes at the corner
            double g = i.Group.GapMm * Mm;

            switch (treat)
            {
                case CornerTreatment.Miter:
                {
                    var sum = di + dj;
                    if (sum.Length < 1e-6) { rep.Warn(i.Name + "/" + j.Name + ": U-turn, cannot miter"); return; }
                    var m = sum.Normalized();                      // miter plane normal
                    double tanHalf = Math.Tan(phi / 2);
                    if (tanHalf < 1e-6) { rep.Warn(i.Name + "/" + j.Name + ": corner too sharp to miter"); return; }
                    i.EndShift = i.RadialM / tanHalf + Margin;
                    j.StartShift = j.RadialM / tanHalf + Margin;
                    cuts.Add(new PlaneCut { Member = i, Point = c - m * (g / 2), Normal = m, KeepSign = -1 });
                    cuts.Add(new PlaneCut { Member = j, Point = c + m * (g / 2), Normal = m, KeepSign = +1 });
                    rep.Log(string.Format("corner {0}/{1}: miter, angle {2:0.#}°, extend {3:0.##}/{4:0.##} mm{5}",
                        i.Name, j.Name, phi * 180 / Math.PI, i.EndShift * 1000, j.StartShift * 1000, g > 0 ? ", gap " + (g * 1000).ToString("0.##") + " mm" : ""));
                    break;
                }
                case CornerTreatment.Butt1:
                {
                    // i runs through to j's far face; j is trimmed by i's outer face
                    i.EndShift = Extent(j, di);
                    var u = (dj - di * Vec3.Dot(dj, di)).Normalized();
                    double w = Extent(i, u) + g;
                    j.StartShift = j.RadialM + Margin;
                    cuts.Add(new PlaneCut { Member = j, Point = c + u * w, Normal = u, KeepSign = +1 });
                    rep.Log(string.Format("corner {0}/{1}: end butt 1 ({0} through), {0} end shift {2:0.##} mm, {1} trimmed at {3:0.##} mm",
                        i.Name, j.Name, i.EndShift * 1000, w * 1000));
                    break;
                }
                case CornerTreatment.Butt2:
                {
                    // j runs through; i is trimmed by j's outer face
                    j.StartShift = Extent(i, -dj);
                    var u = (-di - dj * Vec3.Dot(-di, dj)).Normalized();
                    double w = Extent(j, u) + g;
                    i.EndShift = i.RadialM + Margin;
                    cuts.Add(new PlaneCut { Member = i, Point = c + u * w, Normal = u, KeepSign = +1 });
                    rep.Log(string.Format("corner {0}/{1}: end butt 2 ({1} through), {1} start shift {2:0.##} mm, {0} trimmed at {3:0.##} mm",
                        i.Name, j.Name, j.StartShift * 1000, w * 1000));
                    break;
                }
            }
        }

        // ------------------------------------------------------------------ B. planes
        void CreatePlanes(List<PlaneReq> reqs)
        {
            if (reqs.Count == 0) return;
            // The user may still be editing the path sketch (selecting lines from inside it is natural).
            // Insert3DSketch would then EXIT that sketch instead of opening ours, so leave it first.
            if (skm.ActiveSketch != null)
            {
                model.ClearSelection2(true);
                if (skm.ActiveSketch.Is3D()) skm.Insert3DSketch(true); else skm.InsertSketch(true);
                if (skm.ActiveSketch != null) throw new InvalidOperationException("exit the sketch you are editing, then run the command again");
                rep.Log("left the active sketch before building");
            }
            model.ClearSelection2(true);
            skm.Insert3DSketch(true);
            if (skm.ActiveSketch == null || !skm.ActiveSketch.Is3D())
                throw new InvalidOperationException("could not open the construction 3D sketch (is a sketch or another command still active?)");
            bool addToDb = skm.AddToDB;
            skm.AddToDB = true;      // exact coordinates, no inference
            try
            {
                foreach (var r in reqs)
                {
                    var e = r.Point + r.Normal * HelperLineLen;
                    var line = skm.CreateLine(r.Point.X, r.Point.Y, r.Point.Z, e.X, e.Y, e.Z);
                    if (line == null) throw new InvalidOperationException("could not draw a construction line for " + r.Name);
                    line.ConstructionGeometry = true;
                    r.Line = line;
                }
            }
            finally { skm.AddToDB = addToDb; }
            // Capture the sketch as a feature WHILE it is active: in a sheet-metal part the Flat-Pattern
            // folder is always the last feature, so "last feature in the tree" would return that instead.
            var helper = (IFeature)skm.ActiveSketch;
            model.ClearSelection2(true);
            skm.Insert3DSketch(true);
            helper.Name = rep.Prefix + " Construction";
            created.Add(helper);

            foreach (var r in reqs)
            {
                model.ClearSelection2(true);
                var sd0 = (SelectData)selMgr.CreateSelectData(); sd0.Mark = 0;
                if (!r.Line!.Select4(false, sd0)) throw new InvalidOperationException("could not select the construction line for " + r.Name);
                var sp = (ISketchPoint)((ISketchLine)r.Line).GetStartPoint2();
                var sd1 = (SelectData)selMgr.CreateSelectData(); sd1.Mark = 1;
                if (!sp.Select4(true, sd1)) throw new InvalidOperationException("could not select the construction point for " + r.Name);

                const int perp = (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Perpendicular;
                const int coin = (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident;
                const int origin = (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_OptionOriginOnCurve;
                object o = fm.InsertRefPlane(perp | origin, 0, coin, 0, 0, 0);
                if (o == null)
                {
                    model.ClearSelection2(true);
                    r.Line.Select4(false, sd0); sp.Select4(true, sd1);
                    o = fm.InsertRefPlane(perp, 0, coin, 0, 0, 0);
                }
                if (o == null) throw new InvalidOperationException("InsertRefPlane failed for " + r.Name);
                var f = o as IFeature ?? LastFeatureOfType("RefPlane") ?? throw new InvalidOperationException("new plane not found in the tree for " + r.Name);
                f.Name = r.Name;
                r.Feature = f;
                created.Add(f);

                // where did the plane actually land?
                r.ActualNormal = r.Normal;
                try
                {
                    var rp = (IRefPlane)f.GetSpecificFeature2();
                    var data = (double[])rp.Transform.ArrayData;
                    var n = new Vec3(data[6], data[7], data[8]).Normalized();
                    var org = new Vec3(data[9], data[10], data[11]);
                    r.ActualNormal = n;
                    if (Math.Abs(Vec3.Dot(n, r.Normal)) < 0.999) rep.Warn(r.Name + " normal is off: wanted " + r.Normal.Mm() + " got " + n.Mm());
                    if (Vec3.Distance(org, r.Point) > 1e-5) rep.Log(r.Name + ": plane origin at " + org.Mm() + ", point " + r.Point.Mm() + " (profile placed via transform anyway)");
                }
                catch { /* diagnostics only */ }
            }
            model.ClearSelection2(true);
        }

        // ------------------------------------------------------------------ C. members
        void BuildMember(Plan plan, Vec3 planeNormal)
        {
            // profile sketch on the plane
            model.ClearSelection2(true);
            if (!plan.Plane!.Select2(false, 0)) throw new InvalidOperationException("could not select " + plan.Plane.Name);
            skm.InsertSketch(true);
            var sk = skm.ActiveSketch;
            var toSketch = sk.ModelToSketchTransform;
            var pts3 = plan.Mid.Select(m => plan.Start + (plan.P * m.X + plan.Q * m.Y) * Mm).ToList();
            var pts2 = pts3.Select(v => ToSketch(v, toSketch)).ToList();
            double maxZ = pts2.Max(v => Math.Abs(v.Z));
            if (maxZ > 1e-6) rep.Warn(string.Format("{0}: profile points are {1:0.###} mm off the sketch plane", plan.Name, maxZ * 1000));
            for (int k = 0; k + 1 < pts2.Count; k++)
            {
                var s = skm.CreateLine(pts2[k].X, pts2[k].Y, 0, pts2[k + 1].X, pts2[k + 1].Y, 0);
                if (s == null) throw new InvalidOperationException(plan.Name + ": CreateLine failed for profile segment " + (k + 1));
            }
            var sketch = (IFeature)sk;       // captured while active (see CreatePlanes)
            model.ClearSelection2(true);
            skm.InsertSketch(true);
            sketch.Name = rep.Prefix + " Profile" + plan.Index;
            plan.Sketch = sketch;
            created.Add(sketch);

            // base flange: the API-help argument pattern (the only one SolidWorks accepts for an open profile)
            double t = spec.ThicknessMm * Mm, r = spec.RadiusMm * Mm, len = plan.Length;
            // SolidWorks accepts only the API-help argument pattern for an open profile (FlipExtruDir included),
            // so the direction is fixed afterwards through the definition if the plane normal points the wrong way.
            bool flipDir = false;
            if (Vec3.Dot(planeNormal, plan.D) < 0) rep.Log(plan.Name + ": plane normal opposes the path; expect a direction reversal");
            rep.Log(string.Format("{0}: D={1} P={2} Q={3}; sketch pts {4}", plan.Name, Dir(plan.D), Dir(plan.P), Dir(plan.Q),
                string.Join(" ", pts2.Select(v => string.Format("({0:0.##},{1:0.##})", v.X * 1000, v.Y * 1000)))));
            model.ClearSelection2(true);
            sketch.Select2(false, 0);
            var cba = fm.CreateCustomBendAllowance();
            cba.Type = (int)swBendAllowanceTypes_e.swBendAllowanceKFactor;
            cba.KFactor = spec.KFactor;
            var flange = InsertFlange(plan, sketch, t, r, len, flipDir, cba);
            if (flange == null)
                throw new InvalidOperationException(plan.Name + ": Base Flange failed. Check the profile is open and not self-intersecting, and that the bend radius suits the thickness.");
            flange.Name = plan.Name;
            plan.Flange = flange;
            created.Add(flange);

            // where is the material? fix direction and thickness side in one definition change
            var body = BodyOf(plan);
            var com = CoM(body);
            double along = Vec3.Dot(com - plan.Start, plan.D);
            bool wrongDir = along < 0;
            int side = SideOf(plan, com);
            bool wrongSide = side != plan.Side;
            if (wrongDir || wrongSide)
            {
                var data = (IBaseFlangeFeatureData)flange.GetDefinition();
                data.AccessSelections(model, null);
                if (wrongDir) data.ReverseDirection = !data.ReverseDirection;
                if (wrongSide) data.ReverseThickness = !data.ReverseThickness;
                if (!flange.ModifyDefinition(data, model, null))
                {
                    data.ReleaseSelectionAccess();
                    throw new InvalidOperationException(plan.Name + ": could not change the flange definition (direction/thickness side)");
                }
                body = BodyOf(plan);
                com = CoM(body);
                along = Vec3.Dot(com - plan.Start, plan.D);
                side = SideOf(plan, com);
            }
            rep.Log(string.Format("{0}: {1} → {2}, L={3:0.##} mm, centre of mass {4:0.##} mm along{5}, thickness {6}{7}",
                plan.Name, plan.Start.Mm(), plan.End.Mm(), len * 1000, along * 1000,
                wrongDir ? " (direction reversed)" : "", side > 0 ? "inside" : "outside", wrongSide ? " (flipped)" : ""));
            if (along < 0 || side != plan.Side) rep.Warn(plan.Name + ": geometry still not as asked after the definition change");

            // gauge table
            if (!string.IsNullOrEmpty(spec.GaugeTablePath))
            {
                try
                {
                    var data = (IBaseFlangeFeatureData)flange.GetDefinition();
                    data.AccessSelections(model, null);
                    data.UseGaugeTable = true;
                    data.GaugeTablePath = spec.GaugeTablePath;
                    if (!string.IsNullOrEmpty(spec.GaugeName)) data.ThicknessTableName = spec.GaugeName;
                    bool ok = flange.ModifyDefinition(data, model, null);
                    if (!ok) { data.ReleaseSelectionAccess(); rep.Warn(plan.Name + ": gauge table could not be attached (thickness and radius are still as specified)"); }
                    else
                    {
                        var check = (IBaseFlangeFeatureData)flange.GetDefinition();
                        rep.Log(string.Format("{0}: gauge table {1}, gauge '{2}', table thickness {3:0.###} mm",
                            plan.Name, check.UseGaugeTable ? "ON" : "off", check.ThicknessTableName, check.Thickness * 1000));
                    }
                }
                catch (Exception e) { rep.Warn(plan.Name + ": gauge table: " + e.Message); }
            }
        }

        /// <summary>
        /// The Base Flange call, as a ladder of variants: SolidWorks is picky about the argument pattern for an
        /// open profile, and the first body and the following bodies of a multibody sheet-metal part do not
        /// always accept the same call. Logs which rung worked so the ladder can be shortened later.
        /// </summary>
        IFeature? InsertFlange(Plan plan, IFeature sketch, double t, double r, double len, bool flipDir, CustomBendAllowance cba)
        {
            const int blind = (int)swEndConditions_e.swEndCondBlind;
            const int rect = (int)swSheetMetalReliefTypes_e.swSheetMetalReliefRectangular;
            var tried = new List<string>();
            Func<string, Func<IFeature?>, IFeature?> Try = (label, f) =>
            {
                model.ClearSelection2(true);
                sketch.Select2(false, 0);
                IFeature? res = null;
                try { res = f(); tried.Add(label + ":" + (res == null ? "null" : "OK")); }
                catch (Exception e) { tried.Add(label + ":threw " + e.Message); }
                return res;
            };
            IFeature? flange =
                Try("help-pattern", () => fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, false, false, false))
                ?? Try("merge", () => fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, true, false, false))
                ?? Try("scope+auto", () => fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, false, true, true))
                ?? Try("v1-overload", () => fm.InsertSheetMetalBaseFlange(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true))
                ?? Try("dir0", () => fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 0, cba, true, rect, 1, 1, 0.5, true, false, false, false))
                ?? Try("byid", () =>
                {
                    model.ClearSelection2(true);
                    model.Extension.SelectByID2(sketch.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                    return fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, false, false, false);
                })
                ?? Try("rebuild-first", () =>
                {
                    model.ForceRebuild3(false);
                    model.ClearSelection2(true);
                    sketch.Select2(false, 0);
                    return fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, false, false, false);
                })
                ?? Try("hide-others", () =>
                {
                    foreach (var b in Bodies()) b.HideBody(true);
                    model.ClearSelection2(true);
                    sketch.Select2(false, 0);
                    var res = fm.InsertSheetMetalBaseFlange2(t, false, r, len, 0, flipDir, blind, blind, 1, cba, true, rect, 1, 1, 0.5, true, false, false, false);
                    foreach (var b in Bodies()) b.HideBody(false);
                    return res;
                });
            if (flange == null || tried.Count > 1) rep.Log(plan.Name + ": base flange ladder: " + string.Join("; ", tried));
            return flange;
        }

        Vec3 ToSketch(Vec3 v, IMathTransform xf)
        {
            var p = (IMathPoint)mu.CreatePoint(v.ToArray());
            return Vec3.From(((IMathPoint)p.MultiplyTransform(xf)).ArrayData);
        }

        /// <summary>The member's body, re-fetched through its flange's faces (body objects go stale across rebuilds).</summary>
        IBody2 BodyOf(Plan plan)
        {
            var faces = plan.Flange!.GetFaces() as object[];
            if (faces == null || faces.Length == 0) throw new InvalidOperationException(plan.Name + ": flange has no faces (body lost?)");
            var body = ((IFace2)faces[0]).GetBody() as IBody2;
            if (body == null) throw new InvalidOperationException(plan.Name + ": face has no body");
            return body;
        }

        static Vec3 CoM(IBody2 body)
        {
            var mp = (double[])body.GetMassProperties(1.0);
            return new Vec3(mp[0], mp[1], mp[2]);
        }

        /// <summary>+1 if the body's centre of mass sits on the inside-thickness candidate, -1 for outside.</summary>
        int SideOf(Plan plan, Vec3 com)
        {
            double t = spec.ThicknessMm;
            double total = 0; var cIn = new P2(0, 0); var cOut = new P2(0, 0);
            int k = 0;
            for (int s = 0; s + 1 < plan.Mid.Count; s++)
            {
                var a = plan.Mid[s]; var b = plan.Mid[s + 1];
                var len = (b - a).Length; if (len < 1e-9) continue;
                var mid = (a + b) * 0.5;
                var n = plan.InNormal[k++];
                cIn = cIn + (mid + n * (t / 2)) * len;
                cOut = cOut + (mid - n * (t / 2)) * len;
                total += len;
            }
            if (total <= 0) return plan.Side;
            cIn = cIn * (1 / total); cOut = cOut * (1 / total);
            // compare in the cross-section: drop the along-path component
            var rel = com - plan.Start;
            var perp = rel - plan.D * Vec3.Dot(rel, plan.D);
            var vIn = (plan.P * cIn.X + plan.Q * cIn.Y) * Mm;
            var vOut = (plan.P * cOut.X + plan.Q * cOut.Y) * Mm;
            return Vec3.Distance(perp, vIn) <= Vec3.Distance(perp, vOut) ? 1 : -1;
        }

        // ------------------------------------------------------------------ D. corners
        void ApplyCut(PlaneCut cut)
        {
            var plan = cut.Member;
            IBody2 body;
            try { body = BodyOf(plan); }
            catch (Exception e) { rep.Warn(cut.Label + ": " + e.Message); return; }

            model.ClearSelection2(true);
            if (!cut.Plane!.Select2(false, 0)) { rep.Warn(cut.Label + ": could not select " + cut.Plane.Name); return; }
            int err = 0;
            IFeature? f = null;
            var tried = new List<string>();
            // The Bodies argument is a SAFEARRAY of IDispatch: plain object[] is marshalled as VARIANT[] and SolidWorks
            // answers "bodies not specified" (error 1), so wrap each body in a DispatchWrapper. Fallbacks follow.
            var attempts = new List<KeyValuePair<string, Func<IFeature?>>>
            {
                new KeyValuePair<string, Func<IFeature?>>("dispatch[]", () => fm.InsertCutSurface(false, 0, true, false,
                    new System.Runtime.InteropServices.DispatchWrapper[] { new System.Runtime.InteropServices.DispatchWrapper(body) }, out err)),
                new KeyValuePair<string, Func<IFeature?>>("select-body-mark1", () =>
                {
                    var sd = (SelectData)selMgr.CreateSelectData(); sd.Mark = 1;
                    body.Select2(true, sd);
                    return fm.InsertCutSurface(false, 0, true, false, null, out err);
                }),
                new KeyValuePair<string, Func<IFeature?>>("select-body-mark0", () =>
                {
                    var sd = (SelectData)selMgr.CreateSelectData(); sd.Mark = 0;
                    body.Select2(true, sd);
                    return fm.InsertCutSurface(false, 0, true, false, null, out err);
                }),
            };
            foreach (var at in attempts)
            {
                model.ClearSelection2(true);
                if (!cut.Plane.Select2(false, 0)) break;
                body = BodyOf(plan);
                try { f = at.Value(); tried.Add(at.Key + ":" + (f == null ? "err " + err : "OK")); }
                catch (Exception e) { tried.Add(at.Key + ":threw " + e.Message); }
                if (f != null) break;
            }
            if (tried.Count > 1 || f == null) rep.Log(cut.Label + ": cut ladder: " + string.Join("; ", tried));
            if (f == null) { rep.Warn(cut.Label + ": Cut With Surface failed (error " + err + ") on " + plan.Name); return; }
            f.Name = cut.Label;
            created.Add(f);
            rep.Cuts++;

            // kept the right side?
            var com = CoM(BodyOf(plan));
            int s = Math.Sign(Vec3.Dot(com - cut.Point, cut.Normal));
            bool flipped = false;
            if (s != cut.KeepSign)
            {
                var data = (ISurfaceCutFeatureData)f.GetDefinition();
                data.AccessSelections(model, null);
                data.Flip = !data.Flip;
                if (!f.ModifyDefinition(data, model, null)) { data.ReleaseSelectionAccess(); rep.Warn(cut.Label + ": could not flip the cut"); return; }
                flipped = true;
                com = CoM(BodyOf(plan));
                s = Math.Sign(Vec3.Dot(com - cut.Point, cut.Normal));
            }
            rep.Log(string.Format("{0}: {1} cut at {2}, normal {3}{4}{5}", cut.Label, plan.Name, cut.Point.Mm(), cut.Normal.Mm(),
                flipped ? " (flipped)" : "", s != cut.KeepSign ? " — WRONG SIDE KEPT" : ""));
            if (s != cut.KeepSign) rep.Warn(cut.Label + ": the wrong side was kept");
        }

        // ------------------------------------------------------------------ E. tidy
        void Tidy(List<PlaneReq> reqs)
        {
            if (spec.FolderFeatures) MakeFolders();
            if (spec.HideConstruction)
            {
                try
                {
                    foreach (var r in reqs)
                    {
                        model.ClearSelection2(true);
                        if (r.Feature != null && r.Feature.Select2(false, 0)) model.BlankRefGeom();
                    }
                    var helper = created.FirstOrDefault(c => c.Name.EndsWith(" Construction"));
                    model.ClearSelection2(true);
                    if (helper != null && helper.Select2(false, 0)) model.BlankSketch();
                }
                catch (Exception e) { rep.Log("hide construction: " + e.Message); }
                model.ClearSelection2(true);
            }
            foreach (var f in created) rep.Features.Add(f.Name);
        }

        /// <summary>
        /// Two feature-tree folders: the construction run (helper sketch + planes) and the member run
        /// (flanges + corner cuts). They cannot be one folder: SolidWorks inserts its own Sheet-Metal&lt;n&gt;
        /// features between them and those cannot be foldered. Profile sketches are absorbed, never selected.
        /// </summary>
        void MakeFolders()
        {
            var mine = new HashSet<string>(created.Select(c => c.Name));
            var top = new List<IFeature>();
            for (var tf = model.FirstFeature() as IFeature; tf != null; tf = tf.GetNextFeature() as IFeature) top.Add(tf);
            var firstFlange = top.FindIndex(f => mine.Contains(f.Name) && f.GetTypeName2() == "SMBaseFlange");
            var runs = new List<KeyValuePair<string, List<IFeature>>>();
            if (firstFlange >= 0)
            {
                var members = top.Skip(firstFlange).Where(f => mine.Contains(f.Name)).ToList();
                runs.Add(new KeyValuePair<string, List<IFeature>>(rep.Prefix + " " + spec.Name, members));
                var constr = top.Take(firstFlange).Where(f => mine.Contains(f.Name) && f.GetTypeName2() != "ProfileFeature").ToList();
                runs.Add(new KeyValuePair<string, List<IFeature>>(rep.Prefix + " Construction", constr));
            }
            foreach (var run in runs)
            {
                if (run.Value.Count < 2) continue;
                try
                {
                    model.ClearSelection2(true);
                    bool any = false;
                    foreach (var f in run.Value) any |= f.Select2(true, 0);
                    if (!any) continue;
                    var folder = fm.InsertFeatureTreeFolder2((int)swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing);
                    if (folder != null) { folder.Name = run.Key; rep.Log("folder '" + run.Key + "': " + run.Value.Count + " features"); }
                    else rep.Log("folder '" + run.Key + "' not created (" + run.Value.Count + " features stay loose)");
                }
                catch (Exception e) { rep.Log("folder '" + run.Key + "': " + e.Message); }
            }
            model.ClearSelection2(true);
        }

        // ------------------------------------------------------------------ helpers
        List<IBody2> Bodies()
        {
            var list = new List<IBody2>();
            var arr = part.GetBodies2((int)swBodyType_e.swSolidBody, false) as object[];
            if (arr != null) foreach (var o in arr) { var b = o as IBody2; if (b != null) list.Add(b); }
            return list;
        }

        IFeature? LastFeatureOfType(string typeName)
        {
            IFeature? last = null;
            for (var f = model.FirstFeature() as IFeature; f != null; f = f.GetNextFeature() as IFeature)
                if (f.GetTypeName2() == typeName) last = f;
            return last;
        }

        string NextPrefix()
        {
            int max = 0;
            for (var f = model.FirstFeature() as IFeature; f != null; f = f.GetNextFeature() as IFeature)
            {
                var n = f.Name ?? "";
                if (!n.StartsWith("SSM")) continue;
                int i = 3; int v = 0; bool digits = false;
                while (i < n.Length && char.IsDigit(n[i])) { v = v * 10 + (n[i] - '0'); i++; digits = true; }
                if (digits) max = Math.Max(max, v);
            }
            return "SSM" + (max + 1);
        }
    }
}
