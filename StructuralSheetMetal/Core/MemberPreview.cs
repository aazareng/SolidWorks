using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace StructuralSheetMetal
{
    /// <summary>
    /// Preview of the members as TEMPORARY bodies (no features): one thin box per profile leg per member,
    /// built with the modeler (planar sheet trimmed by four lines, extruded along the path) and shown with
    /// IBody2.Display3. Bends and corner treatments are left out: it shows where each profile sits and which
    /// way it faces, which is what you need before pressing OK. Clear() removes everything.
    /// </summary>
    public sealed class MemberPreview : IDisposable
    {
        public static int ColorOther = Rgb(255, 214, 80);     // SolidWorks-ish preview yellow
        public static int ColorCurrent = Rgb(80, 200, 255);   // the group being edited

        readonly ISldWorks sw;
        readonly IModelDoc2 model;
        readonly IModeler modeler;
        readonly IMathUtility mu;
        readonly List<IBody2> bodies = new List<IBody2>();

        public int Members;
        public List<string> Problems = new List<string>();

        static int Rgb(int r, int g, int b) { return r | (g << 8) | (b << 16); }   // COLORREF

        public MemberPreview(ISldWorks sw, IModelDoc2 model)
        {
            this.sw = sw; this.model = model;
            modeler = (IModeler)sw.GetModeler();
            mu = (IMathUtility)sw.GetMathUtility();
        }

        /// <summary>Rebuild the preview for the whole spec; the group at highlightGroup is drawn in ColorCurrent.</summary>
        public void Show(MemberSpec spec, int highlightGroup)
        {
            Clear();
            Members = 0;
            Problems.Clear();
            if (spec.Profile == null || spec.Profile.Points.Count < 2 || spec.ThicknessMm <= 0) return;
            for (int gi = 0; gi < spec.Groups.Count; gi++)
            {
                var group = spec.Groups[gi];
                int color = gi == highlightGroup ? ColorCurrent : ColorOther;
                foreach (var chain in PathChain.Chain(group.Segments))
                    foreach (var seg in chain.Segments)
                    {
                        MemberFrame f;
                        try { f = MemberFrame.Compute(spec, group, seg); }
                        catch (Exception e) { Problems.Add(e.Message); continue; }
                        for (int k = 0; k + 1 < f.Mid.Count; k++)
                        {
                            var body = Leg(f, seg, k, spec.ThicknessMm);
                            if (body == null) continue;
                            try { body.Display3(model, color, (int)swTempBodySelectOptions_e.swTempBodySelectOptionNone); bodies.Add(body); }
                            catch (Exception e) { Problems.Add("display: " + e.Message); }
                        }
                        Members++;
                    }
            }
            try { model.GraphicsRedraw2(); } catch { }
        }

        IBody2? Leg(MemberFrame f, PathSegment seg, int k, double t)
        {
            string step = "setup";
            try
            {
                var a = f.Mid[k]; var b = f.Mid[k + 1]; var n = f.InNormal[k];
                if ((b - a).Length < 1e-9) return null;
                var off = n * (t * f.Side);
                var pts = new[] { a, b, b + off, a + off }.Select(v => f.ToModel(seg.A, v)).ToArray();

                step = "CreatePlanarSurface2";
                var surf = modeler.CreatePlanarSurface2(seg.A.ToArray(), f.D.ToArray(), f.P.ToArray()) as ISurface;
                if (surf == null) { Problems.Add("planar surface failed"); return null; }
                var curves = new object[4];
                for (int i = 0; i < 4; i++)
                {
                    var p0 = pts[i]; var p1 = pts[(i + 1) % 4];
                    step = "CreateLine " + i;
                    var line = modeler.CreateLine(p0.ToArray(), (p1 - p0).Normalized().ToArray()) as ICurve;
                    if (line == null) { Problems.Add("line failed"); return null; }
                    step = "CreateTrimmedCurve2 " + i;
                    curves[i] = line.CreateTrimmedCurve2(p0.X, p0.Y, p0.Z, p1.X, p1.Y, p1.Z);
                }
                IBody2? sheet = null;
                var wrapped = curves.Select(c => new DispatchWrapper(c)).ToArray();
                var typed = curves.Cast<Curve>().ToArray();
                var ladder = new List<KeyValuePair<string, Func<object?>>>
                {
                    new KeyValuePair<string, Func<object?>>("CreateTrimmedSheet4 dispatch[]", () => surf.CreateTrimmedSheet4(wrapped, true)),
                    new KeyValuePair<string, Func<object?>>("CreateTrimmedSheet4 Curve[]", () => surf.CreateTrimmedSheet4(typed, true)),
                    new KeyValuePair<string, Func<object?>>("CreateTrimmedSheet4 object[]", () => surf.CreateTrimmedSheet4(curves, true)),
                    new KeyValuePair<string, Func<object?>>("CreateTrimmedSheet5 dispatch[]", () => surf.CreateTrimmedSheet5(wrapped, true, 1e-6)),
                    new KeyValuePair<string, Func<object?>>("CreateTrimmedSheet dispatch[]", () => surf.CreateTrimmedSheet(wrapped)),
                };
                var tried = new List<string>();
                foreach (var rung in ladder)
                {
                    step = rung.Key;
                    try { sheet = rung.Value() as IBody2; tried.Add(rung.Key + (sheet == null ? ":null" : ":OK")); }
                    catch (Exception e) { tried.Add(rung.Key + ":" + e.Message); }
                    if (sheet != null) break;
                }
                if (sheet == null) { Problems.Add("trimmed sheet failed: " + string.Join("; ", tried)); return null; }
                if (tried.Count > 1) Problems.Add("sheet ladder: " + string.Join("; ", tried));
                step = "CreateExtrudedBody";
                var dir = (IMathVector)mu.CreateVector(f.D.ToArray());
                var solid = modeler.CreateExtrudedBody((Body2)sheet, (MathVector)dir, seg.Length) as IBody2;
                if (solid == null) { Problems.Add("extrude failed"); return null; }
                return solid;
            }
            catch (Exception e) { Problems.Add(step + ": " + e.Message); return null; }
        }

        public void Clear()
        {
            foreach (var b in bodies)
            {
                try { b.Hide(model); } catch { }
                try { Marshal.FinalReleaseComObject(b); } catch { }
            }
            bodies.Clear();
            try { model.GraphicsRedraw2(); } catch { }
        }

        public void Dispose() { Clear(); }
    }
}
