using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace StructuralSheetMetal
{
    /// <summary>An ordered, consistently oriented run of connected segments.</summary>
    public sealed class Chain
    {
        public List<PathSegment> Segments = new List<PathSegment>();
        public bool Closed;
        public override string ToString() { return Segments.Count + " segment(s)" + (Closed ? ", closed" : ""); }
    }

    /// <summary>Turns SolidWorks selections (sketch lines, linear edges) into path segments and chains them.</summary>
    public static class PathChain
    {
        /// <summary>Endpoint coincidence tolerance, metres (1 µm).</summary>
        public const double Tol = 1e-6;

        /// <summary>All selected entities with the given mark (-1 = any mark) that are straight lines/edges.</summary>
        public static List<PathSegment> FromSelection(ISldWorks sw, IModelDoc2 model, int mark, List<string>? rejected)
        {
            var selMgr = (ISelectionMgr)model.SelectionManager;
            int n = selMgr.GetSelectedObjectCount2(mark);
            var list = new List<PathSegment>();
            for (int i = 1; i <= n; i++)
            {
                object ent = selMgr.GetSelectedObject6(i, mark);
                string why;
                var seg = FromEntity(sw, ent, out why);
                if (seg != null) list.Add(seg);
                else if (rejected != null) rejected.Add(string.Format("selection {0}: {1}", i, why));
            }
            return list;
        }

        public static List<PathSegment> FromEntities(ISldWorks sw, IEnumerable<object> entities, List<string>? rejected)
        {
            var list = new List<PathSegment>();
            int i = 0;
            foreach (var ent in entities)
            {
                i++;
                string why;
                var seg = FromEntity(sw, ent, out why);
                if (seg != null) list.Add(seg);
                else if (rejected != null) rejected.Add(string.Format("entity {0}: {1}", i, why));
            }
            return list;
        }

        /// <summary>A sketch line (2D or 3D sketch) or a linear model edge → segment in model metres.</summary>
        public static PathSegment? FromEntity(ISldWorks sw, object entity, out string why)
        {
            why = "";
            if (entity == null) { why = "null"; return null; }

            var sseg = entity as ISketchSegment;
            if (sseg != null)
            {
                if (sseg.GetType() != (int)swSketchSegments_e.swSketchLINE) { why = "sketch segment is not a line (arcs come in a later phase)"; return null; }
                var line = (ISketchLine)sseg;
                var sp = (ISketchPoint)line.GetStartPoint2();
                var ep = (ISketchPoint)line.GetEndPoint2();
                var sketch = (ISketch)sseg.GetSketch();
                var mu = (IMathUtility)sw.GetMathUtility();
                IMathTransform? toModel = null;
                Vec3? normal = null;
                var xf = sketch.ModelToSketchTransform;
                if (xf != null)
                {
                    toModel = (IMathTransform)xf.Inverse();
                    if (!sketch.Is3D())
                    {
                        var nv = (IMathVector)mu.CreateVector(new double[] { 0, 0, 1 });
                        normal = Vec3.From(((IMathVector)nv.MultiplyTransform(toModel)).ArrayData).Normalized();
                    }
                }
                var a = Point(mu, toModel, sp.X, sp.Y, sp.Z);
                var b = Point(mu, toModel, ep.X, ep.Y, ep.Z);
                if (Vec3.Distance(a, b) < Tol) { why = "zero-length line"; return null; }
                string name = "";
                try { name = sseg.GetName(); } catch { }
                return new PathSegment(a, b) { SwEntity = sseg, SketchNormal = normal, Label = name };
            }

            var edge = entity as IEdge;
            if (edge != null)
            {
                var curve = (ICurve)edge.GetCurve();
                if (curve == null || !curve.IsLine()) { why = "edge is not straight"; return null; }
                var sv = edge.GetStartVertex() as IVertex;
                var ev = edge.GetEndVertex() as IVertex;
                if (sv == null || ev == null) { why = "edge has no end vertices"; return null; }
                var a = Vec3.From(sv.GetPoint());
                var b = Vec3.From(ev.GetPoint());
                if (Vec3.Distance(a, b) < Tol) { why = "zero-length edge"; return null; }
                return new PathSegment(a, b) { SwEntity = edge, Label = "Edge" };
            }

            why = "not a sketch line or model edge";
            return null;
        }

        static Vec3 Point(IMathUtility mu, IMathTransform? toModel, double x, double y, double z)
        {
            if (toModel == null) return new Vec3(x, y, z);
            var p = (IMathPoint)mu.CreatePoint(new double[] { x, y, z });
            return Vec3.From(((IMathPoint)p.MultiplyTransform(toModel)).ArrayData);
        }

        /// <summary>
        /// Groups segments into chains by shared endpoints and orients every segment so that
        /// segment[i].B == segment[i+1].A. A chain whose last end meets its first start is Closed.
        /// Branching (three segments at one point) follows the first match; the rest start new chains.
        /// </summary>
        public static List<Chain> Chain(IEnumerable<PathSegment> input)
        {
            var pool = input.ToList();
            var chains = new List<Chain>();
            while (pool.Count > 0)
            {
                var run = new List<PathSegment> { pool[0] };
                pool.RemoveAt(0);
                // forward
                bool grew = true;
                while (grew)
                {
                    grew = false;
                    var tail = run[run.Count - 1].B;
                    for (int i = 0; i < pool.Count; i++)
                    {
                        var s = pool[i];
                        if (Vec3.Near(s.A, tail, Tol)) { run.Add(s); pool.RemoveAt(i); grew = true; break; }
                        if (Vec3.Near(s.B, tail, Tol)) { run.Add(s.Reversed()); pool.RemoveAt(i); grew = true; break; }
                    }
                }
                // backward
                grew = true;
                while (grew)
                {
                    grew = false;
                    var head = run[0].A;
                    if (Vec3.Near(run[run.Count - 1].B, head, Tol) && run.Count >= 3) break; // already closed
                    for (int i = 0; i < pool.Count; i++)
                    {
                        var s = pool[i];
                        if (Vec3.Near(s.B, head, Tol)) { run.Insert(0, s); pool.RemoveAt(i); grew = true; break; }
                        if (Vec3.Near(s.A, head, Tol)) { run.Insert(0, s.Reversed()); pool.RemoveAt(i); grew = true; break; }
                    }
                }
                var c = new Chain { Segments = run };
                c.Closed = run.Count >= 3 && Vec3.Near(run[run.Count - 1].B, run[0].A, Tol);
                chains.Add(c);
            }
            return chains;
        }
    }
}
