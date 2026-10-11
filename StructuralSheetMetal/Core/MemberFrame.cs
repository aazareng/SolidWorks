using System;
using System.Collections.Generic;
using System.Linq;

namespace StructuralSheetMetal
{
    /// <summary>
    /// Where a member's profile sits in model space: D along the path, Q "up", P = D × Q, rotated by the
    /// group's angle; the mirrored profile; the pierce point; and the thickened section as 2D data in that
    /// frame. Shared by the builder (real features) and the preview (temporary bodies) so both agree.
    /// </summary>
    public sealed class MemberFrame
    {
        const double Mm = 0.001;

        public Vec3 D, P, Q;
        public Profile Prof = null!;
        public P2 Pierce;
        /// <summary>+1 inside, -1 outside.</summary>
        public int Side;
        /// <summary>Mid-surface points relative to the pierce point, mm.</summary>
        public List<P2> Mid = new List<P2>();
        /// <summary>Per profile segment, unit normal toward the inside (mm space).</summary>
        public List<P2> InNormal = new List<P2>();
        /// <summary>Mid-surface plus the offset by the sheet thickness on the chosen side, mm.</summary>
        public List<P2> Cloud = new List<P2>();
        /// <summary>Largest reach of the thickened section from the path point, metres.</summary>
        public double RadialM;

        public static MemberFrame Compute(MemberSpec spec, MemberGroup group, PathSegment seg)
        {
            var f = new MemberFrame();
            f.Prof = spec.Profile.Mirrored(group.MirrorHorizontal, group.MirrorVertical);
            f.Pierce = f.Prof.Pierce(group.PiercePoint ?? f.Prof.DefaultPierce);
            f.Side = group.Side == MaterialSide.Inside ? 1 : -1;

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
            f.D = d; f.P = p; f.Q = q;

            double t = spec.ThicknessMm;
            int sgn = f.Prof.SignedArea() >= 0 ? 1 : -1;
            f.Mid = f.Prof.Points.Select(x => x - f.Pierce).ToList();
            for (int k = 0; k + 1 < f.Mid.Count; k++)
            {
                var a = f.Mid[k]; var b = f.Mid[k + 1];
                var dir = b - a; var len = dir.Length;
                if (len < 1e-9) { f.InNormal.Add(new P2(0, 0)); continue; }
                var n = new P2(-dir.Y / len, dir.X / len) * sgn;   // toward the inside
                f.InNormal.Add(n);
                var off = n * (t * f.Side);
                f.Cloud.Add(a); f.Cloud.Add(b); f.Cloud.Add(a + off); f.Cloud.Add(b + off);
            }
            f.RadialM = f.Cloud.Count > 0 ? f.Cloud.Max(c => c.Length) * Mm : 0;
            return f;
        }

        /// <summary>A 2D point of the section (mm, relative to the pierce point) → model metres at the given path point.</summary>
        public Vec3 ToModel(Vec3 origin, P2 v) { return origin + (P * v.X + Q * v.Y) * Mm; }
    }
}
