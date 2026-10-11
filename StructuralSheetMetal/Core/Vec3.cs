using System;

namespace StructuralSheetMetal
{
    /// <summary>Plain 3-vector in model METRES (the SolidWorks API unit).</summary>
    public struct Vec3
    {
        public double X, Y, Z;
        public Vec3(double x, double y, double z) { X = x; Y = y; Z = z; }

        public static readonly Vec3 Zero = new Vec3(0, 0, 0);
        public static readonly Vec3 UnitX = new Vec3(1, 0, 0);
        public static readonly Vec3 UnitY = new Vec3(0, 1, 0);
        public static readonly Vec3 UnitZ = new Vec3(0, 0, 1);

        public static Vec3 operator +(Vec3 a, Vec3 b) { return new Vec3(a.X + b.X, a.Y + b.Y, a.Z + b.Z); }
        public static Vec3 operator -(Vec3 a, Vec3 b) { return new Vec3(a.X - b.X, a.Y - b.Y, a.Z - b.Z); }
        public static Vec3 operator -(Vec3 a) { return new Vec3(-a.X, -a.Y, -a.Z); }
        public static Vec3 operator *(Vec3 a, double s) { return new Vec3(a.X * s, a.Y * s, a.Z * s); }
        public static Vec3 operator *(double s, Vec3 a) { return a * s; }
        public static Vec3 operator /(Vec3 a, double s) { return new Vec3(a.X / s, a.Y / s, a.Z / s); }

        public static double Dot(Vec3 a, Vec3 b) { return a.X * b.X + a.Y * b.Y + a.Z * b.Z; }
        public static Vec3 Cross(Vec3 a, Vec3 b)
        {
            return new Vec3(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
        }
        public double Length { get { return Math.Sqrt(X * X + Y * Y + Z * Z); } }
        public Vec3 Normalized()
        {
            var l = Length;
            if (l < 1e-15) throw new InvalidOperationException("cannot normalise a zero vector");
            return this / l;
        }
        public static double Distance(Vec3 a, Vec3 b) { return (a - b).Length; }
        public static bool Near(Vec3 a, Vec3 b, double tol) { return Distance(a, b) <= tol; }

        /// <summary>Any unit vector perpendicular to this one.</summary>
        public Vec3 AnyPerpendicular()
        {
            var d = Normalized();
            var t = Math.Abs(d.Z) < 0.9 ? UnitZ : UnitX;
            return Cross(d, t).Normalized();
        }

        public double[] ToArray() { return new[] { X, Y, Z }; }

        /// <summary>From a SolidWorks double[] (at least 3 entries).</summary>
        public static Vec3 From(object arr)
        {
            var a = arr as double[];
            if (a == null || a.Length < 3) throw new ArgumentException("expected a double[3] from SolidWorks");
            return new Vec3(a[0], a[1], a[2]);
        }

        public string Mm() { return string.Format("({0:0.###}, {1:0.###}, {2:0.###}) mm", X * 1000, Y * 1000, Z * 1000); }
        public override string ToString() { return Mm(); }
    }
}
