using System;
using System.Collections.Generic;
using System.Text;

namespace StructuralSheetMetal
{
    /// <summary>Which side of the open profile the sheet thickness goes.</summary>
    public enum MaterialSide
    {
        /// <summary>Toward the concave side of the profile (the L's inner corner). Two members with the same
        /// path and profile, one Inside and one Outside, touch along the profile: the outside one "hugs" the inside one.</summary>
        Inside = 0,
        Outside = 1
    }

    /// <summary>Corner treatment between consecutive path segments of one group (Structural Member's vocabulary).</summary>
    public enum CornerTreatment
    {
        None = 0,
        /// <summary>Both members cut on the bisecting plane.</summary>
        Miter = 1,
        /// <summary>The FIRST member runs through to the second's outer face; the second butts against it.</summary>
        Butt1 = 2,
        /// <summary>The SECOND member runs through; the first butts against it.</summary>
        Butt2 = 3
    }

    /// <summary>A straight path segment in model metres. SwEntity is the ISketchSegment / IEdge it came from (may be null).</summary>
    public sealed class PathSegment
    {
        public Vec3 A, B;
        public object? SwEntity;
        /// <summary>Normal of the 2D sketch the segment lives in, if it came from one. The default "up" for the profile.</summary>
        public Vec3? SketchNormal;
        public string Label = "";

        public PathSegment(Vec3 a, Vec3 b) { A = a; B = b; }
        public Vec3 Dir { get { return (B - A).Normalized(); } }
        public double Length { get { return (B - A).Length; } }
        public PathSegment Reversed() { return new PathSegment(B, A) { SwEntity = SwEntity, SketchNormal = SketchNormal, Label = Label }; }
        public override string ToString() { return (Label.Length > 0 ? Label + " " : "") + A.Mm() + " -> " + B.Mm(); }
    }

    /// <summary>One group of connected path segments sharing settings (like a Structural Member group).</summary>
    public sealed class MemberGroup
    {
        public string Name = "Group";
        public List<PathSegment> Segments = new List<PathSegment>();
        public MaterialSide Side = MaterialSide.Inside;
        public double RotationDeg = 0;
        public bool MirrorHorizontal, MirrorVertical;
        /// <summary>Named point of the profile that sits on the path; null = the profile's default.</summary>
        public string? PiercePoint;
        public CornerTreatment Corner = CornerTreatment.Miter;
        public double GapMm = 0;
        /// <summary>Explicit "up" direction for the profile's +Y (from an alignment selection); null = sketch normal / model Z.</summary>
        public Vec3? AlignUp;
    }

    /// <summary>Everything the builder needs. Units: mm in the spec, metres in the model.</summary>
    public sealed class MemberSpec
    {
        public string Name = "Structural Sheet Metal";
        public Profile Profile = ProfileShapes.Build("angle", new Dictionary<string, double>(), null);
        public double ThicknessMm = 1.5;
        /// <summary>Inside bend radius; null = thickness (or the gauge table's smallest usable radius).</summary>
        public double? BendRadiusMm;
        public double KFactor = 0.5;
        /// <summary>Gauge table to attach to each base flange, and the gauge row to pick in it. Optional.</summary>
        public string? GaugeTablePath;
        public string? GaugeName;
        public List<MemberGroup> Groups = new List<MemberGroup>();
        public bool HideConstruction = true;
        public bool FolderFeatures = true;

        public double RadiusMm { get { return BendRadiusMm ?? ThicknessMm; } }

        public void Validate()
        {
            if (Profile == null || Profile.Points.Count < 2) throw new InvalidOperationException("profile needs at least two points");
            if (ThicknessMm <= 0) throw new InvalidOperationException("thickness must be positive");
            if (RadiusMm < 0) throw new InvalidOperationException("bend radius cannot be negative");
            if (Groups.Count == 0) throw new InvalidOperationException("no groups");
            int n = 0;
            foreach (var g in Groups) n += g.Segments.Count;
            if (n == 0) throw new InvalidOperationException("no path segments selected");
        }
    }

    public sealed class BuildReport
    {
        public bool Ok;
        public string Prefix = "";
        public List<string> Lines = new List<string>();
        public List<string> Features = new List<string>();
        public int Members, Cuts, Warnings;
        public Exception? Error;

        public void Log(string s) { Lines.Add(s); }
        public void Warn(string s) { Warnings++; Lines.Add("WARNING: " + s); }

        public string Summary
        {
            get
            {
                if (!Ok) return "FAILED: " + (Error != null ? Error.Message : "unknown error");
                return string.Format("{0}: {1} member(s), {2} corner cut(s){3}", Prefix, Members, Cuts, Warnings > 0 ? ", " + Warnings + " warning(s)" : "");
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            foreach (var l in Lines) sb.AppendLine(l);
            sb.Append(Summary);
            return sb.ToString();
        }
    }
}
