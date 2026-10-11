using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PolyApp
{
    /// <summary>
    /// Icon strips for the CommandManager, built at startup from the catalogue so a new command never
    /// needs artwork to show up. A command uses AddIn\Icons\&lt;Id&gt;_&lt;size&gt;.png (or &lt;Id&gt;.png, scaled)
    /// when present, otherwise a coloured tile with its Code. Strips land in %LOCALAPPDATA%\PolyApp\icons.
    /// </summary>
    static class IconFactory
    {
        public static readonly int[] Sizes = { 20, 32, 40, 64, 96, 128 };

        public static string IconDir { get { return Path.Combine(SwAddin.BaseDir, "Icons"); } }
        public static string OutDir { get { return Path.Combine(Log.Dir, "icons"); } }

        /// <summary>One horizontal strip per size, the commands in the given order.</summary>
        public static string[] Strip(string name, IList<Command> cmds)
        {
            Directory.CreateDirectory(OutDir);
            var paths = new List<string>();
            foreach (var size in Sizes)
            {
                using (var bmp = new Bitmap(Math.Max(1, cmds.Count) * size, size, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    for (int i = 0; i < cmds.Count; i++)
                        using (var tile = Tile(cmds[i], size)) g.DrawImage(tile, i * size, 0, size, size);
                    var path = Path.Combine(OutDir, Safe(name) + "_" + size + ".png");
                    bmp.Save(path, ImageFormat.Png);
                    paths.Add(path);
                }
            }
            return paths.ToArray();
        }

        /// <summary>The single icon of a group or flyout button.</summary>
        public static string[] Main(string name, string code, string box)
        {
            Directory.CreateDirectory(OutDir);
            var paths = new List<string>();
            foreach (var size in Sizes)
            {
                var custom = Find(name, size);
                using (var bmp = custom ?? DrawTile(code, BoxColor(box), size))
                {
                    var path = Path.Combine(OutDir, Safe(name) + "_main_" + size + ".png");
                    bmp.Save(path, ImageFormat.Png);
                    paths.Add(path);
                }
            }
            return paths.ToArray();
        }

        static Bitmap Tile(Command c, int size)
        {
            return Find(c.Id, size) ?? DrawTile(c.Code, BoxColor(c.Box), size);
        }

        /// <summary>A hand-made icon from AddIn\Icons: exact size first, then any size scaled.</summary>
        static Bitmap? Find(string id, int size)
        {
            try
            {
                var exact = Path.Combine(IconDir, id + "_" + size + ".png");
                if (File.Exists(exact)) using (var img = Image.FromFile(exact)) return new Bitmap(img, size, size);
                var any = Path.Combine(IconDir, id + ".png");
                if (File.Exists(any)) using (var img = Image.FromFile(any)) return new Bitmap(img, size, size);
            }
            catch (Exception e) { Log.Write("icon " + id + ": " + e.Message); }
            return null;
        }

        static Bitmap DrawTile(string code, Color color, int size)
        {
            var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                float pad = Math.Max(1f, size * 0.06f);
                float r = Math.Max(2f, size * 0.18f);
                using (var path = Rounded(new RectangleF(pad, pad, size - 2 * pad, size - 2 * pad), r))
                using (var brush = new SolidBrush(color))
                    g.FillPath(brush, path);

                var text = (code ?? "").Trim();
                if (text.Length == 0) return bmp;
                float scale = text.Length <= 2 ? 0.46f : text.Length == 3 ? 0.32f : 0.25f;
                using (var font = new Font("Segoe UI", Math.Max(5f, size * scale), FontStyle.Bold, GraphicsUnit.Pixel))
                using (var fmt = new StringFormat(StringFormatFlags.NoWrap | StringFormatFlags.NoClip) { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.None })
                    g.DrawString(text, font, Brushes.White, new RectangleF(0, 0, size, size), fmt);
            }
            return bmp;
        }

        static GraphicsPath Rounded(RectangleF rc, float r)
        {
            var p = new GraphicsPath();
            float d = r * 2;
            p.AddArc(rc.X, rc.Y, d, d, 180, 90);
            p.AddArc(rc.Right - d, rc.Y, d, d, 270, 90);
            p.AddArc(rc.Right - d, rc.Bottom - d, d, d, 0, 90);
            p.AddArc(rc.X, rc.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        public static Color BoxColor(string box)
        {
            switch (box)
            {
                case "PDM": return Color.FromArgb(0x2E, 0x6F, 0xBF);
                case "Make": return Color.FromArgb(0x3A, 0x9D, 0x5D);
                case "Select": return Color.FromArgb(0xC8, 0x77, 0x1A);
                case "Export": return Color.FromArgb(0x7A, 0x4F, 0xBF);
                case "Drawing": return Color.FromArgb(0x2A, 0x8C, 0x8C);
                case "Setup": return Color.FromArgb(0x44, 0x44, 0x44);
                default: return Color.FromArgb(0x6B, 0x72, 0x80);
            }
        }

        static string Safe(string name) { return Regex.Replace(name, @"[^A-Za-z0-9]+", "_"); }
    }
}
