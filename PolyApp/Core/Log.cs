using System;
using System.IO;

namespace PolyApp
{
    /// <summary>One append-only log for the add-in and the CLI: %LOCALAPPDATA%\PolyApp\PolyApp.log.</summary>
    public static class Log
    {
        public static string Dir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PolyApp"); }
        }

        public static string Path_ { get { return Path.Combine(Dir, "PolyApp.log"); } }

        /// <summary>Set by the CLI so log lines also reach the console.</summary>
        public static Action<string>? Echo;

        public static void Write(string text)
        {
            var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + text;
            try
            {
                Directory.CreateDirectory(Dir);
                File.AppendAllText(Path_, line + Environment.NewLine);
            }
            catch { }
            try { Echo?.Invoke(line); } catch { }
        }
    }
}
