using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PolyApp
{
    /// <summary>
    /// PolyApp.ini, layered: compiled defaults, the file beside the DLL/exe, %APPDATA%\PolyApp\PolyApp.ini,
    /// then the file named by POLYAPP_INI. Keys are case-insensitive. %Name% expands to a key in the same
    /// section, then in [Paths], then to an environment variable.
    /// </summary>
    public sealed class Settings
    {
        public const string FileName = "PolyApp.ini";

        readonly Dictionary<string, Dictionary<string, string>> data = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, string> origin = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);   // "section/key" -> file

        public List<string> Files { get; } = new List<string>();

        // ------------------------------------------------------------------ loading
        /// <summary>Load the layered configuration; <paramref name="baseDir"/> is the folder holding the shipped ini.</summary>
        public static Settings Load(string baseDir)
        {
            var s = new Settings();
            s.Merge(Defaults, "(defaults)");
            var candidates = new List<string>
            {
                Path.Combine(baseDir, FileName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PolyApp", FileName),
            };
            var env = Environment.GetEnvironmentVariable("POLYAPP_INI");
            if (!string.IsNullOrWhiteSpace(env)) candidates.Add(env!);
            foreach (var f in candidates)
            {
                if (!File.Exists(f)) continue;
                try { s.Merge(File.ReadAllText(f), f); s.Files.Add(f); }
                catch (Exception e) { Log.Write("settings: could not read " + f + ": " + e.Message); }
            }
            return s;
        }

        void Merge(string text, string source)
        {
            string section = "";
            foreach (var raw in text.Replace("\r", "").Split('\n'))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line[0] == ';' || line[0] == '#') continue;
                if (line[0] == '[' && line[line.Length - 1] == ']') { section = line.Substring(1, line.Length - 2).Trim(); continue; }
                int eq = line.IndexOf('=');
                if (eq < 0) continue;
                var key = line.Substring(0, eq).Trim();
                var val = line.Substring(eq + 1).Trim();
                Set(section, key, val, source);
            }
        }

        public void Set(string section, string key, string value, string source = "(runtime)")
        {
            if (!data.TryGetValue(section, out var sec)) data[section] = sec = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            sec[key] = value;
            origin[section + "/" + key] = source;
        }

        // ------------------------------------------------------------------ reading
        public string Raw(string section, string key)
        {
            return data.TryGetValue(section, out var sec) && sec.TryGetValue(key, out var v) ? v : "";
        }

        public string Get(string section, string key, string fallback = "")
        {
            var v = Raw(section, key);
            return v.Length == 0 ? fallback : Expand(section, v, 0);
        }

        public bool GetBool(string section, string key, bool fallback = false)
        {
            var v = Get(section, key).Trim().ToLowerInvariant();
            if (v == "yes" || v == "true" || v == "1" || v == "on") return true;
            if (v == "no" || v == "false" || v == "0" || v == "off") return false;
            return fallback;
        }

        public int GetInt(string section, string key, int fallback = 0)
        {
            return int.TryParse(Get(section, key), out var n) ? n : fallback;
        }

        public string Origin(string section, string key)
        {
            return origin.TryGetValue(section + "/" + key, out var o) ? o : "";
        }

        public IEnumerable<string> Sections { get { return data.Keys; } }
        public IEnumerable<string> Keys(string section) { return data.TryGetValue(section, out var sec) ? sec.Keys : new string[0]; }

        static readonly Regex Var = new Regex(@"%([A-Za-z0-9_]+)%");

        string Expand(string section, string value, int depth)
        {
            if (depth > 8) return value;
            return Var.Replace(value, m =>
            {
                var name = m.Groups[1].Value;
                var v = Raw(section, name);
                if (v.Length == 0 && !section.Equals("Paths", StringComparison.OrdinalIgnoreCase)) v = Raw("Paths", name);
                if (v.Length > 0) return Expand(section, v, depth + 1);
                var env = Environment.GetEnvironmentVariable(name);
                return env ?? m.Value;
            });
        }

        // ------------------------------------------------------------------ typed shortcuts
        public string MacroFile { get { return Get("Macro", "File"); } }
        public string MacroModule { get { return Get("Macro", "Module", "mod_PolyAppEntry"); } }
        public string MacroProcedure { get { return Get("Macro", "Procedure", "PA_Run"); } }
        public bool SelectTree { get { return GetBool("Options", "SelectTree"); } }
        public bool LogCommands { get { return GetBool("Debug", "LogCommands"); } }
        public int InterceptMakeAssembly { get { return GetInt("Intercept", "MakeAssembly"); } }

        /// <summary>PDM mode: "auto", "yes" or "no".</summary>
        public string PdmMode { get { return Get("PDM", "Enabled", "auto").Trim().ToLowerInvariant(); } }

        /// <summary>Per-user override file, created on demand by the add-in's Settings command.</summary>
        public static string UserFile
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PolyApp", FileName); }
        }

        // The compiled-in floor: enough to run without any ini file at all.
        const string Defaults = @"
[Paths]
VaultRoot=C:\_PolyVault
SharedDocs=%VaultRoot%\SOLIDWORKS SHARED DOCUMENTS
Projects=%VaultRoot%\PROJECTS\
Server=\\polyfleng01
[Macro]
File=%SharedDocs%\Macros\PolyAppPro.swp
Module=mod_PolyAppEntry
Procedure=PA_Run
[PDM]
Enabled=auto
Vault=_PolyVault
[Options]
SelectTree=no
[Intercept]
MakeAssembly=0
[Debug]
LogCommands=no
";
    }
}
