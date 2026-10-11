// ============================================================================
//  polyapp — PolyApp from the terminal, against the SOLIDWORKS running as you.
//
//    polyapp attach                     SOLIDWORKS version, pid, active document, PDM status
//    polyapp ini                        the resolved settings and which file each value came from
//    polyapp commands                   the catalogue: id, title, box/flyout, documents, native or macro mapping
//    polyapp run <Id> [--tree]          run one command exactly as the add-in button would
//    polyapp load-addin [dll]           ask SOLIDWORKS to load the registered add-in now (no restart)
//    polyapp unload-addin [dll]
//
//  --start on attach/run launches SOLIDWORKS when none is running; --trace prints stack traces.
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PolyApp;
using PolyApp.Native;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        var rest = new List<string>();
        var opts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var a in args)
        {
            if (a.StartsWith("--")) { int eq = a.IndexOf('='); if (eq > 0) opts[a.Substring(2, eq - 2)] = a.Substring(eq + 1); else opts[a.Substring(2)] = "1"; }
            else rest.Add(a);
        }
        var cmd = rest.Count > 0 ? rest[0].ToLowerInvariant() : "help";
        Log.Echo = s => Console.WriteLine("  log: " + s);
        Setup.BaseDir = AppContext.BaseDirectory;
        var settings = Settings.Load(AppContext.BaseDirectory);

        try
        {
            switch (cmd)
            {
                case "attach":
                {
                    var sw = SwSession.Connect(opts.ContainsKey("start"));
                    Console.WriteLine(SwSession.Describe(sw));
                    PdmSession.Init(settings);
                    Console.WriteLine("PDM: " + PdmSession.Status);
                    Console.WriteLine("Macro: " + settings.MacroFile + (File.Exists(settings.MacroFile) ? "" : "  (NOT FOUND)"));
                    return 0;
                }
                case "ini":
                {
                    Console.WriteLine("files: " + (settings.Files.Count == 0 ? "(none: compiled defaults only)" : string.Join(" ; ", settings.Files)));
                    Console.WriteLine("user overrides go in: " + Settings.UserFile + (File.Exists(Settings.UserFile) ? "" : "  (not present)"));
                    foreach (var sec in settings.Sections)
                    {
                        Console.WriteLine();
                        Console.WriteLine("[" + sec + "]");
                        foreach (var key in settings.Keys(sec))
                        {
                            var origin = settings.Origin(sec, key);
                            var tag = origin == "(defaults)" ? "default" : Path.GetFileName(Path.GetDirectoryName(origin) ?? "") + "\\" + Path.GetFileName(origin);
                            Console.WriteLine("  " + key + " = " + settings.Get(sec, key) + "    [" + tag + "]");
                        }
                    }
                    return 0;
                }
                case "commands":
                {
                    Console.WriteLine($"{"Id",-14} {"Title",-18} {"Box / Flyout",-26} {"Docs",-9} {"Runs as",-8} Macro control(s)");
                    foreach (var c in Catalog.All)
                    {
                        var opt = MacroRunner.ResolveOption(settings, c, out var opt2);
                        var where = c.Box + (c.Flyout != null ? " / " + c.Flyout : "");
                        var how = Dispatcher.HasBuiltIn(c.Id) ? "native" : opt.Length > 0 ? "macro" : "UNMAPPED";
                        var docs = c.Docs == Docs.None ? "always" : c.Docs.ToString();
                        Console.WriteLine($"{c.Id,-14} {c.Title,-18} {where,-26} {docs,-9} {how,-8} {opt}{(opt2.Length > 0 ? " + " + opt2 : "")}");
                    }
                    Console.WriteLine();
                    Console.WriteLine(Catalog.All.Length + " commands; macro = " + settings.MacroFile);
                    return 0;
                }
                case "run":
                {
                    if (rest.Count < 2) throw new ArgumentException("run <CommandId> [--tree]");
                    var sw = SwSession.Connect(opts.ContainsKey("start"));
                    PdmSession.Init(settings);
                    var d = new Dispatcher(sw, settings);
                    if (opts.ContainsKey("tree")) d.SelectTree = true;
                    var r = d.Run(rest[1]);
                    Console.WriteLine((r.Ok ? "OK: " : "FAILED: ") + r.Message);
                    return r.Ok ? 0 : 1;
                }
                case "load-addin":
                {
                    var sw = SwSession.Connect(false);
                    var dll = Path.GetFullPath(rest.Count > 1 ? rest[1] : DefaultDll());
                    int r = sw.LoadAddIn(dll);
                    Console.WriteLine(r == 0 ? "LoadAddIn ok: " + dll : $"LoadAddIn returned {r} for {dll} (0 = ok, 1 = failed/not registered, 2 = already loaded)");
                    object? obj = null;
                    try { obj = sw.GetAddInObject("PolyApp.AddIn"); } catch (Exception e) { Console.WriteLine("GetAddInObject threw " + e.Message); }
                    Console.WriteLine(obj != null ? "add-in object is live (ConnectToSW ran)" : "add-in object NOT found: SOLIDWORKS did not instantiate it (registration or load failure; see " + Log.Path_ + ")");
                    return r == 0 && obj != null ? 0 : 1;
                }
                case "unload-addin":
                {
                    var sw = SwSession.Connect(false);
                    var dll = Path.GetFullPath(rest.Count > 1 ? rest[1] : DefaultDll());
                    int r = sw.UnloadAddIn(dll);
                    Console.WriteLine($"UnloadAddIn returned {r}");
                    return r == 0 ? 0 : 1;
                }
                default:
                    Console.WriteLine("polyapp attach | ini | commands | run <Id> [--tree] | load-addin [dll] | unload-addin [dll]   (--start, --trace)");
                    Console.WriteLine("log: " + Log.Path_);
                    return cmd == "help" ? 0 : 2;
            }
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("ERROR: " + e.Message);
            if (opts.ContainsKey("trace")) Console.Error.WriteLine(e);
            return 1;
        }
    }

    static string DefaultDll()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PolyApp", "AddIn", "PolyApp.AddIn.dll");
    }
}
