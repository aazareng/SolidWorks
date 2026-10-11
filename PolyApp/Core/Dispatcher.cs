using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;

namespace PolyApp
{
    public sealed class CommandResult
    {
        public bool Ok;
        public string Message = "";
        public static CommandResult Done(string message) { return new CommandResult { Ok = true, Message = message }; }
        public static CommandResult Fail(string message) { return new CommandResult { Ok = false, Message = message }; }
    }

    /// <summary>
    /// Runs a command by id: a native (C#) handler first, the macro otherwise. A native handler may return
    /// null to say "not my case" and fall through to the macro, which is how a command is ported a step at a time.
    /// </summary>
    public sealed class Dispatcher
    {
        public delegate CommandResult? NativeHandler(Dispatcher d, Command c, DocContext ctx);

        public readonly ISldWorks Sw;
        public readonly Settings Settings;
        /// <summary>Power Select walks the whole tree (true) or the top level only; toggled by the Tree button.</summary>
        public bool SelectTree;

        readonly Dictionary<string, NativeHandler> native;

        /// <summary>The native handlers every Dispatcher starts with. Port a command: add a line here.</summary>
        static readonly Dictionary<string, NativeHandler> BuiltIn = new Dictionary<string, NativeHandler>(StringComparer.OrdinalIgnoreCase)
        {
            { "OpenDrawing", Native.OpenDrawing.Run },
            { "Settings", Native.Setup.OpenSettings },
            { "ShowLog", Native.Setup.OpenLog },
            { "SelectTree", (d, c, ctx) =>
                {
                    d.SelectTree = !d.SelectTree;
                    return CommandResult.Done("Power Select: " + (d.SelectTree ? "whole tree" : "top level only"));
                } },
        };

        public static bool HasBuiltIn(string id) { return BuiltIn.ContainsKey(id); }

        public Dispatcher(ISldWorks sw, Settings settings)
        {
            Sw = sw;
            Settings = settings;
            SelectTree = settings.SelectTree;
            native = new Dictionary<string, NativeHandler>(BuiltIn, StringComparer.OrdinalIgnoreCase);
        }

        public void RegisterNative(string id, NativeHandler handler) { native[id] = handler; }
        public bool IsNative(string id) { return native.ContainsKey(id); }

        public CommandResult Run(string id)
        {
            var c = Catalog.Find(id);
            return c == null ? CommandResult.Fail("unknown command " + id) : Run(c);
        }

        public CommandResult Run(Command c)
        {
            var ctx = DocContext.Capture(Sw);
            try
            {
                if (native.TryGetValue(c.Id, out var h))
                {
                    var r = h(this, c, ctx);
                    if (r != null) { Log.Write("native " + c.Id + (r.Ok ? " ok: " : " FAILED: ") + r.Message); return r; }
                    Log.Write("native " + c.Id + " passed to the macro");
                }
                if (c.Docs != Docs.None && !ctx.HasDoc) return CommandResult.Fail(c.Title + " needs an open document.");
                if (c.Docs != Docs.None && (c.Docs & ctx.Kind) == 0) return CommandResult.Fail(c.Title + " does not apply to a " + ctx.Kind.ToString().ToLowerInvariant() + ".");
                var m = MacroRunner.Run(Sw, Settings, c, ctx, SelectTree);
                if (!m.Ok) Log.Write("macro " + c.Id + " FAILED: " + m.Message);
                return m;
            }
            catch (Exception e)
            {
                Log.Write(c.Id + " threw: " + e);
                return CommandResult.Fail(c.Title + ": " + e.Message);
            }
        }
    }
}
