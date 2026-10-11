using System;
using System.IO;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace PolyApp
{
    /// <summary>
    /// Runs a command through the PolyAppPro macro. RunMacro2 takes no arguments, so the request goes
    /// through %TEMP%\PolyApp.request.ini ([Request] Command, Option, Option2, Tree, Doc) and the macro's
    /// PA_Run reads it back with GetPrivateProfileString (see macro\mod_PolyAppEntry.bas).
    /// </summary>
    public static class MacroRunner
    {
        public static string RequestFile { get { return Path.Combine(Path.GetTempPath(), "PolyApp.request.ini"); } }

        /// <summary>The frm_Main control(s) for a command: [MacroOptions] in the ini overrides the catalogue.</summary>
        public static string ResolveOption(Settings s, Command c, out string option2)
        {
            option2 = c.MacroOption2;
            var over = s.Get("MacroOptions", c.Id);
            if (over.Length == 0) return c.MacroOption;
            var parts = over.Split(',');
            option2 = parts.Length > 1 ? parts[1].Trim() : "";
            return parts[0].Trim();
        }

        public static CommandResult Run(ISldWorks sw, Settings s, Command c, DocContext ctx, bool tree)
        {
            var opt = ResolveOption(s, c, out var opt2);
            if (opt.Length == 0)
                return CommandResult.Fail(c.Title + " is not mapped to the macro yet. Set [MacroOptions] " + c.Id + "=opt_... in " + Settings.FileName + ".");
            var file = s.MacroFile;
            if (!File.Exists(file))
                return CommandResult.Fail("PolyAppPro macro not found at " + file + ". Set [Macro] File in " + Settings.FileName + ".");

            var sb = new StringBuilder();
            sb.AppendLine("[Request]");
            sb.AppendLine("Command=" + c.Id);
            sb.AppendLine("Option=" + opt);
            sb.AppendLine("Option2=" + opt2);
            sb.AppendLine("Tree=" + (tree ? "1" : "0"));
            sb.AppendLine("Doc=" + ctx.Path);
            sb.AppendLine("Title=" + ctx.Title);
            sb.AppendLine("Time=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            try { File.WriteAllText(RequestFile, sb.ToString()); }
            catch (Exception e) { return CommandResult.Fail("could not write " + RequestFile + ": " + e.Message); }

            Log.Write("macro " + c.Id + " -> " + opt + (opt2.Length > 0 ? " + " + opt2 : "") + " tree=" + tree + " on " + ctx);
            int err = 0;
            bool ok;
            try { ok = sw.RunMacro2(file, s.MacroModule, s.MacroProcedure, (int)swRunMacroOption_e.swRunMacroUnloadAfterRun, out err); }
            catch (Exception e) { return CommandResult.Fail("RunMacro2 threw: " + e.Message); }
            if (!ok) return CommandResult.Fail("RunMacro2 failed (" + Describe(err) + ") for " + s.MacroModule + "." + s.MacroProcedure + " in " + file);
            return CommandResult.Done(c.Title + " ran through the macro.");
        }

        static string Describe(int err)
        {
            try { return ((swRunMacroError_e)err).ToString(); } catch { return "error " + err; }
        }
    }
}
