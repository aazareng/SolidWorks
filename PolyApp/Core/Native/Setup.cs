using System;
using System.Diagnostics;
using System.IO;

namespace PolyApp.Native
{
    /// <summary>The Setup box: open the per-user ini (created from the shipped one on first use) and the log.</summary>
    public static class Setup
    {
        /// <summary>Folder holding the shipped PolyApp.ini; the add-in and the CLI set it at startup.</summary>
        public static string BaseDir = AppDomain.CurrentDomain.BaseDirectory;

        public static CommandResult? OpenSettings(Dispatcher d, Command c, DocContext ctx)
        {
            var user = Settings.UserFile;
            try
            {
                if (!File.Exists(user))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(user)!);
                    var shipped = Path.Combine(BaseDir, Settings.FileName);
                    var text = File.Exists(shipped) ? File.ReadAllText(shipped) : "[Paths]\r\n";
                    File.WriteAllText(user, "; Per-user overrides. Only the keys you change need to stay; the shipped file beside the add-in supplies the rest.\r\n" + text);
                }
                Process.Start("notepad.exe", "\"" + user + "\"");
                return CommandResult.Done("opened " + user + " (restart SOLIDWORKS, or run `polyapp ini`, to see the result)");
            }
            catch (Exception e) { return CommandResult.Fail("could not open " + user + ": " + e.Message); }
        }

        public static CommandResult? OpenLog(Dispatcher d, Command c, DocContext ctx)
        {
            try
            {
                if (!File.Exists(Log.Path_)) Log.Write("log opened");
                Process.Start("notepad.exe", "\"" + Log.Path_ + "\"");
                return CommandResult.Done("opened " + Log.Path_);
            }
            catch (Exception e) { return CommandResult.Fail(e.Message); }
        }
    }
}
