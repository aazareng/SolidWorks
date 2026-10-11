using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace StructuralSheetMetal
{
    /// <summary>
    /// The SOLIDWORKS add-in: one command, "Structural Sheet Metal Member", in its own CommandManager
    /// tab (part documents) and under Tools. The command opens <see cref="MemberPage"/>.
    /// </summary>
    // No ClassInterface(None) here: SOLIDWORKS reaches the callbacks (ShowMemberPage, EnableMemberPage) by
    // name through IDispatch, and SetAddinCallbackInfo2 throws InvalidCastException without a dispatch interface.
    [ComVisible(true)]
    [Guid(AddInGuid)]
    [ProgId("StructuralSheetMetal.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class SwAddin : ISwAddin
    {
        public const string AddInGuid = "8F2B6C3A-5D41-4E7B-9A10-2C6F4B1E7D30";
        public const string Title = "Structural Sheet Metal";
        public const string Description = "Sheet-metal members from OPEN profiles (L, C, Z, hat...) along a sketch path, with corner treatments. Structural Member for sheet metal.";

        const int GroupId = 7421;
        const int CmdMember = 0;

        ISldWorks? sw;
        ICommandManager? cmdMgr;
        ICommandGroup? group;
        int cookie;
        MemberPage? page;

        public static string BaseDir
        {
            get { return Path.GetDirectoryName(typeof(SwAddin).Assembly.Location) ?? System.Environment.CurrentDirectory; }
        }

        // SOLIDWORKS probes its own folder for assemblies, not ours: resolve System.Text.Json and friends
        // from next to this DLL (any requested version: the files shipped with the build are the ones to use).
        static SwAddin()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            {
                try
                {
                    var name = new System.Reflection.AssemblyName(e.Name).Name;
                    var path = Path.Combine(BaseDir, name + ".dll");
                    return File.Exists(path) ? System.Reflection.Assembly.LoadFrom(path) : null;
                }
                catch { return null; }
            };
        }

        // ------------------------------------------------------------------ ISwAddin
        public bool ConnectToSW(object thisSw, int cookie)
        {
            Log("ConnectToSW: cookie " + cookie + ", add-in " + typeof(SwAddin).Assembly.GetName().Version + " from " + BaseDir);
            try
            {
                sw = (ISldWorks)thisSw;
                this.cookie = cookie;
                sw.SetAddinCallbackInfo2(0, this, cookie);
                cmdMgr = sw.GetCommandManager(cookie);
                Log("SOLIDWORKS " + sw.RevisionNumber());
            }
            catch (Exception e) { Log("ConnectToSW failed: " + e); return false; }
            try { AddCommands(); Log("commands added"); }
            catch (Exception e) { Log("AddCommands: " + e); sw.SendMsgToUser2(Title + ": could not create the toolbar: " + e.Message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk); }
            return true;
        }

        public bool DisconnectFromSW()
        {
            try { if (page != null) page.Close(); } catch { }
            page = null;
            try { if (cmdMgr != null) cmdMgr.RemoveCommandGroup2(GroupId, false); } catch { }
            group = null;
            cmdMgr = null;
            sw = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        // ------------------------------------------------------------------ commands
        void AddCommands()
        {
            int err = 0;
            group = cmdMgr!.CreateCommandGroup2(GroupId, Title, "Sheet metal members along a sketch path", "", -1, true, ref err);
            if (group == null) throw new InvalidOperationException("CreateCommandGroup2 failed (" + err + ")");

            var sizes = new[] { 20, 32, 40, 64, 96, 128 };
            var iconDir = Path.Combine(BaseDir, "Icons");
            var cmdIcons = sizes.Select(s => Path.Combine(iconDir, "member_" + s + ".png")).ToArray();
            var mainIcons = sizes.Select(s => Path.Combine(iconDir, "main_" + s + ".png")).ToArray();
            if (cmdIcons.All(File.Exists)) { group.IconList = cmdIcons; group.MainIconList = mainIcons.All(File.Exists) ? mainIcons : cmdIcons; }
            else Log("icons not found in " + iconDir);

            int idx = group.AddCommandItem2("Structural Sheet Metal Member", -1,
                "Sheet-metal members from open profiles along selected sketch lines or edges, with corner treatments",
                "Structural Sheet Metal Member", 0, nameof(ShowMemberPage), nameof(EnableMemberPage), CmdMember,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            group.HasToolbar = true;
            group.HasMenu = true;
            group.Activate();

            // a CommandManager tab in part documents
            try
            {
                int docType = (int)swDocumentTypes_e.swDocPART;
                var tab = cmdMgr.GetCommandTab(docType, Title) ?? cmdMgr.AddCommandTab(docType, Title);
                if (tab != null && tab.GetCommandTabBoxCount() == 0)
                {
                    var box = tab.AddCommandTabBox();
                    box.AddCommands(new[] { group.get_CommandID(idx) }, new[] { (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow });
                }
            }
            catch (Exception e) { Log("command tab: " + e.Message); }
        }

        /// <summary>Callback: opens the PropertyManager page.</summary>
        public void ShowMemberPage()
        {
            try
            {
                if (page != null) { try { page.Close(); } catch { } }
                page = new MemberPage(sw!);
                page.Show();
            }
            catch (Exception e)
            {
                Log("ShowMemberPage: " + e);
                sw!.SendMsgToUser2(Title + ": " + e.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        /// <summary>Callback: the command is enabled only in a part document.</summary>
        public int EnableMemberPage()
        {
            var doc = sw?.ActiveDoc as IModelDoc2;
            return doc != null && doc.GetType() == (int)swDocumentTypes_e.swDocPART ? 1 : 0;
        }

        // ------------------------------------------------------------------ registration
        // RegAsm calls these. Machine-wide keys need an elevated prompt; install\register.ps1 writes
        // the per-user equivalents (HKCU) when run without -Machine.
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            foreach (var root in new[] { Registry.LocalMachine, Registry.CurrentUser })
            {
                try
                {
                    using (var k = root.CreateSubKey(@"SOFTWARE\SolidWorks\AddIns\{" + AddInGuid + "}"))
                    {
                        k!.SetValue(null, 0, RegistryValueKind.DWord);
                        k.SetValue("Title", Title);
                        k.SetValue("Description", Description);
                    }
                    using (var k = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\SolidWorks\AddInsStartup\{" + AddInGuid + "}"))
                        k!.SetValue(null, 1, RegistryValueKind.DWord);
                    break;
                }
                catch (UnauthorizedAccessException) { /* not elevated: fall through to HKCU */ }
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            foreach (var root in new[] { Registry.LocalMachine, Registry.CurrentUser })
            {
                try { root.DeleteSubKeyTree(@"SOFTWARE\SolidWorks\AddIns\{" + AddInGuid + "}", false); } catch { }
                try { root.DeleteSubKeyTree(@"SOFTWARE\SolidWorks\AddInsStartup\{" + AddInGuid + "}", false); } catch { }
            }
        }

        // ------------------------------------------------------------------ log
        public static string LogPath { get { return Path.Combine(Path.GetTempPath(), "StructuralSheetMetal.log"); } }

        public static void Log(string text)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + text + System.Environment.NewLine); } catch { }
        }
    }
}
