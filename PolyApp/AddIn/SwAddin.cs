using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace PolyApp
{
    /// <summary>
    /// The PolyApp add-in. SOLIDWORKS loads it at startup and calls ConnectToSW; from there the catalogue
    /// (<see cref="Catalog"/>) becomes a "PolyApp" CommandManager tab in parts, assemblies and drawings
    /// (boxes PDM / Make / Select / Export / Drawing / More / Setup, flyouts for the families), a Tools
    /// menu, and entries on the right-click menu of a component. Every button goes through
    /// <see cref="Dispatcher"/>: a native handler when one exists, the PolyAppPro macro otherwise.
    /// </summary>
    // No ClassInterface(None): SOLIDWORKS reaches the callbacks (Cmd_*, Enable_*) by name through IDispatch,
    // and SetAddinCallbackInfo2 throws InvalidCastException without a dispatch interface.
    [ComVisible(true)]
    [Guid(AddInGuid)]
    [ProgId("PolyApp.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class SwAddin : ISwAddin
    {
        public const string AddInGuid = "5C1D7A2E-9B64-4F3A-8E21-7D0A3B9C4E55";
        public const string Title = "PolyApp";
        public const string Description = "PolyApp engineering commands: PDM state, Branch, Open, Make Drawing, Hotsheet, Shipping Order, Cutsheet, Power Select, exports.";

        const int GroupId = 7501;
        const int FlyoutBaseId = 7510;

        ISldWorks? sw;
        SldWorks? events;
        ICommandManager? cmdMgr;
        ICommandGroup? group;
        int cookie;
        Settings? settings;
        Dispatcher? dispatcher;
        readonly Dictionary<string, int> buttonIndex = new Dictionary<string, int>();        // command id -> index in the group
        readonly Dictionary<string, IFlyoutGroup> flyouts = new Dictionary<string, IFlyoutGroup>();
        readonly List<Command> popupItems = new List<Command>();
        bool reentering;

        public static string BaseDir
        {
            get { return Path.GetDirectoryName(typeof(SwAddin).Assembly.Location) ?? System.Environment.CurrentDirectory; }
        }

        // SOLIDWORKS probes its own folder for assemblies, not ours: resolve anything we ship from next to this DLL.
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
            Log.Write("ConnectToSW: cookie " + cookie + ", add-in " + typeof(SwAddin).Assembly.GetName().Version + " from " + BaseDir);
            try
            {
                sw = (ISldWorks)thisSw;
                this.cookie = cookie;
                sw.SetAddinCallbackInfo2(0, this, cookie);
                cmdMgr = sw.GetCommandManager(cookie);
                Log.Write("SOLIDWORKS " + sw.RevisionNumber());
                Native.Setup.BaseDir = BaseDir;
                settings = Settings.Load(BaseDir);
                Log.Write("settings from " + (settings.Files.Count == 0 ? "(defaults only)" : string.Join(" ; ", settings.Files)));
                dispatcher = new Dispatcher(sw, settings);
                PdmSession.Init(settings);
            }
            catch (Exception e) { Log.Write("ConnectToSW failed: " + e); return false; }

            CheckCallbacks();
            try { AddCommands(); Log.Write("commands added: " + buttonIndex.Count + " buttons, " + flyouts.Count + " flyouts"); }
            catch (Exception e) { Log.Write("AddCommands: " + e); sw.SendMsgToUser2(Title + ": could not create the toolbar: " + e.Message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk); }
            try { AddContextMenu(); }
            catch (Exception e) { Log.Write("context menu: " + e.Message); }
            try { HookEvents(); }
            catch (Exception e) { Log.Write("events: " + e.Message); }
            return true;
        }

        public bool DisconnectFromSW()
        {
            try { if (events != null) events.CommandOpenPreNotify -= OnCommandOpen; } catch { }
            events = null;
            try
            {
                foreach (var c in popupItems)
                    sw!.RemoveMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, cookie, (int)swSelectType_e.swSelCOMPONENTS, c.Title, "Cmd_" + c.Id, EnableFor(c), c.Hint, "");
            }
            catch { }
            popupItems.Clear();
            try { foreach (var f in flyouts.Values) cmdMgr!.RemoveFlyoutGroup(f.CmdID); } catch { }
            flyouts.Clear();
            try { if (cmdMgr != null) cmdMgr.RemoveCommandGroup2(GroupId, false); } catch { }
            group = null;
            cmdMgr = null;
            dispatcher = null;
            sw = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        // ------------------------------------------------------------------ commands
        /// <summary>Every catalogue command needs a Cmd_&lt;Id&gt; method here; log the ones that are missing so a new entry cannot silently do nothing.</summary>
        void CheckCallbacks()
        {
            var missing = Catalog.All.Where(c => GetType().GetMethod("Cmd_" + c.Id) == null).Select(c => c.Id).ToList();
            if (missing.Count > 0) Log.Write("NO CALLBACK for: " + string.Join(", ", missing) + " (add Cmd_<Id> to SwAddin.cs)");
        }

        static string EnableFor(Command c)
        {
            return c.Id == "SelectTree" ? "Enable_SelectTree" : Catalog.EnableMethod(c.Docs);
        }

        void AddCommands()
        {
            int err = 0;
            group = cmdMgr!.CreateCommandGroup2(GroupId, Title, "PolyApp commands", "", -1, true, ref err);
            if (group == null) throw new InvalidOperationException("CreateCommandGroup2 failed (" + err + ")");

            // plain buttons, one command group
            var plain = Catalog.All.Where(c => c.Flyout == null).ToList();
            try
            {
                group.IconList = IconFactory.Strip("group", plain);
                group.MainIconList = IconFactory.Main("main", "PA", "PDM");
            }
            catch (Exception e) { Log.Write("icons: " + e.Message); }
            for (int i = 0; i < plain.Count; i++)
            {
                var c = plain[i];
                int idx = group.AddCommandItem2(c.Title, -1, c.Hint, c.Title, i, "Cmd_" + c.Id, EnableFor(c), i,
                    (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
                buttonIndex[c.Id] = idx;
            }
            group.HasToolbar = true;
            group.HasMenu = true;
            group.Activate();

            // flyouts: one per family (PDM State, Make Assembly, the three Power Select groups, More)
            int fid = FlyoutBaseId;
            foreach (var box in Catalog.Boxes)
                foreach (var title in Catalog.FlyoutsIn(box))
                {
                    var items = Catalog.InFlyout(title).ToList();
                    var fg = cmdMgr.CreateFlyoutGroup2(fid++, title, title, title,
                        IconFactory.Main("flyout_" + title, Abbrev(title), box), IconFactory.Strip("flyout_" + title, items),
                        nameof(Flyout_Main), Catalog.EnableMethod(Catalog.DocsOf(items)));
                    if (fg == null) { Log.Write("flyout " + title + " not created"); continue; }
                    for (int j = 0; j < items.Count; j++)
                        fg.AddCommandItem(items[j].Title, items[j].Hint, j, "Cmd_" + items[j].Id, EnableFor(items[j]));
                    fg.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;
                    flyouts[title] = fg;
                }

            // a PolyApp tab per document type; rebuilt each start so catalogue changes show up
            foreach (var kv in new[] { Tuple.Create(Docs.Part, swDocumentTypes_e.swDocPART), Tuple.Create(Docs.Assembly, swDocumentTypes_e.swDocASSEMBLY), Tuple.Create(Docs.Drawing, swDocumentTypes_e.swDocDRAWING) })
            {
                try
                {
                    var kind = kv.Item1;
                    int docType = (int)kv.Item2;
                    var old = cmdMgr.GetCommandTab(docType, Title);
                    if (old != null) cmdMgr.RemoveCommandTab(old);
                    var tab = cmdMgr.AddCommandTab(docType, Title);
                    if (tab == null) { Log.Write("no tab for doc type " + docType); continue; }
                    foreach (var box in Catalog.Boxes)
                    {
                        var ids = new List<int>();
                        var styles = new List<int>();
                        foreach (var c in Catalog.InBox(box).Where(c => c.Flyout == null && (c.Docs == Docs.None || (c.Docs & kind) != 0)))
                        {
                            ids.Add(group.get_CommandID(buttonIndex[c.Id]));
                            styles.Add((int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow);
                        }
                        foreach (var title in Catalog.FlyoutsIn(box))
                        {
                            if ((Catalog.DocsOf(Catalog.InFlyout(title)) & kind) == 0 || !flyouts.ContainsKey(title)) continue;
                            ids.Add(flyouts[title].CmdID);
                            styles.Add((int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_SimpleFlyout);
                        }
                        if (ids.Count == 0) continue;
                        var tabBox = tab.AddCommandTabBox();
                        tabBox.AddCommands(ids.ToArray(), styles.ToArray());
                    }
                }
                catch (Exception e) { Log.Write("command tab " + kv.Item1 + ": " + e.Message); }
            }
        }

        static string Abbrev(string title)
        {
            var words = title.Split(' ').Where(w => w.Length > 0).ToArray();
            return words.Length >= 2 ? string.Concat(words.Take(3).Select(w => char.ToUpperInvariant(w[0]))) : title.Substring(0, Math.Min(3, title.Length)).ToUpperInvariant();
        }

        /// <summary>Right-click on a component in an assembly: the commands flagged ContextMenu.</summary>
        void AddContextMenu()
        {
            foreach (var c in Catalog.All.Where(c => c.ContextMenu))
            {
                int id = sw!.AddMenuPopupItem3((int)swDocumentTypes_e.swDocASSEMBLY, cookie, (int)swSelectType_e.swSelCOMPONENTS, c.Title, "Cmd_" + c.Id, EnableFor(c), c.Hint, "");
                if (id > 0) popupItems.Add(c); else Log.Write("popup item " + c.Id + " returned " + id);
            }
        }

        // ------------------------------------------------------------------ events
        void HookEvents()
        {
            if (!settings!.LogCommands && settings.InterceptMakeAssembly == 0) return;
            events = (SldWorks)sw!;
            events.CommandOpenPreNotify += OnCommandOpen;
            Log.Write("command events hooked (LogCommands=" + settings.LogCommands + ", Intercept MakeAssembly=" + settings.InterceptMakeAssembly + ")");
        }

        int OnCommandOpen(int command, int userCommand)
        {
            try
            {
                if (settings!.LogCommands) Log.Write("command " + command + " user " + userCommand);
                if (!reentering && command == settings.InterceptMakeAssembly && settings.InterceptMakeAssembly != 0)
                {
                    var ctx = DocContext.Capture(sw!);
                    if ((ctx.Kind & Docs.Model) != 0) { ShowMakeMenu(command); return 1; }
                }
            }
            catch (Exception e) { Log.Write("OnCommandOpen: " + e.Message); }
            return 0;
        }

        /// <summary>In place of Make Assembly: the PolyApp "Make Assembly" family, then the plain SOLIDWORKS command.</summary>
        void ShowMakeMenu(int swCommand)
        {
            var menu = new ContextMenuStrip();
            foreach (var c in Catalog.InFlyout("Make Assembly"))
            {
                var cmd = c;
                menu.Items.Add(cmd.Title, null, (s, e) => Run(cmd));
            }
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Make Assembly (SOLIDWORKS)", null, (s, e) =>
            {
                reentering = true;
                try { sw!.RunCommand(swCommand, ""); }
                catch (Exception ex) { Log.Write("RunCommand " + swCommand + ": " + ex.Message); }
                finally { reentering = false; }
            });
            menu.Show(Cursor.Position);
        }

        // ------------------------------------------------------------------ run
        void Run(string id)
        {
            var c = Catalog.Find(id);
            if (c == null) { Log.Write("unknown command " + id); return; }
            Run(c);
        }

        void Run(Command c)
        {
            if (dispatcher == null || sw == null) return;
            var r = dispatcher.Run(c);
            try { ((IFrame)sw.Frame()).SetStatusBarText(Title + ": " + r.Message); } catch { }
            if (!r.Ok) sw.SendMsgToUser2(Title + " - " + c.Title + "\n\n" + r.Message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
        }

        // ------------------------------------------------------------------ callbacks (by name, through IDispatch)
        public void Flyout_Main() { }

        public int Enable_Always() { return 1; }
        public int Enable_Any() { return Kind() != Docs.None ? 1 : 0; }
        public int Enable_Model() { return (Kind() & Docs.Model) != 0 ? 1 : 0; }
        public int Enable_Part() { return Kind() == Docs.Part ? 1 : 0; }
        public int Enable_Assembly() { return Kind() == Docs.Assembly ? 1 : 0; }
        public int Enable_Drawing() { return Kind() == Docs.Drawing ? 1 : 0; }
        /// <summary>The Tree toggle shows pressed (3) while Power Select walks the whole tree.</summary>
        public int Enable_SelectTree() { return Kind() != Docs.Assembly ? 0 : (dispatcher != null && dispatcher.SelectTree ? 3 : 1); }

        Docs Kind()
        {
            try { var doc = sw?.ActiveDoc as IModelDoc2; return doc == null ? Docs.None : DocContext.KindOf(doc.GetType()); }
            catch { return Docs.None; }
        }

        // PDM
        public void Cmd_PdmCheckOut() { Run("PdmCheckOut"); }
        public void Cmd_PdmCheckIn() { Run("PdmCheckIn"); }
        public void Cmd_PdmIssue() { Run("PdmIssue"); }
        public void Cmd_PdmIssueNext() { Run("PdmIssueNext"); }
        public void Cmd_PdmBanish() { Run("PdmBanish"); }
        public void Cmd_PdmObsolete() { Run("PdmObsolete"); }
        public void Cmd_Branch() { Run("Branch"); }
        public void Cmd_Open() { Run("Open"); }
        public void Cmd_OpenDrawing() { Run("OpenDrawing"); }
        // Make
        public void Cmd_MakeDrawing() { Run("MakeDrawing"); }
        public void Cmd_Hotsheet() { Run("Hotsheet"); }
        public void Cmd_ShippingOrder() { Run("ShippingOrder"); }
        public void Cmd_Shipping() { Run("Shipping"); }
        public void Cmd_Por() { Run("Por"); }
        public void Cmd_Manual() { Run("Manual"); }
        public void Cmd_Spec() { Run("Spec"); }
        public void Cmd_CutsheetAssy() { Run("CutsheetAssy"); }
        public void Cmd_Modified() { Run("Modified"); }
        public void Cmd_Cutsheet() { Run("Cutsheet"); }
        public void Cmd_MateOrigins() { Run("MateOrigins"); }
        public void Cmd_Mirrorize() { Run("Mirrorize"); }
        // Select
        public void Cmd_SelSheetMetal() { Run("SelSheetMetal"); }
        public void Cmd_SelAssemblies() { Run("SelAssemblies"); }
        public void Cmd_SelParts() { Run("SelParts"); }
        public void Cmd_SelVirtual() { Run("SelVirtual"); }
        public void Cmd_SelExcluded() { Run("SelExcluded"); }
        public void Cmd_SelEnvelope() { Run("SelEnvelope"); }
        public void Cmd_SelFixed() { Run("SelFixed"); }
        public void Cmd_SelExtRef() { Run("SelExtRef"); }
        public void Cmd_SelPre13() { Run("SelPre13"); }
        public void Cmd_SelUnderDef() { Run("SelUnderDef"); }
        public void Cmd_SelVisible() { Run("SelVisible"); }
        public void Cmd_SelHidden() { Run("SelHidden"); }
        public void Cmd_SelRouter() { Run("SelRouter"); }
        public void Cmd_SelMill() { Run("SelMill"); }
        public void Cmd_SelLathe() { Run("SelLathe"); }
        public void Cmd_Sel3D() { Run("Sel3D"); }
        public void Cmd_SelShear() { Run("SelShear"); }
        public void Cmd_SelManual() { Run("SelManual"); }
        public void Cmd_SelPurchased() { Run("SelPurchased"); }
        public void Cmd_SelStandard() { Run("SelStandard"); }
        public void Cmd_SelSpec() { Run("SelSpec"); }
        public void Cmd_SelSerial() { Run("SelSerial"); }
        public void Cmd_SelIssued() { Run("SelIssued"); }
        public void Cmd_SelRight() { Run("SelRight"); }
        public void Cmd_SelConfig() { Run("SelConfig"); }
        public void Cmd_SelRedFace() { Run("SelRedFace"); }
        public void Cmd_SelFlat() { Run("SelFlat"); }
        public void Cmd_SelectTree() { Run("SelectTree"); }
        // Export
        public void Cmd_Dxf() { Run("Dxf"); }
        public void Cmd_Pdf() { Run("Pdf"); }
        public void Cmd_EDrawings() { Run("EDrawings"); }
        public void Cmd_Step() { Run("Step"); }
        public void Cmd_ExportBom() { Run("ExportBom"); }
        public void Cmd_ExportHtml() { Run("ExportHtml"); }
        // Drawing
        public void Cmd_ReloadFormat() { Run("ReloadFormat"); }
        public void Cmd_CheckBalloons() { Run("CheckBalloons"); }
        // More
        public void Cmd_AddConfig() { Run("AddConfig"); }
        public void Cmd_SyncConfig() { Run("SyncConfig"); }
        public void Cmd_Configger() { Run("Configger"); }
        public void Cmd_Colorize() { Run("Colorize"); }
        public void Cmd_Pinkify() { Run("Pinkify"); }
        public void Cmd_PurgeConfigs() { Run("PurgeConfigs"); }
        public void Cmd_TrimProps() { Run("TrimProps"); }
        public void Cmd_TrimConfigs() { Run("TrimConfigs"); }
        public void Cmd_CutSize() { Run("CutSize"); }
        public void Cmd_SetMaterial() { Run("SetMaterial"); }
        public void Cmd_CsvAssembly() { Run("CsvAssembly"); }
        public void Cmd_ErrorCheck() { Run("ErrorCheck"); }
        public void Cmd_ReplaceRef() { Run("ReplaceRef"); }
        public void Cmd_Dissolve() { Run("Dissolve"); }
        public void Cmd_ExcludeBom() { Run("ExcludeBom"); }
        public void Cmd_AddSpec() { Run("AddSpec"); }
        public void Cmd_SyncCard() { Run("SyncCard"); }
        public void Cmd_DeleteDoc() { Run("DeleteDoc"); }
        // Setup
        public void Cmd_Settings() { Run("Settings"); }
        public void Cmd_ShowLog() { Run("ShowLog"); }

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
    }
}
