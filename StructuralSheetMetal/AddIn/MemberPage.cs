using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace StructuralSheetMetal
{
    /// <summary>
    /// The PropertyManager page, laid out like Structural Member:
    ///
    ///   Profile      Standard / Type / Size, gauge table + gauge (or a typed thickness), bend radius
    ///   Groups       list of groups, New / Delete; each group has its own path selection and settings
    ///   Path         the selection box (sketch lines, straight edges) of the current group;
    ///                corner treatment (none / miter / butt 1 / butt 2) and gap
    ///   Locate       material side Inside / Outside (the hug toggle), rotation, pierce point, mirror,
    ///                optional alignment line for the profile's "up"
    ///
    /// OK builds through <see cref="MemberBuilder"/> (in AfterClose, once the page is gone).
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class MemberPage : IPropertyManagerPage2Handler9
    {
        // control ids
        const int GrpProfile = 100, IdStandard = 101, IdType = 102, IdSize = 103, IdTable = 104, IdGauge = 105, IdThickness = 106, IdRadius = 107, IdProfileInfo = 108;
        const int GrpGroups = 200, IdGroupList = 201, IdNewGroup = 202, IdDelGroup = 203;
        const int GrpPath = 300, IdPathSel = 301, IdCorner = 302, IdGap = 303, IdPreview = 304;
        const int GrpLocate = 400, IdSideIn = 401, IdSideOut = 402, IdRotation = 403, IdPierce = 404, IdMirrorH = 405, IdMirrorV = 406, IdAlignSel = 407;
        const int MarkPath = 1, MarkAlign = 2;
        const string NoTable = "(none: type the thickness)";

        readonly ISldWorks sw;
        readonly IModelDoc2 model;
        readonly ISelectionMgr selMgr;
        IPropertyManagerPage2? page;

        IPropertyManagerPageCombobox? cbStandard, cbType, cbSize, cbTable, cbGauge, cbCorner, cbPierce;
        IPropertyManagerPageNumberbox? nbThickness, nbRadius, nbGap, nbRotation;
        IPropertyManagerPageSelectionbox? selPath, selAlign;
        IPropertyManagerPageListbox? lbGroups;
        IPropertyManagerPageOption? optIn, optOut;
        IPropertyManagerPageCheckbox? ckMirrorH, ckMirrorV, ckPreview;
        IPropertyManagerPageLabel? lblProfile;
        MemberPreview? preview;

        readonly ProfileCatalog catalog;
        readonly string catalogSource;
        readonly List<string> tableFiles;
        GaugeTable? table;
        Profile? profile;
        ProfileSize? size;

        sealed class GroupState
        {
            public MemberGroup G = new MemberGroup();
            public List<object> Path = new List<object>();
            public object? Align;
        }
        readonly List<GroupState> groups = new List<GroupState>();
        int cur = 0;
        bool okPressed;
        bool loading;       // true while the controls are being written from state (ignore change events)

        public MemberPage(ISldWorks sw)
        {
            this.sw = sw;
            model = sw.ActiveDoc as IModelDoc2 ?? throw new InvalidOperationException("open a part first");
            if (model.GetType() != (int)swDocumentTypes_e.swDocPART) throw new InvalidOperationException("Structural Sheet Metal works in a part document");
            selMgr = (ISelectionMgr)model.SelectionManager;
            catalog = ProfileCatalog.LoadOrDefault(SwAddin.BaseDir, out catalogSource);
            var folders = GaugeTable.SplitFolders(sw.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetMetalGaugeTable));
            tableFiles = GaugeTable.ListFiles(folders);
            groups.Add(new GroupState { G = new MemberGroup { Name = "Group1" } });
        }

        // ------------------------------------------------------------------ build the page
        public void Show()
        {
            int errors = 0;
            int options = (int)(swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton
                              | swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton
                              | swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton);
            page = sw.CreatePropertyManagerPage(SwAddin.Title, options, this, ref errors) as IPropertyManagerPage2;
            if (page == null) throw new InvalidOperationException("CreatePropertyManagerPage failed (" + errors + ")");
            page.SetMessage3("Select sketch lines or straight edges for the path, pick a profile and a gauge, then OK. " +
                             "Inside puts the sheet on the profile's concave side; Outside flips it, so two members on the same path hug each other.",
                             (int)swPropertyManagerPageMessageVisibility.swMessageBoxVisible, (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand, "Structural Sheet Metal Member");

            const int vis = (int)(swAddControlOptions_e.swControlOptions_Visible | swAddControlOptions_e.swControlOptions_Enabled);
            const short label = (short)swPropertyManagerPageControlType_e.swControlType_Label;
            const short button = (short)swPropertyManagerPageControlType_e.swControlType_Button;
            const short listbox = (short)swPropertyManagerPageControlType_e.swControlType_Listbox;
            const short selbox = (short)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
            const short option = (short)swPropertyManagerPageControlType_e.swControlType_Option;
            const short check = (short)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            const short left = (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            const short indent = (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
            int grpOpts = (int)(swAddGroupBoxOptions_e.swGroupBoxOptions_Visible | swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded);

            // --- Profile
            var gp = (IPropertyManagerPageGroup)page.AddGroupBox(GrpProfile, "Profile", grpOpts);
            Lbl(gp, "Standard"); cbStandard = Combo(gp, IdStandard, "Catalogue standard");
            Lbl(gp, "Type"); cbType = Combo(gp, IdType, "Profile type (angle, channel, zee, hat...)");
            Lbl(gp, "Size"); cbSize = Combo(gp, IdSize, "Size row of the catalogue");
            lblProfile = (IPropertyManagerPageLabel)gp.AddControl2(IdProfileInfo, label, "", left, vis, "");
            Lbl(gp, "Gauge table"); cbTable = Combo(gp, IdTable, "Sheet metal gauge table (Tools > Options > File Locations)");
            Lbl(gp, "Gauge"); cbGauge = Combo(gp, IdGauge, "Gauge row: thickness and the smallest usable radius come from the table");
            Lbl(gp, "Thickness"); nbThickness = Number(gp, IdThickness, "Sheet thickness", (int)swNumberboxUnitType_e.swNumberBox_Length, 0.0001, 0.05, 0.0005, 0.0015);
            Lbl(gp, "Inside bend radius"); nbRadius = Number(gp, IdRadius, "Inside bend radius", (int)swNumberboxUnitType_e.swNumberBox_Length, 0, 0.1, 0.0005, 0.0015);

            // --- Groups
            var gg = (IPropertyManagerPageGroup)page.AddGroupBox(GrpGroups, "Groups", grpOpts);
            lbGroups = (IPropertyManagerPageListbox)gg.AddControl2(IdGroupList, listbox, "Groups", left, vis, "Each group has its own path, corner treatment and material side");
            lbGroups.Height = 40;
            gg.AddControl2(IdNewGroup, button, "New Group", left, vis, "Start another group (for example the member that hugs the first one)");
            gg.AddControl2(IdDelGroup, button, "Delete Group", left, vis, "Remove the current group");

            // --- Path
            var gs = (IPropertyManagerPageGroup)page.AddGroupBox(GrpPath, "Path segments", grpOpts);
            selPath = (IPropertyManagerPageSelectionbox)gs.AddControl2(IdPathSel, selbox, "Path segments", left, vis, "Sketch lines (2D or 3D) or straight model edges; connected segments get corner treatment");
            selPath.SetSelectionFilters(new[] { (int)swSelectType_e.swSelSKETCHSEGS, (int)swSelectType_e.swSelEXTSKETCHSEGS, (int)swSelectType_e.swSelEDGES });
            selPath.Mark = MarkPath;
            selPath.SingleEntityOnly = false;
            selPath.AllowMultipleSelectOfSameEntity = false;
            selPath.Height = 60;
            Lbl(gs, "Corner treatment"); cbCorner = Combo(gs, IdCorner, "What happens where two segments of this group meet");
            cbCorner.AddItems(new[] { "None", "End miter", "End butt 1 (first through)", "End butt 2 (second through)" });
            Lbl(gs, "Gap"); nbGap = Number(gs, IdGap, "Gap between the two members at a corner", (int)swNumberboxUnitType_e.swNumberBox_Length, 0, 0.05, 0.0005, 0);
            ckPreview = (IPropertyManagerPageCheckbox)gs.AddControl2(IdPreview, check, "Show preview", left, vis, "Temporary bodies: every member as flat legs (no bends, no corner cuts); the current group in blue");
            ckPreview.Checked = true;

            // --- Locate profile
            var gl = (IPropertyManagerPageGroup)page.AddGroupBox(GrpLocate, "Locate profile", grpOpts);
            Lbl(gl, "Material side");
            optIn = (IPropertyManagerPageOption)gl.AddControl2(IdSideIn, option, "Inside (concave side of the profile)", indent, vis, "Sheet goes toward the profile's inner corner");
            optOut = (IPropertyManagerPageOption)gl.AddControl2(IdSideOut, option, "Outside (reverse)", indent, vis, "Sheet goes away from the profile: hugs an Inside member on the same path");
            // the angle box works in RADIANS through the API (it displays degrees); 5° steps
            Lbl(gl, "Rotation angle"); nbRotation = Number(gl, IdRotation, "Rotate the profile about the path", (int)swNumberboxUnitType_e.swNumberBox_Angle, -2 * Math.PI, 2 * Math.PI, 5 * Math.PI / 180, 0);
            Lbl(gl, "Pierce point"); cbPierce = Combo(gl, IdPierce, "Which point of the profile sits on the path");
            ckMirrorH = (IPropertyManagerPageCheckbox)gl.AddControl2(IdMirrorH, check, "Mirror horizontal", left, vis, "Flip the profile left/right");
            ckMirrorV = (IPropertyManagerPageCheckbox)gl.AddControl2(IdMirrorV, check, "Mirror vertical", left, vis, "Flip the profile up/down");
            selAlign = (IPropertyManagerPageSelectionbox)gl.AddControl2(IdAlignSel, selbox, "Align (up direction)", left, vis, "Optional: a line or edge that defines the profile's +Y; default is the path sketch's normal");
            selAlign.SetSelectionFilters(new[] { (int)swSelectType_e.swSelSKETCHSEGS, (int)swSelectType_e.swSelEXTSKETCHSEGS, (int)swSelectType_e.swSelEDGES });
            selAlign.Mark = MarkAlign;
            selAlign.SingleEntityOnly = true;
            selAlign.Height = 20;

            // populate
            loading = true;
            cbStandard.AddItems(catalog.Standards().ToArray());
            cbStandard.CurrentSelection = 0;
            FillTypes();
            cbTable.AddItems(new[] { NoTable }.Concat(tableFiles.Select(Path.GetFileNameWithoutExtension)).ToArray());
            cbTable.CurrentSelection = 0;
            cbGauge.AddItems(new[] { "-" });
            cbGauge.CurrentSelection = 0;
            RefreshGroupList();
            WriteGroupControls();
            loading = false;

            // pre-selected lines become the first group's path
            var pre = new List<object>();
            int n = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= n; i++) { string why; var ent = selMgr.GetSelectedObject6(i, -1); if (PathChain.FromEntity(sw, ent, out why) != null) pre.Add(ent); }
            model.ClearSelection2(true);
            preview = new MemberPreview(sw, model);
            page.Show2(0);
            if (pre.Count > 0) { groups[cur].Path = pre; Reselect(groups[cur]); }
            RefreshPreview();
        }

        public void Close() { try { page?.Close(false); } catch { } }

        /// <summary>Redraw the temporary-body preview from the current state (cheap: a few boxes per member).</summary>
        void RefreshPreview()
        {
            if (preview == null) return;
            try
            {
                if (ckPreview == null || !ckPreview.Checked || profile == null) { preview.Clear(); return; }
                var spec = BuildSpec();
                preview.Show(spec, cur);
            }
            catch (Exception e) { SwAddin.Log("preview: " + e.Message); try { preview.Clear(); } catch { } }
        }

        void Lbl(IPropertyManagerPageGroup g, string text)
        {
            g.AddControl2(0, (short)swPropertyManagerPageControlType_e.swControlType_Label, text, (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                (int)(swAddControlOptions_e.swControlOptions_Visible | swAddControlOptions_e.swControlOptions_Enabled), "");
        }

        IPropertyManagerPageCombobox Combo(IPropertyManagerPageGroup g, int id, string tip)
        {
            var c = (IPropertyManagerPageCombobox)g.AddControl2(id, (short)swPropertyManagerPageControlType_e.swControlType_Combobox, "", (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                (int)(swAddControlOptions_e.swControlOptions_Visible | swAddControlOptions_e.swControlOptions_Enabled), tip);
            c.Style = (int)swPropMgrPageComboBoxStyle_e.swPropMgrPageComboBoxStyle_EditBoxReadOnly;
            c.Height = 120;
            return c;
        }

        IPropertyManagerPageNumberbox Number(IPropertyManagerPageGroup g, int id, string tip, int unit, double min, double max, double inc, double value)
        {
            var nb = (IPropertyManagerPageNumberbox)g.AddControl2(id, (short)swPropertyManagerPageControlType_e.swControlType_Numberbox, "", (short)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge,
                (int)(swAddControlOptions_e.swControlOptions_Visible | swAddControlOptions_e.swControlOptions_Enabled), tip);
            nb.SetRange2(unit, min, max, true, inc, inc * 10, inc / 10);
            nb.Value = value;
            return nb;
        }

        // ------------------------------------------------------------------ catalogue cascade
        string Current(IPropertyManagerPageCombobox? cb)
        {
            if (cb == null) return "";
            short i = cb.CurrentSelection;
            return i < 0 ? "" : cb.get_ItemText(i);
        }

        void FillTypes()
        {
            cbType!.Clear();
            cbType.AddItems(catalog.Types(Current(cbStandard)).ToArray());
            cbType.CurrentSelection = 0;
            FillSizes();
        }

        void FillSizes()
        {
            cbSize!.Clear();
            cbSize.AddItems(catalog.SizesOf(Current(cbStandard), Current(cbType)).Select(s => s.Size).ToArray());
            cbSize.CurrentSelection = 0;
            ResolveProfile();
        }

        void ResolveProfile()
        {
            size = catalog.Find(Current(cbStandard), Current(cbType), Current(cbSize));
            profile = size?.Build();
            var keep = Current(cbPierce);
            cbPierce!.Clear();
            if (profile != null)
            {
                var names = profile.NamedPoints.Select(kv => kv.Key).ToArray();
                cbPierce.AddItems(names);
                int at = Array.IndexOf(names, keep);
                cbPierce.CurrentSelection = (short)(at >= 0 ? at : 0);
                lblProfile!.Caption = string.Format("{0}: {1}", size!.Shape, string.Join(", ", size.Params.Select(kv => kv.Key + " = " + kv.Value.ToString("0.##") + " mm")));
            }
            else lblProfile!.Caption = "(no profile)";
        }

        void TableChanged()
        {
            short i = cbTable!.CurrentSelection;
            table = null;
            cbGauge!.Clear();
            if (i > 0 && i - 1 < tableFiles.Count)
            {
                try { table = GaugeTable.Load(tableFiles[i - 1]); }
                catch (Exception e) { sw.SendMsgToUser2("Gauge table could not be read: " + e.Message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk); }
            }
            if (table != null)
            {
                cbGauge.AddItems(table.Rows.Select(r => r.Name + "  (" + r.ThicknessMm.ToString("0.###") + " mm)").ToArray());
                // default: the row nearest the current thickness
                double t = nbThickness!.Value * 1000;
                int best = 0; double bd = double.MaxValue;
                for (int k = 0; k < table.Rows.Count; k++) { double d = Math.Abs(table.Rows[k].ThicknessMm - t); if (d < bd) { bd = d; best = k; } }
                cbGauge.CurrentSelection = (short)best;
                GaugeChanged();
            }
            else
            {
                cbGauge.AddItems(new[] { "-" });
                cbGauge.CurrentSelection = 0;
            }
            ((IPropertyManagerPageControl)nbThickness!).Enabled = table == null;
        }

        void GaugeChanged()
        {
            if (table == null) return;
            short i = cbGauge!.CurrentSelection;
            if (i < 0 || i >= table.Rows.Count) return;
            var row = table.Rows[i];
            nbThickness!.Value = row.ThicknessMm / 1000;
            nbRadius!.Value = GaugeTable.DefaultRadiusMm(row) / 1000;
        }

        // ------------------------------------------------------------------ groups
        void RefreshGroupList()
        {
            lbGroups!.Clear();
            lbGroups.AddItems(groups.Select(g => g.G.Name + "  (" + g.Path.Count + ")").ToArray());
            lbGroups.CurrentSelection = (short)cur;
        }

        /// <summary>Controls ← current group.</summary>
        void WriteGroupControls()
        {
            var g = groups[cur].G;
            bool was = loading; loading = true;
            cbCorner!.CurrentSelection = (short)g.Corner;
            nbGap!.Value = g.GapMm / 1000;
            optIn!.Checked = g.Side == MaterialSide.Inside;
            optOut!.Checked = g.Side == MaterialSide.Outside;
            nbRotation!.Value = g.RotationDeg * Math.PI / 180;
            ckMirrorH!.Checked = g.MirrorHorizontal;
            ckMirrorV!.Checked = g.MirrorVertical;
            if (g.PiercePoint != null && profile != null)
            {
                var names = profile.NamedPoints.Select(kv => kv.Key).ToList();
                int at = names.IndexOf(g.PiercePoint);
                if (at >= 0) cbPierce!.CurrentSelection = (short)at;
            }
            loading = was;
        }

        /// <summary>Current group ← controls.</summary>
        void ReadGroupControls()
        {
            var g = groups[cur].G;
            g.Corner = (CornerTreatment)Math.Max(0, (int)cbCorner!.CurrentSelection);
            g.GapMm = nbGap!.Value * 1000;
            g.Side = optOut!.Checked ? MaterialSide.Outside : MaterialSide.Inside;
            g.RotationDeg = nbRotation!.Value * 180 / Math.PI;
            g.MirrorHorizontal = ckMirrorH!.Checked;
            g.MirrorVertical = ckMirrorV!.Checked;
            g.PiercePoint = Current(cbPierce);
        }

        /// <summary>Selection set ← the group's stored entities.</summary>
        void Reselect(GroupState g)
        {
            model.ClearSelection2(true);
            foreach (var ent in g.Path) SelectEntity(ent, MarkPath);
            if (g.Align != null) SelectEntity(g.Align, MarkAlign);
        }

        void SelectEntity(object ent, int mark)
        {
            var sd = (SelectData)selMgr.CreateSelectData(); sd.Mark = mark;
            var seg = ent as ISketchSegment;
            if (seg != null) { seg.Select4(true, sd); return; }
            var e = ent as IEntity;
            if (e != null) e.Select4(true, sd);
        }

        /// <summary>The group's stored entities ← selection set (called when a selection box changes).</summary>
        void SyncFromSelection()
        {
            var g = groups[cur];
            g.Path.Clear();
            int n = selMgr.GetSelectedObjectCount2(MarkPath);
            for (int i = 1; i <= n; i++) g.Path.Add(selMgr.GetSelectedObject6(i, MarkPath));
            g.Align = selMgr.GetSelectedObjectCount2(MarkAlign) > 0 ? selMgr.GetSelectedObject6(1, MarkAlign) : null;
            RefreshGroupList();
        }

        void SwitchGroup(int to)
        {
            if (to < 0 || to >= groups.Count || to == cur) return;
            ReadGroupControls();
            SyncFromSelection();
            cur = to;
            Reselect(groups[cur]);
            WriteGroupControls();
            RefreshGroupList();
            RefreshPreview();
        }

        // ------------------------------------------------------------------ the build
        MemberSpec BuildSpec()
        {
            if (profile == null) throw new InvalidOperationException("no profile selected");
            var spec = new MemberSpec
            {
                Name = "Structural Sheet Metal",
                Profile = profile,
                ThicknessMm = nbThickness!.Value * 1000,
                BendRadiusMm = nbRadius!.Value * 1000,
                KFactor = table != null && table.KFactor.HasValue ? table.KFactor.Value : 0.5
            };
            if (table != null)
            {
                short i = cbGauge!.CurrentSelection;
                if (i >= 0 && i < table.Rows.Count) { spec.GaugeTablePath = table.Path; spec.GaugeName = table.Rows[i].RawName; }
            }
            var rejected = new List<string>();
            foreach (var gs in groups)
            {
                var g = gs.G;
                g.Segments = PathChain.FromEntities(sw, gs.Path, rejected);
                g.AlignUp = null;
                if (gs.Align != null)
                {
                    string why;
                    var a = PathChain.FromEntity(sw, gs.Align, out why);
                    if (a != null) g.AlignUp = a.Dir;
                }
                if (g.Segments.Count > 0) spec.Groups.Add(g);
            }
            if (rejected.Count > 0) SwAddin.Log("rejected: " + string.Join("; ", rejected));
            return spec;
        }

        void RunBuild()
        {
            BuildReport rep;
            try
            {
                var spec = BuildSpec();
                rep = MemberBuilder.Build(sw, model, spec);
            }
            catch (Exception e)
            {
                rep = new BuildReport { Ok = false, Error = e };
                rep.Log("ERROR: " + e.Message);
            }
            SwAddin.Log(rep.ToString());
            if (!rep.Ok || rep.Warnings > 0)
            {
                var tail = string.Join("\n", rep.Lines.Where(l => l.StartsWith("WARNING") || l.StartsWith("ERROR")).Take(8));
                sw.SendMsgToUser2(rep.Summary + "\n\n" + tail + "\n\nFull log: " + SwAddin.LogPath,
                    rep.Ok ? (int)swMessageBoxIcon_e.swMbWarning : (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        // ------------------------------------------------------------------ IPropertyManagerPage2Handler9
        public void AfterActivation() { }
        public void OnClose(int reason)
        {
            okPressed = reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay;
            if (okPressed)
            {
                // read everything now: the controls are gone by AfterClose
                ReadGroupControls();
                SyncFromSelection();
            }
            try { preview?.Clear(); } catch { }
        }
        public void AfterClose()
        {
            try { preview?.Dispose(); } catch { }
            preview = null;
            if (okPressed) RunBuild();
            page = null;
        }
        public bool OnHelp() { return false; }
        public bool OnPreviousPage() { return false; }
        public bool OnNextPage() { return false; }
        public bool OnPreview() { return false; }
        public void OnWhatsNew() { }
        public void OnUndo() { }
        public void OnRedo() { }
        public bool OnTabClicked(int id) { return false; }
        public void OnGroupExpand(int id, bool expanded) { }
        public void OnGroupCheck(int id, bool chk) { }
        public void OnCheckboxCheck(int id, bool chk) { if (loading) return; if (id != IdPreview) ReadGroupControls(); RefreshPreview(); }
        public void OnOptionCheck(int id) { if (loading) return; ReadGroupControls(); RefreshPreview(); }
        public void OnButtonPress(int id)
        {
            if (id == IdNewGroup)
            {
                ReadGroupControls();
                SyncFromSelection();
                var g = new GroupState { G = new MemberGroup { Name = "Group" + (groups.Count + 1) } };
                // a new group starts with the current group's settings but no path
                var c = groups[cur].G;
                g.G.Side = c.Side; g.G.Corner = c.Corner; g.G.GapMm = c.GapMm; g.G.RotationDeg = c.RotationDeg; g.G.PiercePoint = c.PiercePoint;
                g.G.MirrorHorizontal = c.MirrorHorizontal; g.G.MirrorVertical = c.MirrorVertical;
                groups.Add(g);
                cur = groups.Count - 1;
                model.ClearSelection2(true);
                WriteGroupControls();
                RefreshGroupList();
                selPath!.SetSelectionFocus();
            }
            else if (id == IdDelGroup && groups.Count > 1)
            {
                groups.RemoveAt(cur);
                cur = Math.Max(0, cur - 1);
                for (int i = 0; i < groups.Count; i++) groups[i].G.Name = "Group" + (i + 1);
                Reselect(groups[cur]);
                WriteGroupControls();
                RefreshGroupList();
            }
            RefreshPreview();
        }
        public void OnTextboxChanged(int id, string text) { }
        public void OnNumberboxChanged(int id, double value) { if (loading) return; ReadGroupControls(); RefreshPreview(); }
        public void OnNumberBoxTrackingCompleted(int id, double value) { }
        public void OnComboboxEditChanged(int id, string text) { }
        public void OnComboboxSelectionChanged(int id, int item)
        {
            if (loading) return;
            loading = true;
            try
            {
                switch (id)
                {
                    case IdStandard: FillTypes(); break;
                    case IdType: FillSizes(); break;
                    case IdSize: ResolveProfile(); break;
                    case IdTable: TableChanged(); break;
                    case IdGauge: GaugeChanged(); break;
                    default: ReadGroupControls(); break;
                }
            }
            finally { loading = false; }
            RefreshPreview();
        }
        public void OnListboxSelectionChanged(int id, int item) { if (id == IdGroupList && !loading) SwitchGroup(item); }
        public void OnListboxRMBUp(int id, int posX, int posY) { }
        public void OnSelectionboxFocusChanged(int id) { }
        public void OnSelectionboxListChanged(int id, int count) { if (loading) return; SyncFromSelection(); RefreshPreview(); }
        public void OnSelectionboxCalloutCreated(int id) { }
        public void OnSelectionboxCalloutDestroyed(int id) { }
        public bool OnSubmitSelection(int id, object selection, int selType, ref string itemText)
        {
            string why;
            if (PathChain.FromEntity(sw, selection, out why) != null) return true;
            page?.SetMessage3("Not usable: " + why, (int)swPropertyManagerPageMessageVisibility.swImportantMessageBox, (int)swPropertyManagerPageMessageExpanded.swMessageBoxExpand, "Structural Sheet Metal Member");
            return false;
        }
        public int OnActiveXControlCreated(int id, bool status) { return -1; }
        public void OnSliderPositionChanged(int id, double value) { }
        public void OnSliderTrackingCompleted(int id, double value) { }
        public bool OnKeystroke(int wparam, int message, int lparam, int id) { return false; }
        public void OnPopupMenuItem(int id) { }
        public void OnPopupMenuItemUpdate(int id, ref int retval) { }
        public void OnGainedFocus(int id) { }
        public void OnLostFocus(int id) { }
        public int OnWindowFromHandleControlCreated(int id, bool status) { return 0; }
    }
}
