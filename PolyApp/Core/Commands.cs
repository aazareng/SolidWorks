using System;
using System.Collections.Generic;
using System.Linq;

namespace PolyApp
{
    /// <summary>Document types a command is enabled in (own flags; see <see cref="DocContext"/> for the SOLIDWORKS mapping).</summary>
    [Flags]
    public enum Docs
    {
        None = 0,
        Part = 1,
        Assembly = 2,
        Drawing = 4,
        Model = Part | Assembly,
        Any = Part | Assembly | Drawing,
    }

    /// <summary>
    /// One PolyApp command. Until it is ported to C#, a command runs the PolyAppPro macro: the add-in
    /// writes the request (which frm_Main option(s) to set) and calls the macro's PA_Run entry point.
    /// </summary>
    public sealed class Command
    {
        public string Id = "";
        public string Title = "";
        public string Hint = "";
        /// <summary>Two to four letters drawn on the generated icon until a real one is dropped in AddIn\Icons.</summary>
        public string Code = "";
        /// <summary>CommandManager tab box ("PDM", "Make", ...).</summary>
        public string Box = "";
        /// <summary>Flyout (dropdown) inside the box, or null for a plain button.</summary>
        public string? Flyout;
        public Docs Docs = Docs.Any;
        /// <summary>frm_Main control the macro sets before Run; a second control for the Make Assembly family.</summary>
        public string MacroOption = "";
        public string MacroOption2 = "";
        /// <summary>Also offered on the right-click menu of a component.</summary>
        public bool ContextMenu;

        public override string ToString() { return Id; }
    }

    /// <summary>The command catalogue: the tab layout, enable rules and macro mapping all derive from this list.</summary>
    public static class Catalog
    {
        static Command C(string id, string title, string code, string box, string? flyout, Docs docs, string opt, string opt2 = "", string hint = "", bool ctx = false)
        {
            return new Command { Id = id, Title = title, Code = code, Box = box, Flyout = flyout, Docs = docs, MacroOption = opt, MacroOption2 = opt2, Hint = hint.Length > 0 ? hint : title, ContextMenu = ctx };
        }

        public static readonly Command[] All =
        {
            // ---------------------------------------------------------------- PDM
            C("PdmCheckOut",  "Check Out",       "OUT", "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_CheckedOut",    "Check out the active document (and the selected components)", ctx: true),
            C("PdmCheckIn",   "Check In",        "IN",  "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_CheckedIn",     "Error-check, then check in", ctx: true),
            C("PdmIssue",     "Issue",           "ISS", "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_Issued",        "Check in and set the Issued state"),
            C("PdmIssueNext", "Issue Next Rev",  "IS+", "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_IssuedNextRev", "Issue and bump the revision"),
            C("PdmBanish",    "Banish",          "BAN", "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_Banish"),
            C("PdmObsolete",  "Obsolete",        "OBS", "PDM", "PDM State", Docs.Any, "opt_pdmState_GEN", "opt_Obsolete"),
            C("Branch",       "Branch",          "BR",  "PDM", null,        Docs.Model, "opt_Branch_MDL", "", "Copy the document (or the selected components) to the next serial/section number", ctx: true),
            C("Open",         "Open",            "OPN", "PDM", null,        Docs.Any,   "opt_OPEN_GEN",   "", "Open the drawing, right hand, left hand and parent of the active document", ctx: true),
            C("OpenDrawing",  "Open Drawing",    "DWG", "PDM", null,        Docs.Model, "opt_OPEN_GEN",   "", "Open the drawing beside the model (native; falls back to the macro's vault search)", ctx: true),

            // ---------------------------------------------------------------- Make
            C("MakeDrawing",  "Make Drawing",    "MKD", "Make", null,            Docs.Model,    "opt_MAKEDWG_MDL", "", "New drawing from the template, views placed, title block filled"),
            C("Hotsheet",     "Hotsheet",        "HOT", "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_Hotsheet_MDL", "Hotsheet assembly from the selected components", ctx: true),
            C("ShippingOrder","Shipping Order",  "SO",  "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_SO_MDL"),
            C("Shipping",     "Shipping (SH)",   "SH",  "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_SH_MDL"),
            C("Por",          "POR",             "POR", "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_POR_MDL", "Purchase order request assembly"),
            C("Manual",       "Manual",          "MAN", "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_MANUAL_MDL"),
            C("Spec",         "Spec",            "SPC", "Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_SPEC_MDL"),
            C("CutsheetAssy", "Cutsheet Assembly","CSA","Make", "Make Assembly", Docs.Model,    "opt_ASSY_MDL", "opt_Cutsheet_MDL"),
            C("Modified",     "Modified (-M)",   "-M",  "Make", "Make Assembly", Docs.Model,    "", "", "Modified part/assembly (-M): map its frm_Main control in [MacroOptions] at work"),
            C("Cutsheet",     "Cutsheet",        "CUT", "Make", null,            Docs.Assembly, "opt_Cutsheet_ASM", "", "Error check, eDrawing, check in, queue and write the cutsheet"),
            C("MateOrigins",  "Mate Origins",    "MO",  "Make", null,            Docs.Assembly, "opt_OGMate", "", "Mate the inserted components' origins to the envelope assembly"),
            C("Mirrorize",    "Mirrorize",       "MIR", "Make", null,            Docs.Model,    "opt_MIRROR_MDL", "", "Opposite-hand part by mirror"),

            // ---------------------------------------------------------------- Select (Power Select)
            C("SelSheetMetal","Sheet Metal",     "SM",  "Select", "By Type", Docs.Assembly, "opt_SelSM"),
            C("SelAssemblies","Assemblies",      "ASM", "Select", "By Type", Docs.Assembly, "opt_SelASM"),
            C("SelParts",     "Parts",           "PRT", "Select", "By Type", Docs.Assembly, "opt_SelTLAPrt"),
            C("SelVirtual",   "Virtual",         "VIR", "Select", "By Type", Docs.Assembly, "opt_SelVirtual"),
            C("SelExcluded",  "Excluded from BOM","EXC","Select", "By Type", Docs.Assembly, "opt_SelExclude"),
            C("SelEnvelope",  "Envelopes",       "ENV", "Select", "By Type", Docs.Assembly, "opt_SelEnvelope"),
            C("SelFixed",     "Fixed",           "FIX", "Select", "By Type", Docs.Assembly, "opt_SelFixed"),
            C("SelExtRef",    "External Refs",   "REF", "Select", "By Type", Docs.Assembly, "opt_SelExtRef"),
            C("SelPre13",     "Pre-2013",        "P13", "Select", "By Type", Docs.Assembly, "opt_SelPre13"),
            C("SelUnderDef",  "Under Defined",   "UND", "Select", "By Type", Docs.Assembly, "opt_SelUnderDefined"),
            C("SelVisible",   "Visible",         "VIS", "Select", "By Type", Docs.Assembly, "opt_SelVis"),
            C("SelHidden",    "Hidden",          "HID", "Select", "By Type", Docs.Assembly, "opt_SelHid"),

            C("SelRouter",    "Router",          "RTR", "Select", "By Department", Docs.Assembly, "opt_SelROUTER"),
            C("SelMill",      "Mill",            "MIL", "Select", "By Department", Docs.Assembly, "opt_SelMILL"),
            C("SelLathe",     "Lathe",           "LTH", "Select", "By Department", Docs.Assembly, "opt_SelLATHE"),
            C("Sel3D",        "3D Print",        "3D",  "Select", "By Department", Docs.Assembly, "opt_Sel3D"),
            C("SelShear",     "Shear",           "SHR", "Select", "By Department", Docs.Assembly, "opt_SelSHEAR"),
            C("SelManual",    "Manual",          "MNL", "Select", "By Department", Docs.Assembly, "opt_SelMANUAL"),
            C("SelPurchased", "Purchased",       "BUY", "Select", "By Department", Docs.Assembly, "opt_SelPurch"),
            C("SelStandard",  "Standard Parts",  "SP",  "Select", "By Department", Docs.Assembly, "opt_SelSP"),
            C("SelSpec",      "Spec",            "SPC", "Select", "By Department", Docs.Assembly, "opt_SelSPEC"),

            C("SelSerial",    "Serial Number",   "SN",  "Select", "By State", Docs.Assembly, "opt_SelSN", "", "Components carrying the active serial number"),
            C("SelIssued",    "Issued",          "ISS", "Select", "By State", Docs.Assembly, "opt_SelIssued"),
            C("SelRight",     "Right Hand",      "RH",  "Select", "By State", Docs.Assembly, "opt_SelRight"),
            C("SelConfig",    "Configuration",   "CFG", "Select", "By State", Docs.Assembly, "opt_SelConf"),
            C("SelRedFace",   "Red Face",        "RED", "Select", "By State", Docs.Assembly, "opt_SelRed"),
            C("SelFlat",      "Flat Pattern",    "FLT", "Select", "By State", Docs.Assembly, "opt_SelFlat"),
            C("SelectTree",   "Tree",            "TRE", "Select", null,       Docs.Assembly, "", "", "Toggle: Power Select walks the whole tree (on) or the top level only (off)"),

            // ---------------------------------------------------------------- Export
            C("Dxf",          "DXF",             "DXF", "Export", null, Docs.Model,   "opt_DXF_MDL", "", "Sheet-metal flat DXFs (the PolySW DXFer replaces this)"),
            C("Pdf",          "PDF",             "PDF", "Export", null, Docs.Any,     "opt_PDF_GEN"),
            C("EDrawings",    "eDrawings",       "EDW", "Export", null, Docs.Any,     "opt_eDWG_GEN"),
            C("Step",         "STEP",            "STP", "Export", null, Docs.Model,   "opt_STEP_MDL"),
            C("ExportBom",    "Export BOM",      "BOM", "Export", null, Docs.Drawing, "opt_ExportBOM_DWG", "", "Check balloons, then BOM to Excel"),
            C("ExportHtml",   "Export HTML",     "HTM", "Export", null, Docs.Drawing, "opt_ExportHTML_DWG", "", "Manual drawing to HTML"),

            // ---------------------------------------------------------------- Drawing
            C("ReloadFormat", "Reload Format",   "RLD", "Drawing", null, Docs.Drawing, "opt_Reload_DWG", "", "Reload the sheet format and keep the revision"),
            C("CheckBalloons","Check Balloons",  "BLN", "Drawing", null, Docs.Drawing, "opt_CheckBalloons_DWG", "", "Every BOM row has a balloon; hide the ones that do not"),

            // ---------------------------------------------------------------- More
            C("AddConfig",    "Add Config",      "+CF", "More", "More", Docs.Model,    "opt_AddConfig"),
            C("SyncConfig",   "Sync Config",     "SYC", "More", "More", Docs.Assembly, "opt_SyncConfig"),
            C("Configger",    "Configger",       "CFR", "More", "More", Docs.Assembly, "opt_Configger", "", "Mate and component states per configuration token"),
            C("Colorize",     "Colorize",        "COL", "More", "More", Docs.Assembly, "opt_COLOR_ASM"),
            C("Pinkify",      "Pinkify",         "PNK", "More", "More", Docs.Model,    "opt_Pinkify"),
            C("PurgeConfigs", "Purge Configs",   "PRG", "More", "More", Docs.Model,    "opt_PurgeConfigs_MDL"),
            C("TrimProps",    "Trim Properties", "TRP", "More", "More", Docs.Model,    "opt_TrimProp_MDL"),
            C("TrimConfigs",  "Trim Configs",    "TRC", "More", "More", Docs.Model,    "opt_TrimConf_MDL"),
            C("CutSize",      "Cut Size",        "L×W", "More", "More", Docs.Model,    "opt_CUTSIZE_MDL", "", "Sheet-metal length and width properties"),
            C("SetMaterial",  "Set Material",    "MAT", "More", "More", Docs.Model,    "opt_SetMatl_MDL"),
            C("CsvAssembly",  "CSV Assembly",    "CSV", "More", "More", Docs.Assembly, "opt_CSV_ASM", "", "Insert components listed in a CSV"),
            C("ErrorCheck",   "Error Check",     "ERR", "More", "More", Docs.Model,    "opt_ERRORCHK_MDL"),
            C("ReplaceRef",   "Replace Ref",     "RPL", "More", "More", Docs.Any,      "opt_RepRef_GEN"),
            C("Dissolve",     "Dissolve",        "DIS", "More", "More", Docs.Assembly, "opt_Explode_ASM", "", "Purge, delete folders, dissolve patterns and mirrors"),
            C("ExcludeBom",   "Exclude from BOM","XBM", "More", "More", Docs.Assembly, "opt_ExcludeBOM_ASM", "", "Selected components become envelopes"),
            C("AddSpec",      "Add Spec",        "+SP", "More", "More", Docs.Model,    "opt_AddSpec"),
            C("SyncCard",     "Sync Card",       "SYN", "More", "More", Docs.Model,    "opt_SyncCard", "", "Department, stocked part number, spec type/value to the data card"),
            C("DeleteDoc",    "Delete Document", "DEL", "More", "More", Docs.Any,      "opt_DelDoc_GEN"),

            // ---------------------------------------------------------------- Setup (native, always enabled)
            C("Settings",     "Settings",        "INI", "Setup", null, Docs.None, "", "", "Open your PolyApp.ini (per-user overrides of the shipped one)"),
            C("ShowLog",      "Log",             "LOG", "Setup", null, Docs.None, "", "", "Open the PolyApp log"),
        };

        public static Command? Find(string id)
        {
            return All.FirstOrDefault(c => c.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Boxes in tab order.</summary>
        public static IEnumerable<string> Boxes { get { return All.Select(c => c.Box).Distinct(); } }

        public static IEnumerable<Command> InBox(string box) { return All.Where(c => c.Box == box); }

        /// <summary>Flyouts in a box, in catalogue order.</summary>
        public static IEnumerable<string> FlyoutsIn(string box)
        {
            return All.Where(c => c.Box == box && c.Flyout != null).Select(c => c.Flyout!).Distinct();
        }

        public static IEnumerable<Command> InFlyout(string flyout) { return All.Where(c => c.Flyout == flyout); }

        /// <summary>Every box/flyout as a group of commands, in order: plain buttons are their own group of one.</summary>
        public static Docs DocsOf(IEnumerable<Command> group)
        {
            Docs d = Docs.None;
            foreach (var c in group) d |= c.Docs;
            return d;
        }

        /// <summary>Name of the add-in enable callback for a document mask (the add-in defines one per mask).</summary>
        public static string EnableMethod(Docs docs)
        {
            switch (docs)
            {
                case Docs.None: return "Enable_Always";
                case Docs.Part: return "Enable_Part";
                case Docs.Assembly: return "Enable_Assembly";
                case Docs.Drawing: return "Enable_Drawing";
                case Docs.Model: return "Enable_Model";
                default: return "Enable_Any";
            }
        }
    }
}
