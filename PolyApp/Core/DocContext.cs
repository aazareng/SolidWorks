using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace PolyApp
{
    /// <summary>What a command acts on: the active document and the selection, captured at click time.</summary>
    public sealed class DocContext
    {
        public IModelDoc2? Doc;
        public Docs Kind = Docs.None;
        public string Path = "";
        public string Title = "";
        public int Selected;
        public int SelectedComponents;

        public bool HasDoc { get { return Doc != null; } }

        public static Docs KindOf(int swDocType)
        {
            switch (swDocType)
            {
                case (int)swDocumentTypes_e.swDocPART: return Docs.Part;
                case (int)swDocumentTypes_e.swDocASSEMBLY: return Docs.Assembly;
                case (int)swDocumentTypes_e.swDocDRAWING: return Docs.Drawing;
                default: return Docs.None;
            }
        }

        public static DocContext Capture(ISldWorks sw)
        {
            var ctx = new DocContext();
            try
            {
                var doc = sw.ActiveDoc as IModelDoc2;
                if (doc == null) return ctx;
                ctx.Doc = doc;
                ctx.Kind = KindOf(doc.GetType());
                ctx.Path = doc.GetPathName() ?? "";
                ctx.Title = doc.GetTitle() ?? "";
                var sel = doc.SelectionManager as ISelectionMgr;
                if (sel != null)
                {
                    ctx.Selected = sel.GetSelectedObjectCount2(-1);
                    for (int i = 1; i <= ctx.Selected; i++)
                        if (sel.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelCOMPONENTS) ctx.SelectedComponents++;
                }
            }
            catch (Exception e) { Log.Write("DocContext: " + e.Message); }
            return ctx;
        }

        public override string ToString()
        {
            return HasDoc ? Title + " [" + Kind + "] sel=" + Selected + " comps=" + SelectedComponents : "(no document)";
        }
    }
}
