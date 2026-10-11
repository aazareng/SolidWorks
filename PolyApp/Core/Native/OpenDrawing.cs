using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace PolyApp.Native
{
    /// <summary>
    /// The first native command and the pattern for the rest: do the common case in C#, return null for
    /// anything else so the macro (which searches the vault) takes over.
    /// </summary>
    public static class OpenDrawing
    {
        public static CommandResult? Run(Dispatcher d, Command c, DocContext ctx)
        {
            if (!ctx.HasDoc || ctx.Kind == Docs.Drawing) return CommandResult.Fail("Open Drawing works from a part or assembly.");
            if (ctx.Path.Length == 0) return CommandResult.Fail("Save the document first; an unsaved model has no drawing.");
            var dwg = Path.ChangeExtension(ctx.Path, ".SLDDRW");
            if (!File.Exists(dwg)) return null;      // not beside the model: the macro's vault search knows where it lives

            int err = 0, warn = 0;
            var doc = d.Sw.OpenDoc6(dwg, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn) as IModelDoc2;
            if (doc == null) return CommandResult.Fail("could not open " + dwg + " (error " + err + ")");
            try { d.Sw.ActivateDoc3(doc.GetTitle(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref err); } catch { }
            return CommandResult.Done("opened " + Path.GetFileName(dwg));
        }
    }
}
