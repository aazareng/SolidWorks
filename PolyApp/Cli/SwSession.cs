using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace PolyApp
{
    /// <summary>Attach to the SOLIDWORKS already running in this Windows session (same user, same elevation).</summary>
    static class SwSession
    {
        const string ProgId = "SldWorks.Application";

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void CLSIDFromProgID(string progId, out Guid clsid);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public static ISldWorks? TryAttach()
        {
            try
            {
                CLSIDFromProgID(ProgId, out var clsid);
                GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
                return obj as ISldWorks;
            }
            catch (COMException e) when ((uint)e.HResult == 0x800401E3) { return null; }   // MK_E_UNAVAILABLE
        }

        public static ISldWorks Connect(bool startIfMissing)
        {
            var sw = TryAttach();
            if (sw != null) return sw;
            if (!startIfMissing) throw new InvalidOperationException("SOLIDWORKS is not running for this user. Open it first, or pass --start.");
            var type = Type.GetTypeFromProgID(ProgId) ?? throw new InvalidOperationException("SOLIDWORKS is not installed (ProgID not registered).");
            sw = (ISldWorks)Activator.CreateInstance(type)!;
            sw.Visible = true;
            return sw;
        }

        public static string Describe(ISldWorks sw)
        {
            return "SOLIDWORKS " + sw.RevisionNumber() + "  pid=" + sw.GetProcessID() + "\nActive document: " + DocContext.Capture(sw);
        }
    }
}
