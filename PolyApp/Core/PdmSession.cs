using System;
#if PDM
using EPDM.Interop.epdm;
#endif

namespace PolyApp
{
    /// <summary>
    /// The PDM vault, logged in once at startup. Compiled against EPDM.Interop.epdm when the PDM client is
    /// installed on the build machine (the PDM define); otherwise every call reports "no PDM" and the macro
    /// side keeps doing the vault work.
    /// </summary>
    public static class PdmSession
    {
        public static bool Available { get; private set; }
        public static string Status { get; private set; } = "not initialised";
        public static string VaultName { get; private set; } = "";

#if PDM
        static IEdmVault5? vault;
        public static IEdmVault5? Vault { get { return vault; } }
#endif

        public static void Init(Settings s)
        {
            Available = false;
            VaultName = s.Get("PDM", "Vault", "_PolyVault");
            var mode = s.PdmMode;
            if (mode == "no") { Status = "disabled in " + Settings.FileName; Log.Write("PDM: " + Status); return; }
#if PDM
            try
            {
                vault = new EdmVault5();
                if (!vault.IsLoggedIn) vault.LoginAuto(VaultName, 0);
                Available = vault.IsLoggedIn;
                Status = Available ? "logged in to " + VaultName + " as " + SafeUser() : "login to " + VaultName + " did not complete";
            }
            catch (Exception e)
            {
                vault = null;
                Status = "no vault view for " + VaultName + " on this PC (" + e.Message.Trim() + ")";
            }
#else
            Status = "built without the PDM interop (EPDM.Interop.epdm.dll not found at build time)";
#endif
            Log.Write("PDM: " + Status);
            if (mode == "yes" && !Available) Log.Write("PDM: [PDM] Enabled=yes but the vault is not available; vault commands will fail");
        }

#if PDM
        static string SafeUser()
        {
            try { return ((IEdmUserMgr5)vault!).GetLoggedInUser()?.Name ?? "?"; } catch { return "?"; }
        }
#endif
    }
}
