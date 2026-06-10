using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EdmLib;

namespace HelloWorldAddIn
{
    // The CLSID below must match the <Guid> in the .csproj.
    // PDM looks up this GUID in the registry to load the add-in.
    [Guid("5C3F1A2B-7E4D-4A1B-9C3E-1234567890AB")]
    [ComVisible(true)]
    [ProgId("HelloWorldAddIn.AddIn")]
    public class AddIn : IEdmAddIn5
    {
        // Pick any IDs you like, just keep them unique within this add-in.
        private const int CmdID_HelloWorld = 1;
        private const int CmdID_SetToIssued = 2;

        // The PDM state you want files to be moved into. Adjust to match your workflow.
        private const string TargetStateName = "Issued";

        private IEdmVault5 _vault;

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            _vault = poVault;

            poInfo.mbsAddInName = "Hello World Add-in";
            poInfo.mbsCompany = "Aazar Engineering";
            poInfo.mbsDescription = "Sample add-in: shows a Hello World message and changes a file's PDM state.";
            poInfo.mlAddInVersion = 1;
            poInfo.mlRequiredVersionMajor = 1;
            poInfo.mlRequiredVersionMinor = 0;

            // "Hello World" right-click menu command, available on files.
            EdmCmd helloCmd = new EdmCmd
            {
                meType = EdmCmdType.EdmCmd_Menu,
                mlCmdID = CmdID_HelloWorld,
                mbsCmdsetName = "HelloWorldCmdSet",
                mbsName = "Hello World",
                mbsToolTip = "Say hello",
                mbsHint = "Say hello",
                mlCmdType = (int)(EdmMenuFlags.EdmMenu_FileMenu | EdmMenuFlags.EdmMenu_FileTab | EdmMenuFlags.EdmMenu_Enable)
            };
            poCmdMgr.AddCmd(ref helloCmd);

            // "Set to Issued" right-click menu command, available on files.
            EdmCmd stateCmd = new EdmCmd
            {
                meType = EdmCmdType.EdmCmd_Menu,
                mlCmdID = CmdID_SetToIssued,
                mbsCmdsetName = "HelloWorldCmdSet",
                mbsName = "Set to " + TargetStateName,
                mbsToolTip = "Change PDM state of the selected file(s) to " + TargetStateName,
                mbsHint = "Change PDM state to " + TargetStateName,
                mlCmdType = (int)(EdmMenuFlags.EdmMenu_FileMenu | EdmMenuFlags.EdmMenu_FileTab | EdmMenuFlags.EdmMenu_Enable)
            };
            poCmdMgr.AddCmd(ref stateCmd);
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            switch (poCmd.mlCmdID)
            {
                case CmdID_HelloWorld:
                    MessageBox.Show("Hello World!", "HelloWorldAddIn");
                    break;

                case CmdID_SetToIssued:
                    SetFilesState(ppoData, TargetStateName);
                    break;
            }
        }

        private void SetFilesState(EdmCmdData[] selection, string stateName)
        {
            if (selection == null) return;

            foreach (EdmCmdData item in selection)
            {
                try
                {
                    IEdmFile5 file = (IEdmFile5)_vault.GetObject(EdmObjectType.EdmObject_File, item.mlObjectID1);
                    int folderId = item.mlProjID;

                    // Mirrors PDM_src.var_pdmFile.ChangeState in the VBA macro.
                    file.ChangeState(stateName, folderId, "AUTO via HelloWorldAddIn", 0, 0);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to change state: " + ex.Message, "HelloWorldAddIn");
                }
            }
        }
    }
}
