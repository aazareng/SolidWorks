Attribute VB_Name = "mod_PolyAppEntry"
' ============================================================================
'  mod_PolyAppEntry — the PolyApp add-in's way into PolyAppPro.swp.          UNTESTED (written on the home PC, no .swp here)
'
'  The add-in calls   swApp.RunMacro2 "<PolyAppPro.swp>", "mod_PolyAppEntry", "PA_Run", ...
'  RunMacro2 takes no arguments, so the request is in  %TEMP%\PolyApp.request.ini :
'
'      [Request]
'      Command=Branch              the add-in's command id (for the log)
'      Option=opt_Branch_MDL       the frm_Main control to set before Run
'      Option2=                    a second control (the Make Assembly family: opt_ASSY_MDL + opt_Hotsheet_MDL)
'      Tree=0                      Power Select: 1 = whole tree, 0 = top level only  (chk_Tree)
'      Doc=C:\...\A4583.SLDASM     active document when the button was clicked
'
'  PA_Run sets the controls on frm_Main exactly as a user would, then presses Run.
'
'  TO INSTALL IN THE WORK MACRO (three small edits, then import this module):
'
'    1. frm_Main: the Run handler is Private. Add a public wrapper at the end of the form's code:
'
'           Public Sub RunIt()
'               btn_Run_Click
'           End Sub
'
'    2. module PolyAppPro: split Sub PolyAppPro() so the setup can run without showing the form.
'       Move everything above "frm_Main.Show" into
'
'           Public Sub PA_Init()
'               If Environ("username") = "Azar" Then bPDM = False Else bPDM = True
'               bDebug = True
'               Set swApp = Application.SldWorks
'               Set swFrame = swApp.Frame
'               Set swModel = swApp.ActiveDoc
'               Set fso = CreateObject("Scripting.FileSystemObject")
'               Set Shex = CreateObject("Shell.Application")
'               Set swSelMgr = swModel.SelectionManager
'               Set swFeatMgr = swModel.FeatureManager
'               If bPDM Then
'                   Set swPdmVault = New EdmVault5
'                   If Not swPdmVault.IsLoggedIn Then swPdmVault.LoginAuto VAULT_NAME, 0
'               End If
'               DocDate = Format(Date, "mmm-dd-yy")
'           End Sub
'
'       and make Sub PolyAppPro() call PA_Init before frm_Main.Show (the GetDocumentCount / ActiveDoc
'       guards stay in PolyAppPro(); PA_Run repeats them below).
'
'    3. Import this file (VBA editor: File > Import File) and save the .swp.
'
'  Then:  polyapp commands   lists every add-in command with the control it sets; any control name that
'  differs in the work macro is overridden in PolyApp.ini under [MacroOptions] without touching this file.
' ============================================================================
Option Explicit

Private Declare PtrSafe Function GetPrivateProfileStringA Lib "kernel32" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Function IniRead(ByVal path As String, ByVal section As String, ByVal key As String) As String
    Dim buf As String * 2048
    Dim n As Long
    n = GetPrivateProfileStringA(section, key, "", buf, Len(buf), path)
    IniRead = Left$(buf, n)
End Function

Public Function PA_RequestFile() As String
    PA_RequestFile = Environ("TEMP") & "\PolyApp.request.ini"
End Function

' The add-in's entry point.
Public Sub PA_Run()
    Dim req As String
    Dim cmd As String, opt1 As String, opt2 As String, tree As String
    req = PA_RequestFile
    cmd = IniRead(req, "Request", "Command")
    opt1 = IniRead(req, "Request", "Option")
    opt2 = IniRead(req, "Request", "Option2")
    tree = IniRead(req, "Request", "Tree")
    If Len(opt1) = 0 Then
        MsgBox "PolyApp: no request found in " & req, vbExclamation, "PolyApp"
        Exit Sub
    End If

    Set swApp = Application.SldWorks
    If swApp.GetDocumentCount() = 0 Then Exit Sub
    If swApp.ActiveDoc Is Nothing Then Exit Sub
    PA_Init                                  ' module PolyAppPro (edit 2 above)

    Load frm_Main                            ' UserForm_Initialize fills serial number / section from the active doc
    If Not PA_SetControl(opt1, cmd) Then Unload frm_Main: Exit Sub
    If Len(opt2) > 0 Then
        If Not PA_SetControl(opt2, cmd) Then Unload frm_Main: Exit Sub
    End If
    On Error Resume Next
    frm_Main.chk_Tree.value = (tree = "1")
    On Error GoTo 0

    swFrame.SetStatusBarText "PolyApp: " & cmd
    frm_Main.RunIt                           ' frm_Main (edit 1 above): btn_Run_Click
End Sub

' Set one option/check box by name; False (with a message) when the control does not exist in this macro.
Private Function PA_SetControl(ByVal ctrlName As String, ByVal cmd As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = frm_Main.Controls(ctrlName)
    On Error GoTo 0
    If ctrl Is Nothing Then
        MsgBox "PolyApp: frm_Main has no control named '" & ctrlName & "' (command " & cmd & ")." & vbLf & vbLf & _
               "Map it in PolyApp.ini under [MacroOptions]:  " & cmd & "=opt_...", vbExclamation, "PolyApp"
        PA_SetControl = False
        Exit Function
    End If
    If Not ctrl.Enabled Then
        MsgBox "PolyApp: '" & ctrlName & "' is disabled for this document (command " & cmd & ").", vbExclamation, "PolyApp"
        PA_SetControl = False
        Exit Function
    End If
    ctrl.value = True
    PA_SetControl = True
End Function
