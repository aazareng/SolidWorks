Attribute VB_Name = "HideShowComponents"
Option Explicit

' For the selected drawing view, shows all components that have been explicitly
' hidden in that view (right-click > Show/Hide > Hide Component).
Sub ShowExplicitlyHiddenComponents()
On Error GoTo Failure

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim swDoc As SldWorks.ModelDoc2
    Set swDoc = swApp.ActiveDoc
    If swDoc Is Nothing Then Debug.Print "No active document." : Exit Sub

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swDoc.SelectionManager

    Dim swView As SldWorks.View
    Set swView = swSelMgr.GetSelectedObject6(1, -1)
    If swView Is Nothing Then Debug.Print "No drawing view selected." : Exit Sub

    Debug.Print "Drawing view: " & swView.Name

    Dim vComps As Variant
    vComps = swView.GetHiddenComponents
    If IsEmpty(vComps) Then
        Debug.Print "  No hidden components found."
        Exit Sub
    End If

    Dim vComp   As Variant
    Dim swComp  As SldWorks.Component2
    Dim bResult As Boolean
    For Each vComp In vComps
        Set swComp = vComp
        bResult = swView.HideShowComponent(swComp, True)
        Debug.Print "  Shown: " & swComp.Name2 & IIf(bResult, "", "  (returned False)")
    Next

    swDoc.GraphicsRedraw2
    Exit Sub

Failure:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub
