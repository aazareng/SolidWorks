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

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swDoc

    ' Activate the view — required so component selections resolve in view context
    swDraw.ActivateView swView.Name

    Dim vComps As Variant
    vComps = swView.GetHiddenComponents(False)
    If IsEmpty(vComps) Then
        Debug.Print "  No hidden components found."
        swDraw.ActivateView ""
        Exit Sub
    End If

    swDoc.ClearSelection2 True

    Dim vComp  As Variant
    Dim swComp As SldWorks.Component2
    Dim bFirst As Boolean
    bFirst = True
    For Each vComp In vComps
        Set swComp = vComp
        ' Select4 (not Select) is required for drawing-context components
        swComp.Select4 Not bFirst, Nothing, False
        bFirst = False
        Debug.Print "  Selected: " & swComp.Name2
    Next

    ' ShowComponent2 acts on all currently selected components
    swDoc.ShowComponent2

    swDraw.ActivateView ""
    swDoc.GraphicsRedraw2
    Exit Sub

Failure:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub
