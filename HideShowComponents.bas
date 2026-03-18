Attribute VB_Name = "HideShowComponents"
Option Explicit

' Prints the names of all components on the hide/show list for the currently
' selected drawing view in the active drawing document.
Sub PrintHideShowComponents()
On Error GoTo Failure

    ' --- 1. Verify the active document is a drawing ---------------------------
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim swDoc As SldWorks.ModelDoc2
    Set swDoc = swApp.ActiveDoc
    If swDoc Is Nothing Then
        Debug.Print "No active document."
        GoTo Failure
    End If

    If swDoc.GetType() <> swDocDRAWING Then
        Debug.Print "Active document is not a drawing."
        GoTo Failure
    End If

    ' --- 2. Get the selected drawing view -------------------------------------
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swDoc.SelectionManager

    Dim swView As SldWorks.View
    Dim i As Integer
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelDRAWINGVIEWS Then
            Set swView = swSelMgr.GetSelectedObject6(i, -1)
            Exit For
        End If
    Next i

    If swView Is Nothing Then
        Debug.Print "No drawing view is selected."
        GoTo Failure
    End If

    Debug.Print "Drawing view: " & swView.Name

    ' --- 3. Print all components on the hide/show list ------------------------
    Dim vComps As Variant
    vComps = swView.GetHideShowComponents()

    If IsEmpty(vComps) Or IsNull(vComps) Then
        Debug.Print "  (no components on the hide/show list)"
        Exit Sub
    End If

    Dim comp As Variant
    For Each comp In vComps
        Debug.Print "  " & comp
    Next comp

    Exit Sub
Failure:
    Debug.Print "PrintHideShowComponents failed (Err " & Err.Number & "): " & Err.Description
End Sub
