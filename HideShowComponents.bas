Attribute VB_Name = "HideShowComponents"
Option Explicit

' For the selected drawing view, finds every component that has been explicitly
' hidden (visible in model, hidden in view) and shows it, removing that override.
' Also reports any components that are explicitly shown (hidden in model, shown in view).
Sub ShowExplicitlyHiddenComponents()
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

    ' --- 3. Get the referenced assembly and its root component ----------------
    Dim swRefDoc As SldWorks.ModelDoc2
    Set swRefDoc = swView.ReferencedDocument
    If swRefDoc Is Nothing Then
        Debug.Print "View has no referenced document."
        GoTo Failure
    End If

    If swRefDoc.GetType() <> swDocASSEMBLY Then
        Debug.Print "Referenced document is not an assembly."
        GoTo Failure
    End If

    Dim swConfig As SldWorks.Configuration
    Set swConfig = swRefDoc.GetConfigurationByName(swView.ReferencedConfiguration)
    If swConfig Is Nothing Then
        Debug.Print "Could not get referenced configuration: " & swView.ReferencedConfiguration
        GoTo Failure
    End If

    Dim swRoot As SldWorks.Component2
    Set swRoot = swConfig.GetRootComponent3(True)
    If swRoot Is Nothing Then
        Debug.Print "Could not get root component."
        GoTo Failure
    End If

    ' --- 4. Traverse the tree and show any explicitly hidden components -------
    TraverseComponents swRoot, swView, swDoc

    swDoc.GraphicsRedraw2

    Exit Sub
Failure:
    Debug.Print "ShowExplicitlyHiddenComponents failed (Err " & Err.Number & "): " & Err.Description
End Sub

' Recursively walks the component tree.  For each component that is explicitly
' hidden in the view (visible in the model config but hidden in the view) it
' removes the override by calling HideShowComponent(..., True).
Private Sub TraverseComponents(swComp As SldWorks.Component2, _
                               swView As SldWorks.View, _
                               swDoc As SldWorks.ModelDoc2)
    Dim vChildren As Variant
    vChildren = swComp.GetChildren()
    If IsEmpty(vChildren) Then Exit Sub

    Dim child As Variant
    For Each child In vChildren
        Dim swChild As SldWorks.Component2
        Set swChild = child

        Dim swDrawComp As SldWorks.DrawingComponent
        Set swDrawComp = swChild.GetDrawingComponent(swView)

        If Not swDrawComp Is Nothing Then
            Dim bVisInView  As Boolean
            Dim bHidInModel As Boolean
            bVisInView  = swDrawComp.Visible
            bHidInModel = swChild.IsHidden(False)

            If Not bVisInView And Not bHidInModel Then
                ' Explicitly hidden: visible in model, hidden in view — show it
                Dim bResult As Boolean
                bResult = swView.HideShowComponent(swChild, True)
                Debug.Print "  Shown: " & swChild.Name2 & _
                            IIf(bResult, "", "  (HideShowComponent returned False)")

            ElseIf bVisInView And bHidInModel Then
                ' Explicitly shown: hidden in model, visible in view — leave it, just report
                Debug.Print "  [explicitly shown, left as-is]: " & swChild.Name2
            End If
        End If

        TraverseComponents swChild, swView, swDoc
    Next child
End Sub
