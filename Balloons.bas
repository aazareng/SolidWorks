Attribute VB_Name = "Balloons"
Option Explicit

Function GetUnballooned() As Variant
    Dim swApp  As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Set swApp  = Application.SldWorks
    Set swDraw = swApp.ActiveDoc

    ' Find the first BOM feature on the drawing
    Dim swBOMFeat As SldWorks.BomFeature
    Dim swFeat    As SldWorks.Feature
    Set swFeat = swDraw.FirstFeature()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "BomFeat" Then
            Set swBOMFeat = swFeat.GetSpecificFeature2()
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature()
    Loop

    If swBOMFeat Is Nothing Then GetUnballooned = Empty: Exit Function

    Dim vTAnns As Variant
    vTAnns = swBOMFeat.GetTableAnnotations()
    If IsEmpty(vTAnns) Then GetUnballooned = Empty: Exit Function

    Dim swTableAnn As SldWorks.TableAnnotation
    Set swTableAnn = vTAnns(0)

    ' ReferencedConfiguration is required by GetComponents2; "" fails on most assemblies
    Dim swView    As SldWorks.View
    Dim strConfig As String
    Set swView = swDraw.GetFirstView().GetNextView()  ' skip sheet-format placeholder
    If Not swView Is Nothing Then strConfig = swView.ReferencedConfiguration

    Dim vAnn As Variant
    vAnn = GetBalloonVals(swDraw)

    Dim vComps() As SldWorks.Component2
    Dim nCount   As Long
    Dim i        As Long
    For i = 1 To swTableAnn.RowCount - 1
        If Not IsInArray(swTableAnn.Text(i, 0), vAnn) Then
            Dim vRowComps As Variant
            vRowComps = swTableAnn.GetComponents2(i, strConfig)
            If IsEmpty(vRowComps) Then vRowComps = swTableAnn.GetComponents2(i, "")
            If Not IsEmpty(vRowComps) Then
                ReDim Preserve vComps(nCount)
                Set vComps(nCount) = vRowComps(0)
                nCount = nCount + 1
            End If
        End If
    Next i

    If nCount = 0 Then GetUnballooned = Empty Else GetUnballooned = vComps
End Function

' Returns a String array of item-number text for every balloon in the drawing
Function GetBalloonVals(swDraw As SldWorks.DrawingDoc) As Variant
    Const SW_BALLOON         As Long = 18
    Const SW_STACKED_BALLOON As Long = 19

    Dim results() As String
    Dim nCount    As Long
    Dim swView    As SldWorks.View
    Set swView = swDraw.GetFirstView()

    Do While Not swView Is Nothing
        Dim vAnns As Variant
        vAnns = swView.GetAnnotations()
        If Not IsEmpty(vAnns) Then
            Dim j As Long
            For j = 0 To UBound(vAnns)
                Dim swAnn As SldWorks.Annotation
                Set swAnn = vAnns(j)
                If swAnn.GetType() = SW_BALLOON Or swAnn.GetType() = SW_STACKED_BALLOON Then
                    Dim swNote As SldWorks.Note
                    Set swNote = swAnn.GetSpecificAnnotation()
                    If Not swNote Is Nothing Then
                        Dim strText As String
                        strText = Trim(swNote.GetText())
                        If Len(strText) > 0 Then
                            ReDim Preserve results(nCount)
                            results(nCount) = strText
                            nCount = nCount + 1
                        End If
                    End If
                End If
            Next j
        End If
        Set swView = swView.GetNextView()
    Loop

    If nCount = 0 Then GetBalloonVals = Array() Else GetBalloonVals = results
End Function

' Case-insensitive check: True if strVal exists in vArr
Function IsInArray(strVal As String, vArr As Variant) As Boolean
    IsInArray = False
    If IsEmpty(vArr) Then Exit Function
    On Error GoTo Done
    Dim i As Long
    For i = LBound(vArr) To UBound(vArr)
        If StrComp(CStr(vArr(i)), strVal, vbTextCompare) = 0 Then
            IsInArray = True: Exit Function
        End If
    Next i
Done:
End Function
