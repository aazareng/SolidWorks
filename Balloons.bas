Attribute VB_Name = "Balloons"
Option Explicit

Function GetUnballooned(swView As SldWorks.View) As Variant
    On Error GoTo Failure
    Dim swApp  As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Set swApp  = Application.SldWorks
    Set swDraw = swApp.ActiveDoc

    Dim swSheet As SldWorks.sheet
    Set swSheet = swDraw.GetCurrentSheet()

    Dim swBOMFeat As SldWorks.BomFeature
    Set swBOMFeat = GetBomFeature(swSheet)

    If swBOMFeat Is Nothing Then GetUnballooned = Empty: Exit Function

    Dim vTAnns As Variant
    vTAnns = swBOMFeat.GetTableAnnotations()
    If IsEmpty(vTAnns) Then GetUnballooned = Empty: Exit Function

    Dim swTableAnn As SldWorks.TableAnnotation
    Set swTableAnn = vTAnns(0)

    Dim vAnn As Variant
    vAnn = GetBalloonVals(swDraw)

    Dim vComps As Variant
    ReDim vComps(swTableAnn.RowCount - 1)
    Dim nCount    As Long
    Dim i         As Long
    Dim swComp    As SldWorks.Component2
    Dim strConfig As String
    For i = 1 To swTableAnn.RowCount - 1
        If Not swComp Is Nothing Then strConfig = swComp.ReferencedConfiguration Else strConfig = ""
        Set swComp = GetFirstComponentInRow(swTableAnn, i, strConfig)
        If Not IsInArray(swTableAnn.Text(i, 0), vAnn) Then
            If Not swComp Is Nothing Then
                Set vComps(nCount) = swComp
                nCount = nCount + 1
            End If
        End If
    Next i
    ReDim Preserve vComps(nCount - 1)

    If nCount = 0 Then GoTo Failure
Success:
    GetUnballooned = vComps
    Exit Function
Failure:
    GetUnballooned = Empty
    Stop
End Function

Function GetFirstComponentInRow(swTableAnn As SldWorks.TableAnnotation, nRow As Long, strConfig As String) As SldWorks.Component2
    Dim vComps As Variant
    vComps = swTableAnn.GetComponents2(nRow, strConfig)
    If IsEmpty(vComps) Then vComps = swTableAnn.GetComponents2(nRow, "")
    If Not IsEmpty(vComps) Then Set GetFirstComponentInRow = vComps(0)
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
