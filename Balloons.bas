Attribute VB_Name = "Balloons"
Option Explicit

' Module-level app reference, set on first call.
Private swApp As SldWorks.SldWorks

' ─────────────────────────────────────────────────────────────────────────────
' GetUnballooned
' Returns a Component2 array for every BOM row on the current sheet that does
' not yet have a balloon.  Returns Empty if all rows are covered or no BOM
' exists.
'
' Fix applied: the original code passed strConfig = "" to GetComponents2 via
' GetFirstComponentInRow.  On most assemblies GetComponents2 returns nothing
' when given an empty config string.  We now resolve the drawing view's
' ReferencedConfiguration and pass that instead.
' ─────────────────────────────────────────────────────────────────────────────
Function GetUnballooned() As Variant
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw  As SldWorks.DrawingDoc
    Set swModel = swApp.ActiveDoc
    Set swDraw  = swModel

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet()

    Dim swBOMFeat As SldWorks.BomFeature
    Set swBOMFeat = GetBomFeature(swSheet)
    If swBOMFeat Is Nothing Then
        GetUnballooned = Empty
        Exit Function
    End If

    ' Guard: GetTableAnnotations may return Empty if the BOM has no view yet.
    Dim vTableAnns As Variant
    vTableAnns = swBOMFeat.GetTableAnnotations()
    If IsEmpty(vTableAnns) Then
        GetUnballooned = Empty
        Exit Function
    End If

    Dim swTableAnn As SldWorks.TableAnnotation
    Set swTableAnn = vTableAnns(0)

    ' Resolve the assembly configuration shown in the drawing view.
    ' This is what GetComponents2 requires; "" fails on most models.
    Dim strViewConfig As String
    strViewConfig = GetBomViewConfig(swDraw)

    Dim vAnn As Variant
    vAnn = GetBalloonVals(swModel)

    Dim vComps() As SldWorks.Component2
    Dim nCount   As Long

    Dim i As Long
    For i = 1 To swTableAnn.RowCount - 1
        Dim strItemNo As String
        strItemNo = swTableAnn.Text(i, 0)

        If Not IsInArray(strItemNo, vAnn) Then
            Dim swComp As SldWorks.Component2
            Set swComp = GetFirstComponentInRow(swTableAnn, i, strViewConfig)

            If Not swComp Is Nothing Then
                ReDim Preserve vComps(nCount)
                Set vComps(nCount) = swComp
                nCount = nCount + 1
            End If
        End If
    Next i

    If nCount = 0 Then
        GetUnballooned = Empty
    Else
        GetUnballooned = vComps
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' GetBomFeature
' Walks the drawing feature tree and returns the first BomFeature whose table
' annotation belongs to swSheet.  Returns Nothing if none is found.
' ─────────────────────────────────────────────────────────────────────────────
Function GetBomFeature(swSheet As SldWorks.Sheet) As SldWorks.BomFeature
    If swApp Is Nothing Then Set swApp = Application.SldWorks

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swApp.ActiveDoc

    Dim strSheetName As String
    strSheetName = swSheet.GetName()

    Dim swFeat As SldWorks.Feature
    Set swFeat = swDraw.FirstFeature()

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "BomFeat" Then
            Dim swBOM As SldWorks.BomFeature
            Set swBOM = swFeat.GetSpecificFeature2()

            If Not swBOM Is Nothing Then
                Dim vTAnns As Variant
                vTAnns = swBOM.GetTableAnnotations()

                If Not IsEmpty(vTAnns) Then
                    Dim swTA  As SldWorks.TableAnnotation
                    Set swTA = vTAnns(0)

                    Dim swAnnSheet As SldWorks.Sheet
                    Set swAnnSheet = swTA.GetSheet()

                    If Not swAnnSheet Is Nothing Then
                        If swAnnSheet.GetName() = strSheetName Then
                            Set GetBomFeature = swBOM
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature()
    Loop

    Set GetBomFeature = Nothing
End Function

' ─────────────────────────────────────────────────────────────────────────────
' GetBalloonVals
' Returns a String() of item-number text for every balloon annotation found
' across all views in the active drawing.  Returns an empty Array() if none.
' ─────────────────────────────────────────────────────────────────────────────
Function GetBalloonVals(swModel As SldWorks.ModelDoc2) As Variant
    ' swAnnotationType_e: swBalloon = 18, swStackedBalloon = 19
    Const SW_BALLOON        As Long = 18
    Const SW_STACKED_BALLOON As Long = 19

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel

    Dim results() As String
    Dim nCount    As Long

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView()   ' first entry is the sheet-format placeholder

    Do While Not swView Is Nothing
        Dim vAnns As Variant
        vAnns = swView.GetAnnotations()

        If Not IsEmpty(vAnns) Then
            Dim j As Long
            For j = 0 To UBound(vAnns)
                Dim swAnn As SldWorks.Annotation
                Set swAnn = vAnns(j)

                Dim annType As Long
                annType = swAnn.GetType()

                If annType = SW_BALLOON Or annType = SW_STACKED_BALLOON Then
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

    If nCount = 0 Then
        GetBalloonVals = Array()
    Else
        GetBalloonVals = results
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' GetFirstComponentInRow
' Returns the first Component2 for a given BOM table row.
'
' Fix: accepts the resolved strConfig from the caller.  If GetComponents2
' still returns nothing (e.g. indented BOM edge cases), falls back to "".
' ─────────────────────────────────────────────────────────────────────────────
Function GetFirstComponentInRow(swTableAnn As SldWorks.TableAnnotation, _
                                rowIndex   As Long, _
                                strConfig  As String) As SldWorks.Component2
    Dim vComps As Variant
    vComps = swTableAnn.GetComponents2(rowIndex, strConfig)

    If IsEmpty(vComps) And Len(strConfig) > 0 Then
        vComps = swTableAnn.GetComponents2(rowIndex, "")
    End If

    If Not IsEmpty(vComps) Then
        Set GetFirstComponentInRow = vComps(0)
    Else
        Set GetFirstComponentInRow = Nothing
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' GetBomViewConfig
' Returns the ReferencedConfiguration of the first real drawing view
' (skips the sheet-format placeholder that GetFirstView returns).
' ─────────────────────────────────────────────────────────────────────────────
Private Function GetBomViewConfig(swDraw As SldWorks.DrawingDoc) As String
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView()   ' sheet-format placeholder
    Set swView = swView.GetNextView()    ' first real model view

    If Not swView Is Nothing Then
        GetBomViewConfig = swView.ReferencedConfiguration
    Else
        GetBomViewConfig = ""
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' IsInArray
' Case-insensitive check: returns True if strVal exists in vArr.
' ─────────────────────────────────────────────────────────────────────────────
Function IsInArray(strVal As String, vArr As Variant) As Boolean
    IsInArray = False
    If IsEmpty(vArr) Then Exit Function

    On Error GoTo Done   ' guard against zero-length arrays
    Dim i As Long
    For i = LBound(vArr) To UBound(vArr)
        If StrComp(CStr(vArr(i)), strVal, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
Done:
End Function
