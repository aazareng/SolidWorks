Attribute VB_Name = "DrillUp"
Option Explicit

' =====================================================================
' DrillUp — assembly circular-edge drill-up
'
' Paste DrillUpAssy over the existing one in ProTools.swp and add the
' helper functions below it (FindOppositeCircularEdge and friends).
' DrillUp itself is unchanged and included for completeness.
'
' Assumes the module already has:
'   Dim swApp As SldWorks.SldWorks
'   Sub SelectObject(...), Sub LogDrillUp(...), GetNextHigherComponent,
'   FindFolderInComponent, SelectParentFolder, GetLastFeature, GetPreviousFeature
' =====================================================================

Private Sub DrillUpAssy(ByVal swModel As SldWorks.ModelDoc2)
On Error GoTo Failure

    If swModel Is Nothing Then GoTo Failure

        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager

        Dim selCount As Long
        selCount = swSelMgr.GetSelectedObjectCount2(-1)

        If selCount = 0 Then
            Exit Sub
        End If

        Dim latestIdx As Long
        latestIdx = selCount

        Dim swObj As Object
        Set swObj = swSelMgr.GetSelectedObject6(latestIdx, -1)

        If swObj Is Nothing Then
            Exit Sub
        End If

        ' Log the selected object's persist ID
        Dim vId As Variant
        vId = swModel.Extension.GetPersistReference3(swObj)
        Dim selType As Long
        selType = swSelMgr.GetSelectedObjectType3(latestIdx, -1)
        LogDrillUp swModel, "Selected", selType, vId

        ' --- Find the parent ---
        Dim swParentObj As Object
        Set swParentObj = Nothing
        Dim parentType As Long
        parentType = -1

        ' --- FACE / EDGE: parent is the part component ---
        ' Circular EDGE: parent is the next concentric circular edge along the
        ' hole axis, stepping away from the hole wall (e.g. the matching hole in
        ' the plate stacked on top). Falls back to the component.
        If selType = swSelFACES Or selType = swSelEDGES Then

            Dim swFaceComp As SldWorks.Component2
            Set swFaceComp = swSelMgr.GetSelectedObjectsComponent3(latestIdx, -1)

            If selType = swSelEDGES Then
                Dim swOppEdge As SldWorks.Edge
                Set swOppEdge = FindOppositeCircularEdge(swModel, swObj, swFaceComp)
                If Not swOppEdge Is Nothing Then
                    Set swParentObj = swOppEdge
                    parentType = swSelEDGES
                End If
            End If

            If swParentObj Is Nothing Then
                If Not swFaceComp Is Nothing Then
                    Set swParentObj = swFaceComp
                    parentType = swSelCOMPONENTS
                End If
            End If

        ' --- COMPONENT or FEATURE ---
        Else

            Dim parentFound As Boolean
            parentFound = False

            Dim searchName As String
            searchName = ""

            Dim swSelComp As SldWorks.Component2

            If selType = swSelCOMPONENTS Then
                Set swSelComp = swObj
                Dim compParts() As String
                compParts = Split(swSelComp.name2, "/")
                searchName = compParts(UBound(compParts))
            Else
                Dim swFeat As SldWorks.feature
                On Error Resume Next
                Set swFeat = swObj
                On Error GoTo 0
                If Not swFeat Is Nothing Then
                    searchName = swFeat.name
                End If
            End If

            ' Try to find a parent folder
            If Len(searchName) > 0 Then
                Dim swParentFolder As SldWorks.feature

                If selType = swSelCOMPONENTS And swModel.GetType = swDocASSEMBLY Then
                    Dim swAssy As SldWorks.AssemblyDoc
                    Set swAssy = swModel

                    Dim swParentComp As SldWorks.Component2
                    Set swParentComp = GetNextHigherComponent(swAssy, swSelComp)

                    If Not swParentComp Is Nothing Then
                        Set swParentFolder = FindFolderInComponent(swParentComp, searchName)
                    Else
                        Set swParentFolder = SelectParentFolder(swModel, searchName)
                    End If
                Else
                    Dim swFeatComp As SldWorks.Component2
                    Set swFeatComp = swSelMgr.GetSelectedObjectsComponent3(latestIdx, -1)

                    If Not swFeatComp Is Nothing Then
                        Set swParentFolder = FindFolderInComponent(swFeatComp, searchName)
                    Else
                        Set swParentFolder = SelectParentFolder(swModel, searchName)
                    End If
                End If

                If Not swParentFolder Is Nothing Then
                    Set swParentObj = swParentFolder
                    parentType = swSelBODYFEATURES
                    parentFound = True
                End If
            End If

            ' No folder — walk up component hierarchy
            If Not parentFound Then
                If selType = swSelCOMPONENTS Then
                    If swModel.GetType = swDocASSEMBLY Then
                        Dim swAssy2 As SldWorks.AssemblyDoc
                        Set swAssy2 = swModel

                        Dim swParentComp2 As SldWorks.Component2
                        Set swParentComp2 = GetNextHigherComponent(swAssy2, swSelComp)

                        If Not swParentComp2 Is Nothing Then
                            Set swParentObj = swParentComp2
                            parentType = swSelCOMPONENTS
                            parentFound = True
                        End If
                    End If
                Else
                    Dim swFeatComp2 As SldWorks.Component2
                    Set swFeatComp2 = swSelMgr.GetSelectedObjectsComponent3(latestIdx, -1)
                    If Not swFeatComp2 Is Nothing Then
                        Set swParentObj = swFeatComp2
                        parentType = swSelCOMPONENTS
                        parentFound = True
                    End If
                End If
            End If

        End If

        ' --- If no parent was found, item is directly under the TLA — do nothing ---
        If swParentObj Is Nothing Then
            LogDrillUp swModel, "AtTLA", selType, vId
            Debug.Print "Item resides directly below the TLA. Selection unchanged."
            Exit Sub
        End If

        ' Log the parent's persist ID
        Dim vParentId As Variant
        vParentId = swModel.Extension.GetPersistReference3(swParentObj)
        LogDrillUp swModel, "Parent", parentType, vParentId

        ' --- Store persist IDs of all PREVIOUS selections (everything except the latest) ---
        Dim prevCount As Long
        prevCount = selCount - 1

        Dim prevIds() As Variant
        Dim prevTypes() As Long
        Dim prevMarks() As Long

        If prevCount > 0 Then
            ReDim prevIds(0 To prevCount - 1)
            ReDim prevTypes(0 To prevCount - 1)
            ReDim prevMarks(0 To prevCount - 1)

            Dim j As Long
            For j = 1 To prevCount
                prevIds(j - 1) = swModel.Extension.GetPersistReference3(swSelMgr.GetSelectedObject6(j, -1))
                prevTypes(j - 1) = swSelMgr.GetSelectedObjectType3(j, -1)
                prevMarks(j - 1) = swSelMgr.GetSelectedObjectMark(j)
            Next j
        End If

        ' --- Now rebuild the selection: previous items + parent ---
        swModel.ClearSelection2 True

        ' Re-select all previous items using persist references
        If prevCount > 0 Then
            Dim k As Long
            For k = 0 To prevCount - 1
                Dim errCode As Long
                Dim swPrevObj As Object
                Set swPrevObj = swModel.Extension.GetObjectByPersistReference3(prevIds(k), errCode)
                If Not swPrevObj Is Nothing Then
                    Dim swSelData As SldWorks.SelectData
                    Set swSelData = swSelMgr.CreateSelectData
                    swSelData.Mark = prevMarks(k)
                    SelectObject swModel, swPrevObj, prevTypes(k), True, swSelData
                End If
            Next k
        End If

        ' Append the parent to the selection
        Dim swParentSelData As SldWorks.SelectData
        Set swParentSelData = swSelMgr.CreateSelectData
        swParentSelData.Mark = 0
        SelectObject swModel, swParentObj, parentType, True, swParentSelData

Success::
Failure::

End Sub

' =====================================================================
' Circular-edge drill-up helpers
' =====================================================================

' From circular edge A: step along the hole axis, away from the hole wall
' (the non-planar face adjacent to A), and return the first concentric
' circular edge encountered on any resolved, visible part in the assembly.
' Any diameter qualifies. Returns Nothing when A is not a circle or when no
' concentric circle lies ahead.
'
' Entity geometry of a component is in that component's own space, so
' every centre/axis is moved into assembly space with Component2.Transform2
' before comparing.
Private Function FindOppositeCircularEdge(ByVal swModel As SldWorks.ModelDoc2, _
                                          ByVal swObj As Object, _
                                          ByVal swEdgeComp As SldWorks.Component2) As SldWorks.Edge
    On Error GoTo Failure
    Set FindOppositeCircularEdge = Nothing

    If swModel Is Nothing Then Exit Function
    If swModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim swEdge As SldWorks.Edge
    On Error Resume Next
    Set swEdge = swObj
    On Error GoTo Failure
    If swEdge Is Nothing Then Exit Function

    Dim swCurve As SldWorks.Curve
    Set swCurve = swEdge.GetCurve
    If swCurve Is Nothing Then Exit Function
    ' SOLIDWORKS returns some booleans as 1 rather than -1, so "If Not x" is
    ' wrong for API results: Not 1 = -2 = True. Compare against False instead.
    If swCurve.IsCircle = False Then Exit Function

    Dim swMath As SldWorks.MathUtility
    Set swMath = swApp.GetMathUtility

    Dim swXform As SldWorks.MathTransform
    Set swXform = Nothing
    If Not swEdgeComp Is Nothing Then Set swXform = swEdgeComp.Transform2

    ' Circle A: centre and unit axis in assembly space
    Dim cA(2) As Double, axA(2) As Double
    If Not CircleInAssemblySpace(swMath, swCurve, swXform, cA, axA) Then Exit Function

    ' Search direction: +/- axis, pointing away from the hole wall
    Dim dirA(2) As Double
    If Not GetAwayFromHoleDirection(swMath, swEdge, swXform, cA, axA, dirA) Then Exit Function

    Const TOL_LIN As Double = 0.00001     ' 10 um: concentricity + minimum step
    Const TOL_ANG As Double = 0.0001      ' |cos| deviation allowed for parallel axes
    Const BOX_PAD As Double = 0.001       ' 1 mm inflation of bounding boxes in the pre-filter

    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)  ' all levels
    If IsEmpty(vComps) Then Exit Function

    Dim bestDist As Double
    bestDist = 1E+300
    Dim swBest As SldWorks.Edge
    Set swBest = Nothing

    Dim i As Long, b As Long, e As Long
    Dim swComp As SldWorks.Component2
    Dim swCompXform As SldWorks.MathTransform
    Dim vBox As Variant, vBodies As Variant, vBodyInfo As Variant, vEdges As Variant
    Dim swBody As SldWorks.Body2
    Dim swCand As SldWorks.Edge
    Dim swCandCurve As SldWorks.Curve
    Dim cC(2) As Double, axC(2) As Double, v(2) As Double
    Dim t As Double, perp As Double
    Dim sameComp As Boolean

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        If swComp Is Nothing Then GoTo NextComp
        If swComp.GetSuppression2 = swComponentSuppressed Then GoTo NextComp
        If swComp.Visible = swComponentHidden Then GoTo NextComp
        If swComp.IGetChildrenCount > 0 Then GoTo NextComp   ' sub-assembly: its parts are listed on their own

        sameComp = False
        If Not swEdgeComp Is Nothing Then
            If swComp.Name2 = swEdgeComp.Name2 Then sameComp = True
        End If

        ' Cheap pre-filter: skip components whose box the search ray never enters
        vBox = Empty
        On Error Resume Next
        vBox = swComp.GetBox(False, False)
        On Error GoTo Failure
        If Not IsEmpty(vBox) Then
            If Not RayHitsBox(cA, dirA, vBox, BOX_PAD) Then GoTo NextComp
        End If

        Set swCompXform = swComp.Transform2
        vBodies = Empty
        vBodies = swComp.GetBodies3(swSolidBody, vBodyInfo)
        If IsEmpty(vBodies) Then GoTo NextComp

        For b = 0 To UBound(vBodies)
            Set swBody = vBodies(b)
            If swBody Is Nothing Then GoTo NextBody
            vEdges = swBody.GetEdges
            If IsEmpty(vEdges) Then GoTo NextBody

            For e = 0 To UBound(vEdges)
                Set swCand = vEdges(e)
                Set swCandCurve = swCand.GetCurve
                If swCandCurve Is Nothing Then GoTo NextEdge
                If swCandCurve.IsCircle = False Then GoTo NextEdge
                If Not CircleInAssemblySpace(swMath, swCandCurve, swCompXform, cC, axC) Then GoTo NextEdge

                ' Parallel axes
                If Abs(Abs(Dot3(axC, axA)) - 1#) > TOL_ANG Then GoTo NextEdge

                ' Ahead of A along the search direction. A coincident circle (t = 0)
                ' counts only when it belongs to another component: that is the
                ' touching plate. On A's own component t = 0 is A itself.
                v(0) = cC(0) - cA(0): v(1) = cC(1) - cA(1): v(2) = cC(2) - cA(2)
                t = Dot3(v, dirA)
                If t < -TOL_LIN Then GoTo NextEdge
                If t <= TOL_LIN And sameComp Then GoTo NextEdge

                ' Centre lies on A's axis (concentric)
                perp = Sqr(Abs(Dot3(v, v) - t * t))
                If perp > TOL_LIN Then GoTo NextEdge

                ' Keep the nearest one = first circular edge on the next face
                If t < bestDist Then
                    bestDist = t
                    Set swBest = swCand
                End If
NextEdge:
            Next e
NextBody:
        Next b
NextComp:
    Next i

    Set FindOppositeCircularEdge = swBest
    Exit Function

Failure:
    Set FindOppositeCircularEdge = Nothing
End Function

' Centre and unit axis of a circular curve, moved into assembly space when a
' transform is supplied.
Private Function CircleInAssemblySpace(ByVal swMath As SldWorks.MathUtility, _
                                       ByVal swCurve As SldWorks.Curve, _
                                       ByVal swXform As SldWorks.MathTransform, _
                                       ByRef cOut() As Double, _
                                       ByRef axOut() As Double) As Boolean
    CircleInAssemblySpace = False

    Dim vP As Variant
    vP = swCurve.CircleParams            ' cx cy cz  ax ay az  r
    If IsEmpty(vP) Then Exit Function

    Dim p(2) As Double, a(2) As Double
    p(0) = vP(0): p(1) = vP(1): p(2) = vP(2)
    a(0) = vP(3): a(1) = vP(4): a(2) = vP(5)

    If Not swXform Is Nothing Then
        Dim swPt As SldWorks.MathPoint
        Dim swVec As SldWorks.MathVector
        Set swPt = swMath.CreatePoint((p))
        Set swVec = swMath.CreateVector((a))
        Set swPt = swPt.MultiplyTransform(swXform)
        Set swVec = swVec.MultiplyTransform(swXform)
        vP = swPt.ArrayData
        p(0) = vP(0): p(1) = vP(1): p(2) = vP(2)
        vP = swVec.ArrayData
        a(0) = vP(0): a(1) = vP(1): a(2) = vP(2)
    End If

    If Not Normalize3(a) Then Exit Function

    cOut(0) = p(0): cOut(1) = p(1): cOut(2) = p(2)
    axOut(0) = a(0): axOut(1) = a(1): axOut(2) = a(2)
    CircleInAssemblySpace = True
End Function

' Search direction for circular edge A: its axis, signed so that it points
' away from the hole wall (non-planar face adjacent to A). Uses the wall's
' bounding-box centre -> edge centre, which also handles bosses and chamfered
' or countersunk holes. Falls back to the planar face's outward normal when
' A has no non-planar neighbour.
Private Function GetAwayFromHoleDirection(ByVal swMath As SldWorks.MathUtility, _
                                          ByVal swEdge As SldWorks.Edge, _
                                          ByVal swXform As SldWorks.MathTransform, _
                                          ByRef cA() As Double, _
                                          ByRef axA() As Double, _
                                          ByRef dirOut() As Double) As Boolean
    GetAwayFromHoleDirection = False

    Dim vFaces As Variant
    vFaces = swEdge.GetTwoAdjacentFaces2
    If IsEmpty(vFaces) Then Exit Function

    Dim swWall As SldWorks.Face2
    Dim swPlane As SldWorks.Face2
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface
    Dim k As Long
    For k = 0 To UBound(vFaces)
        Set swFace = vFaces(k)
        If swFace Is Nothing Then GoTo NextFace
        Set swSurf = swFace.GetSurface
        If swSurf Is Nothing Then GoTo NextFace
        If swSurf.IsPlane Then
            If swPlane Is Nothing Then Set swPlane = swFace
        ElseIf swSurf.IsCylinder Then
            Set swWall = swFace               ' prefer a cylinder
        ElseIf swWall Is Nothing Then
            Set swWall = swFace               ' cone / torus / other
        End If
NextFace:
    Next k

    Dim d(2) As Double
    Dim vTmp As Variant

    If Not swWall Is Nothing Then
        vTmp = swWall.GetBox
        If IsEmpty(vTmp) Then Exit Function
        Dim m(2) As Double
        m(0) = (vTmp(0) + vTmp(3)) / 2#
        m(1) = (vTmp(1) + vTmp(4)) / 2#
        m(2) = (vTmp(2) + vTmp(5)) / 2#
        If Not swXform Is Nothing Then
            Dim swPt As SldWorks.MathPoint
            Set swPt = swMath.CreatePoint((m))
            Set swPt = swPt.MultiplyTransform(swXform)
            vTmp = swPt.ArrayData
            m(0) = vTmp(0): m(1) = vTmp(1): m(2) = vTmp(2)
        End If
        d(0) = cA(0) - m(0): d(1) = cA(1) - m(1): d(2) = cA(2) - m(2)
    ElseIf Not swPlane Is Nothing Then
        vTmp = swPlane.Normal                 ' outward normal, sense already applied
        d(0) = vTmp(0): d(1) = vTmp(1): d(2) = vTmp(2)
        If Not swXform Is Nothing Then
            Dim swVec As SldWorks.MathVector
            Set swVec = swMath.CreateVector((d))
            Set swVec = swVec.MultiplyTransform(swXform)
            vTmp = swVec.ArrayData
            d(0) = vTmp(0): d(1) = vTmp(1): d(2) = vTmp(2)
        End If
    Else
        Exit Function
    End If

    ' Snap to the axis so the search line is exactly the hole axis
    Dim s As Double
    s = Dot3(d, axA)
    If Abs(s) < 0.000000001 Then Exit Function
    If s < 0 Then
        dirOut(0) = -axA(0): dirOut(1) = -axA(1): dirOut(2) = -axA(2)
    Else
        dirOut(0) = axA(0): dirOut(1) = axA(1): dirOut(2) = axA(2)
    End If
    GetAwayFromHoleDirection = True
End Function

' Slab test: does the half-line o + t*d (t >= 0) pass through the axis-aligned
' box (xmin ymin zmin xmax ymax zmax), inflated by pad?
Private Function RayHitsBox(ByRef o() As Double, ByRef d() As Double, _
                            ByVal vBox As Variant, ByVal pad As Double) As Boolean
    RayHitsBox = False

    Dim tmin As Double, tmax As Double
    Dim t1 As Double, t2 As Double, tmp As Double
    Dim lo As Double, hi As Double
    tmin = 0#
    tmax = 1E+300

    Dim k As Long
    For k = 0 To 2
        lo = vBox(k): hi = vBox(k + 3)
        If lo > hi Then tmp = lo: lo = hi: hi = tmp
        lo = lo - pad: hi = hi + pad

        If Abs(d(k)) < 0.000000000001 Then
            If o(k) < lo Or o(k) > hi Then Exit Function
        Else
            t1 = (lo - o(k)) / d(k)
            t2 = (hi - o(k)) / d(k)
            If t1 > t2 Then tmp = t1: t1 = t2: t2 = tmp
            If t1 > tmin Then tmin = t1
            If t2 < tmax Then tmax = t2
            If tmin > tmax Then Exit Function
        End If
    Next k

    RayHitsBox = True
End Function

Private Function Dot3(ByRef a() As Double, ByRef b() As Double) As Double
    Dot3 = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Private Function Normalize3(ByRef a() As Double) As Boolean
    Dim n As Double
    n = Sqr(a(0) * a(0) + a(1) * a(1) + a(2) * a(2))
    If n < 0.000000000001 Then
        Normalize3 = False
        Exit Function
    End If
    a(0) = a(0) / n: a(1) = a(1) / n: a(2) = a(2) / n
    Normalize3 = True
End Function

' =====================================================================
' Entry point (unchanged)
' =====================================================================
Sub DrillUp()
    On Error GoTo Failure
    Dim bRet As Boolean
    Dim swFeat As SldWorks.feature
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then GoTo Failure

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager

    Dim selCount As Long
    selCount = swSelMgr.GetSelectedObjectCount2(-1)

    Dim isEditingSketch As Boolean
    isEditingSketch = Not (swModel.SketchManager.ActiveSketch Is Nothing)

    Dim selType As Long
    Dim swObj As Object

    ' ==========================================================
    ' DOC TYPE
    ' ==========================================================
    Select Case swModel.GetType
        Case swDocASSEMBLY
                DrillUpAssy swModel
                GoTo Success
        Case swDocPART
        Case swDocDRAWING
        Case Else
            GoTo Failure
    End Select

    ' ==========================================================
    ' GET LAST SELECTED OBJECT
    ' ==========================================================
    selType = swSelMgr.GetSelectedObjectType3(selCount, -1)
    Set swObj = swSelMgr.GetSelectedObject6(selCount, -1)

    Select Case selCount
        Case 0 'Nothing Selected (shouldn't ever reach)
             If isEditingSketch Then

             Else
                Set swFeat = GetLastFeature(swModel) 'swModel.Extension.GetLastFeatureAdded
                Set swFeat = GetPreviousFeature(swModel, swFeat)
                swFeat.Select (False)
                GoTo Success
            End If
        Case 1 'One thing selected
            If isEditingSketch Then
                Select Case selType
                    Case swSelEDGES, swSelFACES
                        'swModel.SelectMidpoint
                        swModel.SelectTangency
                        GoTo Success
                    Case swSelSKETCHSEGS, swSelEXTSKETCHSEGS
                        Dim swSeg As SldWorks.SketchSegment
                        Set swSeg = swObj
                        If swSeg Is Nothing Then GoTo Failure
                        bRet = swSeg.SelectChain(True, Nothing)
                        Select Case swSeg.GetType
                            Case swSketchLine
                            Case swSketchArc
                            Case Else
                                GoTo Success
                        End Select
                    Case Else
                        Debug.Print "Other selected type in sketch: " & selType
                        GoTo Success
                End Select
            Else
                Select Case selType
                    Case swSelFACES, swSelEDGES
                         swModel.SelectTangency
                         GoTo Success
                    Case Else
                        Set swFeat = swObj
                        Set swFeat = GetPreviousFeature(swModel, swFeat)
                        swFeat.Select (False)
                        GoTo Success
                End Select
            End If
        Case Else 'Many things selected

    End Select

    GoTo Success

Success::
    Exit Sub

Failure::

End Sub
