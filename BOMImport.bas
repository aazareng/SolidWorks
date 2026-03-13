Attribute VB_Name = "BOMImport"
Option Compare Database
Option Explicit

Public Function ImportBOMToTable(strFilePath As String, _
                                  strConfig As String, _
                                  strBomName As String, _
                                  strTableName As String) As Boolean
On Error GoTo Failure
    Dim colMap As New Scripting.Dictionary
    colMap.Add "T11", "QTY"
    colMap.Add "T1", "PART_NUMBER"
    colMap.Add "T21", "CONFIGURATION"
    colMap.Add "T6", "STATE"
    colMap.Add "Revision", "REVISION"
    colMap.Add "Description", "DESCRIPTION"
    colMap.Add "Department", "DEPARTMENT"
    colMap.Add "Material", "MATERIAL"
    colMap.Add "Gauge", "GAUGE"
    colMap.Add "Length", "LENGTH"
    colMap.Add "Width", "WIDTH"
    colMap.Add "StockedPartNumber", "STOCKEDPARTNUMBER"
    colMap.Add "PartDestination", "PARTDESTINATION"
    colMap.Add "Weight", "WEIGHT"
    colMap.Add "SpecTyp", "SPECTYP"
    colMap.Add "SpecVal", "SPECVAL"
    ' --- 1. Vault login -------------------------------------------------------
    Dim localVault As New EdmLib.EdmVault5
    localVault.LoginAuto VAULT_NAME, 0
    If Not localVault.IsLoggedIn Then
        MsgBox "Failed to log in to vault: " & VAULT_NAME, vbCritical
        Exit Function
    End If
    ' --- 2. Resolve file ------------------------------------------------------
    Dim localFile As EdmLib.IEdmFile7
    Set localFile = localVault.GetFileFromPath(strFilePath)
    If localFile Is Nothing Then
        MsgBox "File not found in vault:" & vbCrLf & strFilePath, vbCritical
        Exit Function
    End If
    ' --- 3. Get computed BOM view ---------------------------------------------
    Dim localBomView As EdmLib.IEdmBomView3
    Set localBomView = localFile.GetComputedBOM(strBomName, 0, strConfig, EdmBomFlag.EdmBf_ShowSelected)
    If localBomView Is Nothing Then
        MsgBox "Could not retrieve BOM """ & strBomName & """ from:" & vbCrLf & strFilePath, vbCritical
        Exit Function
    End If
    ' --- 4. Columns & rows ----------------------------------------------------
    Dim bomCols() As EdmLib.EdmBomColumn
    localBomView.GetColumns bomCols
    If IsEmpty(bomCols) Then MsgBox "BOM has no columns.", vbExclamation: Exit Function
    Dim bomRows() As Variant
    localBomView.GetRows bomRows
    If IsEmpty(bomRows) Then MsgBox "BOM has no rows.", vbExclamation: Exit Function
    If Not IsArray(bomRows) Then
        Dim singleRow As Variant
        singleRow = bomRows
        ReDim bomRows(0 To 0)
        Set bomRows(0) = singleRow
    End If
    ' --- 5. Build level lookup from reference tree ----------------------------
    Dim levelDict  As New Scripting.Dictionary
    Dim refFolder  As EdmLib.IEdmFolder5
    Dim refFile    As EdmLib.IEdmFile5
    Set refFile = localVault.GetFileFromPath(strFilePath, refFolder)
    If Not refFile Is Nothing Then
        Dim rootRef As IEdmReference10
        Set rootRef = refFile.GetReferenceTree(refFolder.ID)
        BuildLevelDict rootRef, 1, strConfig, levelDict
    End If
    ' --- 6. Resolve each column to its colMap key ----------------------------
    Dim varMgr As EdmLib.IEdmVariableMgr5
    Set varMgr = localVault
    Dim colKeys() As String
    ReDim colKeys(LBound(bomCols) To UBound(bomCols))
    Dim c      As Integer
    Dim pos    As EdmLib.IEdmPos5
    Dim varObj As EdmLib.IEdmVariable5
    For c = LBound(bomCols) To UBound(bomCols)
        If bomCols(c).meType = 0 Then
            If bomCols(c).mlVariableID > 0 Then
                Set pos = varMgr.GetFirstVariablePosition()
                Do While Not pos.IsNull
                    Set varObj = varMgr.GetNextVariable(pos)
                    If varObj.ID = bomCols(c).mlVariableID Then colKeys(c) = varObj.Name: Exit Do
                Loop
            End If
        Else
            colKeys(c) = "T" & bomCols(c).meType
        End If
    Next c
    ' --- 7. Clear table and open recordset -----------------------------------
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb()
    db.Execute "DELETE FROM [" & strTableName & "]", dbFailOnError
    Set rs = db.OpenRecordset(strTableName, dbOpenDynaset)
    ' --- 8. Write rows -------------------------------------------------------
    Dim levelCounters As New Scripting.Dictionary
    Dim r           As Integer
    Dim treeLevel   As Long
    Dim k           As Variant
    Dim lvl         As Long
    Dim itemNo      As String
    Dim rowConfig   As String
    Dim cellVal     As Variant
    Dim rowCell     As EdmLib.IEdmBomCell
    Dim rowPartName As String
    Dim dot         As Integer
    Dim firstRow   As EdmLib.IEdmBomCell
    Dim firstLevel As Long
    Dim startRow   As Integer
    Set firstRow = bomRows(LBound(bomRows))
    firstLevel = firstRow.GetTreeLevel()
    startRow = IIf(firstLevel = 0, LBound(bomRows) + 1, LBound(bomRows))
    For r = startRow To UBound(bomRows)
        Set rowCell = bomRows(r)
        treeLevel = rowCell.GetTreeLevel()
        If treeLevel < 1 Then GoTo NextRow      ' skip before AddNew — no orphan record
        ' Increment this level's counter
        If levelCounters.Exists(treeLevel) Then
            levelCounters(treeLevel) = levelCounters(treeLevel) + 1
        Else
            levelCounters(treeLevel) = 1
        End If
        ' Reset deeper levels so they restart when re-entered
        For Each k In levelCounters.Keys
            If k > treeLevel Then levelCounters.Remove k
        Next k
        ' Build dotted item number
        itemNo = ""
        For lvl = 1 To treeLevel
            itemNo = IIf(Len(itemNo) = 0, "", itemNo & ".") & levelCounters(lvl)
        Next lvl
        rs.AddNew
        rs.Fields("ITEM_NO").Value = itemNo
        rs.Fields("LEVEL").Value = ItemNoSortKey(itemNo)
        'rs.Fields("FilePath").Value = strFilePath
        ' NOTE: do NOT set CONFIGURATION here — T21/GetVar populates it correctly below
        ' rowConfig resets each row; GetVar may write the component's own config back into it
        rowConfig = strConfig
        rowPartName = ""
        For c = LBound(bomCols) To UBound(bomCols)
            If Len(colKeys(c)) > 0 And colMap.Exists(colKeys(c)) Then
                cellVal = Null
                ' False = read live variable values from the current file card (latest),
                ' not the as-built snapshot cached in the BOM.
                rowCell.GetVar bomCols(c).mlVariableID, bomCols(c).meType, _
                               cellVal, Nothing, rowConfig, False
                If bomCols(c).meType = 1 And Not IsNull(cellVal) Then
                    dot = InStrRev(CStr(cellVal), ".")
                    If dot > 0 Then cellVal = Left(CStr(cellVal), dot - 1)
                    rowPartName = UCase(CStr(cellVal))
                End If
                On Error Resume Next
                rs.Fields(colMap(colKeys(c))).Value = cellVal
                On Error GoTo Failure
            End If
        Next c
        rs.Update
NextRow:
    Next r
    rs.Close
    ImportBOMToTable = True
    Exit Function
Failure:
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
        rs.Close
    End If
    MsgBox "ImportBOMToTable error (#" & Err.Number & "):" & vbCrLf & Err.DESCRIPTION, vbCritical
End Function
