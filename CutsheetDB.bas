Attribute VB_Name = "CutsheetDB"
Option Compare Database
Option Explicit

' ============================================================
' Required References (Tools > References in the VBA editor):
'   [x] Microsoft DAO 3.6 Object Library  (dao360.dll)
'
' If calling from SolidWorks VBA instead of Access, also add:
'   [x] Microsoft Office XX.0 Access Database Engine Object Library
'
' Global constants expected in a separate Globals module:
'   Global Const dir_CUTSHEET As String = "C:\Path\To\Your.accdb"
'   Global Const tbl_CUTSHEET As String = "CUTSHEET_BOM"
' ============================================================


' ------------------------------------------------------------
' Helper: Set a single field on an open, AddNew/Edit recordset.
'
'   rs        - DAO.Recordset already in AddNew or Edit mode
'   strField  - name of the field in the table
'   varValue  - value to write (Null is allowed)
'
' Raises a descriptive error if the field name is not found.
' ------------------------------------------------------------
Public Sub SetField(rs As DAO.Recordset, strField As String, varValue As Variant)
    On Error GoTo ErrHandler
    rs.Fields(strField).Value = varValue
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "SetField", _
              "Field """ & strField & """ not found or type mismatch. " & Err.Description
End Sub


' ------------------------------------------------------------
' InsertCutsheetRow
'
' Opens the CUTSHEET database and appends one new row to
' tbl_CUTSHEET.  Pass in a ParamArray of alternating
' field-name / value pairs:
'
'   InsertCutsheetRow "PART_NUMBER", "ABC-123", _
'                     "DESCRIPTION", "Widget", _
'                     "QTY", 5
'
' Returns True on success, False on failure (error is logged
' to the Immediate window).
' ------------------------------------------------------------
Public Function InsertCutsheetRow(ParamArray fields_and_values() As Variant) As Boolean
    If (UBound(fields_and_values) + 1) Mod 2 <> 0 Then
        MsgBox "InsertCutsheetRow: must pass an even number of arguments (field, value pairs).", _
               vbExclamation, "CutsheetDB"
        Exit Function
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo ErrHandler

    Set db = DBEngine.OpenDatabase(dir_CUTSHEET)
    Set rs = db.OpenRecordset(tbl_CUTSHEET, dbOpenTable)

    rs.AddNew

    Dim i As Long
    For i = 0 To UBound(fields_and_values) - 1 Step 2
        SetField rs, CStr(fields_and_values(i)), fields_and_values(i + 1)
    Next i

    rs.Update
    InsertCutsheetRow = True

Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Exit Function

ErrHandler:
    Debug.Print "InsertCutsheetRow error " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Function


' ------------------------------------------------------------
' Test_InsertCutsheetRow
'
' Smoke-test: inserts one hard-coded row and reports the
' result in a message box.  Remove or adapt before shipping.
' ------------------------------------------------------------
Public Sub Test_InsertCutsheetRow()
    Dim bOK As Boolean
    bOK = InsertCutsheetRow( _
        "PART_NUMBER", "TEST-001", _
        "DESCRIPTION", "Test row from VBA", _
        "QTY", 1)

    If bOK Then
        MsgBox "Row inserted successfully into " & tbl_CUTSHEET & ".", _
               vbInformation, "CutsheetDB Test"
    Else
        MsgBox "Insert FAILED. Check the Immediate window for details.", _
               vbCritical, "CutsheetDB Test"
    End If
End Sub
