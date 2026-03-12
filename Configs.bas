Attribute VB_Name = "Configs"
Option Compare Database
Option Explicit

' Returns an array of filtered config names for the given PDM file title.
' Configs containing "@" or "SM-FLAT-PATTERN" are excluded.
' Returns an empty Variant (IsEmpty = True) if no valid configs are found.
Public Function GetConfigs(str_FileTitle As String) As Variant
    Dim pdmResult As IEdmSearchResult5
    Set pdmResult = GetPDMResult(str_FileTitle)
    If pdmResult Is Nothing Then Exit Function

    Dim vConfigs() As Variant
    vConfigs = GetConfigsArr(pdmResult.Path)
    If IsEmpty(vConfigs) Then Exit Function

    Dim col As Collection
    Set col = New Collection

    Dim i As Long
    For i = 0 To UBound(vConfigs)
        If InStr(vConfigs(i), "@") = 0 And InStr(vConfigs(i), "SM-FLAT-PATTERN") = 0 Then
            col.Add vConfigs(i)
        End If
    Next i

    If col.Count = 0 Then Exit Function

    Dim a() As Variant
    ReDim a(0 To col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col.Item(i)
    Next i

    GetConfigs = a
End Function
