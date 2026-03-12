Attribute VB_Name = "Form_Usage_Example"
Option Compare Database
Option Explicit

' Example: call GetConfigs and bind the result to cbo_CONFIG.
' Replace this with the actual event (e.g. txt_CUTSHEET_AfterUpdate) in your form.
Private Sub LoadConfigCombo()
    Dim vConfigs As Variant
    vConfigs = GetConfigs(txt_CUTSHEET.Value)
    If IsEmpty(vConfigs) Then GoTo SkipTo

    cbo_CONFIG.RowSourceType = "Value List"
    cbo_CONFIG.RowSource = ""

    Dim i As Long
    For i = LBound(vConfigs) To UBound(vConfigs)
        cbo_CONFIG.AddItem vConfigs(i)
    Next i

SkipTo:
End Sub
