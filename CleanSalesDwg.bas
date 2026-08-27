Attribute VB_Name = "CleanSalesDwg"
'==========================================================================
' CleanSalesDwg - the active drawing sheet out as a DXF and a PDF at once,
' with the small holes taken out of the DXF.
'
' Run CleanSalesDrawing from Tools > Macro, or call ExportSalesDwg from your
' own code - it returns True only when both files are on disk.
'
' WHAT COMES OUT.  <drawing>.dxf and <drawing>.pdf, in the same folder, both
' of whichever sheet is active when you run it, so the pair always matches.
' The raw export the DXF was cleaned from stays in %TEMP%\CleanSalesDwg next
' to the log, which holds the script's own account of what it removed.
'
' WHERE THE HOLES ACTUALLY GO.  Not here.  CleanSalesDwg.py beside this file
' does that with ezdxf, and this module exports and hands over.  One thing in
' the pipeline understands DXF geometry and it is not the VBA - which is also
' why a surprising result can be reproduced at a console without SolidWorks
' open at all:
'
'     python CleanSalesDwg.py "raw.dxf" --out "C:\somewhere"
'
' THE DIAMETER IS IN DRAWING UNITS AND IS MEASURED ON THE SHEET.  100 is
' 100 mm on a metric drawing and 100 inches on an inch one, and a view at 1:2
' puts a 200 hole into the file as a 100 circle.  Export the views 1:1 - which
' is what this is for - and sheet size is part size.
'
' Only circles go.  A centre mark, a centreline or the dimension calling a
' hole out is not a circle and stays; strip those in the drawing if the
' customer's copy should not carry them.  A balloon RING under the cut-off IS
' a circle and does go - see strip_small_circles in CleanSalesDwg.py, which
' reports what it took out of each block by name, and --no-blocks if you would
' rather it left them alone.
'
' MULTI-SHEET.  The ACTIVE sheet, so Tools > Options > Export > DXF/DWG has to
' be set to export the active sheet only.  Set to all sheets, SolidWorks
' writes the whole drawing into the one file and you get a cleaned copy of all
' of it - no error, just more than you asked for.
'
' Self-contained on purpose: nothing outside this module and the script, so it
' imports into any macro project as one file.  Every private helper is
' cs_-prefixed so a project that already has a Quote() or a Tail() keeps it.
'==========================================================================

Option Explicit

' Where CleanSalesDwg.py is checked out.  The one line to fix per workstation.
Private Const SCRIPT_PATH As String = "C:\Users\Azar\SolidWorks\CleanSalesDwg.py"

' "" = work the interpreter out (see cs_PythonCmd).  Put a full path here if
' this machine has several and one of them is the one with ezdxf in it.
Private Const PYTHON_EXE As String = ""

' Prepopulates the input box.
Private Const MIN_HOLE_DIA As Double = 100


'==========================  ENTRY POINTS  ================================

' The one to run by hand.  Asks for the diameter, writes both files beside
' the drawing.
Sub CleanSalesDrawing()
    Dim ok As Boolean
    ok = ExportSalesDwg("", MIN_HOLE_DIA, True)
End Sub


' sOutFolder  "" = the folder the drawing is saved in.
' dMinDia     circles under this diameter come out, in drawing units.
' bAsk        True shows the input box prepopulated with dMinDia and the
'             report at the end; False takes the number as given and puts up
'             nothing unless something fails.
'
' True only when the DXF and the PDF are both on disk afterwards.  Every
' failure path says why in a message box first, so a caller only has to test
' the result.
Public Function ExportSalesDwg(Optional ByVal sOutFolder As String = "", _
                               Optional ByVal dMinDia As Double = MIN_HOLE_DIA, _
                               Optional ByVal bAsk As Boolean = True) As Boolean

    Dim swApp        As SldWorks.SldWorks
    Dim swModel      As SldWorks.ModelDoc2

    Dim drawingPath  As String
    Dim drawingName  As String
    Dim sheetName    As String
    Dim workDir      As String
    Dim rawDxf       As String
    Dim dxfFile      As String
    Dim pdfFile      As String
    Dim LogFile      As String

    Dim sAnswer      As String
    Dim Cmd          As String
    Dim rc           As Long
    Dim nErr         As Long
    Dim nWarn        As Long

    ' --- 1. The document has to be a drawing, saved to disk -------------------
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Open a drawing first.", vbExclamation, "CleanSalesDwg"
        Exit Function
    End If

    If swModel.GetType <> swDocDRAWING Then
        MsgBox "The active document is not a drawing.", vbExclamation, "CleanSalesDwg"
        Exit Function
    End If

    drawingPath = swModel.GetPathName

    If Len(drawingPath) = 0 Then
        MsgBox "Save the drawing to disk first - the file name drives the " & _
               "output names.", vbExclamation, "CleanSalesDwg"
        Exit Function
    End If

    If Not cs_FileExists(SCRIPT_PATH) Then
        MsgBox "Cannot find:" & vbCrLf & SCRIPT_PATH & vbCrLf & vbCrLf & _
               "Fix SCRIPT_PATH at the top of the module.", vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    drawingName = cs_Title(drawingPath)

    ' --- 2. The diameter -----------------------------------------------------
    If bAsk Then
        sAnswer = InputBox( _
            "Remove full circles smaller than this DIAMETER." & vbCrLf & vbCrLf & _
            "Drawing units, measured on the sheet - so this assumes the " & _
            "views are 1:1." & vbCrLf & _
            "A circle exactly this size is kept.", _
            "Clean sales drawing - DXF + PDF", _
            Trim$(Str$(dMinDia)))

        ' Cancel and an empty box both come back empty, and both mean stop.
        If Len(Trim$(sAnswer)) = 0 Then Exit Function

        ' Val reads a full stop whatever the machine's locale is; the Replace
        ' is for whoever types 100,5 because their Windows told them to.
        dMinDia = Val(Replace$(Trim$(sAnswer), ",", "."))
    End If

    If dMinDia <= 0 Then
        MsgBox "That is not a diameter: " & sAnswer, vbExclamation, "CleanSalesDwg"
        Exit Function
    End If

    ' --- 3. Paths ------------------------------------------------------------
    ' Same basename in two folders, which is the whole convention: the raw
    ' export lands in workDir, the cleaned copy of it lands in the output
    ' folder.  --out takes a FOLDER, and the script refuses to write over the
    ' file it read.
    If Len(sOutFolder) = 0 Then sOutFolder = cs_Folder(drawingPath)
    sOutFolder = cs_Slash(sOutFolder)

    workDir = cs_Slash(Environ$("TEMP") & "\CleanSalesDwg\" & drawingName)

    If Not cs_MakeFolder(sOutFolder) Then
        MsgBox "Cannot create the output folder:" & vbCrLf & sOutFolder, _
               vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    If Not cs_MakeFolder(workDir) Then
        MsgBox "Cannot create the working folder:" & vbCrLf & workDir, _
               vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    rawDxf = workDir & drawingName & ".dxf"
    dxfFile = sOutFolder & drawingName & ".dxf"
    pdfFile = sOutFolder & drawingName & ".pdf"
    LogFile = workDir & drawingName & "_clean.log"

    ' --- 4. Save first, so the exports match what is on screen ---------------
    If swModel.GetSaveFlag Then
        swModel.Save3 1, nErr, nWarn                   ' 1 = swSaveAsOptions_Silent
    End If

    ' --- 5. The two exports, of the one sheet that is active -----------------
    sheetName = swModel.GetCurrentSheet.GetName

    If Not cs_ExportDxf(swModel, rawDxf) Then
        MsgBox "DXF export failed - SolidWorks wrote no file:" & vbCrLf & _
               rawDxf & vbCrLf & vbCrLf & _
               "Check Tools > Options > Export > DXF/DWG.", _
               vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    If Not cs_ExportPdf(swApp, swModel, pdfFile, sheetName) Then
        MsgBox "PDF export failed - SolidWorks wrote no file:" & vbCrLf & pdfFile, _
               vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    ' --- 6. And the script takes the small circles out of the DXF ------------
    ' cs_NoTrailSlash on the folder is NOT cosmetic.  Windows argument parsing
    ' reads \" as an escaped quote, so --out "C:\pub\" hands Python the string
    ' C:\pub" with the quote welded on and the rest of the command line
    ' shifted.  A quoted argument must never end in \.
    '
    ' Str$ rather than Format$ or a bare concatenation: it writes a full stop
    ' for the decimal point whatever the machine's locale is, and Python is not
    ' going to parse 100,5.  It leads with a space for a positive number, hence
    ' the Trim$.
    cs_ClearTarget dxfFile

    Cmd = cs_PythonCmd() & _
          " -u " & cs_Quote(SCRIPT_PATH) & _
          " " & cs_Quote(rawDxf) & _
          " --min-dia " & Trim$(Str$(dMinDia)) & _
          " --out " & cs_Quote(cs_NoTrailSlash(sOutFolder))

    rc = cs_RunAndWait(Cmd, LogFile)

    If rc <> 0 Then
        MsgBox "CleanSalesDwg.py failed, exit code " & rc & _
               IIf(rc = 9009, "  (the interpreter was not found)", "") & _
               vbCrLf & vbCrLf & cs_Tail(LogFile, 24) & vbCrLf & _
               "Full log: " & LogFile, vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    ' --- 7. Ask the disk -----------------------------------------------------
    ' A zero exit code says the script ran, not that the file it was asked for
    ' is sitting there.
    If Not cs_FileExists(dxfFile) Then
        MsgBox "The script reported success but nothing arrived at:" & vbCrLf & _
               dxfFile & vbCrLf & vbCrLf & "Full log: " & LogFile, _
               vbCritical, "CleanSalesDwg"
        Exit Function
    End If

    ExportSalesDwg = True

    If bAsk Then
        ' The script's own count of what it removed, out of the log it just
        ' wrote - rather than this module guessing at it.
        MsgBox sOutFolder & vbCrLf & _
               "   " & drawingName & ".dxf" & vbCrLf & _
               "   " & drawingName & ".pdf" & vbCrLf & vbCrLf & _
               "Sheet: " & sheetName & vbCrLf & vbCrLf & _
               cs_Tail(LogFile, 10), _
               vbInformation, "CleanSalesDwg"
    End If

End Function


'==========================  EXPORT  ======================================

' DXF of whichever sheet is active right now.
'
' The one that matters: Extension.SaveAs returns TRUE having written nothing
' often enough to bite - a suppressed export dialog, an empty sheet, a locked
' target.  Trusting that boolean reports success and leaves the script to
' choke on a file that was never there.  Ask the disk.
Private Function cs_ExportDxf(ByVal swModel As SldWorks.ModelDoc2, _
                              ByVal sTarget As String) As Boolean

    Dim nErr As Long, nWarn As Long
    Dim bOk  As Boolean

    cs_ClearTarget sTarget

    ' Scoped, not blanket.  A COM fault here is information, not noise.
    On Error Resume Next
    ' 0 = swSaveAsCurrentVersion, 1 = swSaveAsOptions_Silent
    bOk = swModel.Extension.SaveAs(sTarget, 0, 1, Nothing, nErr, nWarn)
    On Error GoTo 0

    cs_ExportDxf = bOk And cs_FileExists(sTarget)

End Function


' One sheet's PDF, without launching Acrobat.
'
' The sheet is NAMED rather than left to an "all sheets" flag, so the PDF is
' of the same sheet the DXF just came from.  GetExportFileData / SetSheets are
' the calls that legitimately vary between SolidWorks versions, so the Resume
' Next goes around those two and nothing else - a failing SaveAs has to be
' able to report itself.
Private Function cs_ExportPdf(ByVal swApp As SldWorks.SldWorks, _
                              ByVal swModel As SldWorks.ModelDoc2, _
                              ByVal sTarget As String, _
                              ByVal sSheet As String) As Boolean

    Dim nErr As Long, nWarn As Long
    Dim swExportData As Object
    Dim bOk As Boolean

    cs_ClearTarget sTarget

    Set swExportData = Nothing
    On Error Resume Next
    Set swExportData = swApp.GetExportFileData(1)       ' 1 = swExportPdfData
    If Not swExportData Is Nothing Then
        swExportData.ViewPdfAfterSaving = False
        ' 3 = swExportData_ExportSpecifiedSheets
        swExportData.SetSheets 3, Array(sSheet)
    End If
    On Error GoTo 0

    bOk = swModel.Extension.SaveAs(sTarget, 0, 1, swExportData, nErr, nWarn)

    cs_ExportPdf = bOk And cs_FileExists(sTarget)

End Function


'==========================  PROCESS  =====================================

' The interpreter, ready to drop into a command line.
'
' No searching the disk: the py launcher ships with every python.org install
' and picks the newest 3.x itself, and PATH covers the rest.  If the machine
' has several and only one of them has ezdxf, name it in PYTHON_EXE above.
Private Function cs_PythonCmd() As String

    If Len(PYTHON_EXE) > 0 Then
        cs_PythonCmd = cs_Quote(PYTHON_EXE)
    ElseIf cs_FileExists(Environ$("WINDIR") & "\py.exe") Then
        cs_PythonCmd = "py -3"
    Else
        cs_PythonCmd = "python"
    End If

End Function


' Run a command line, capture stdout AND stderr into sLog, wait, return the
' exit code.
'
' The log is truncated rather than appended to, or cs_Tail shows last week's
' traceback and you chase a bug you already fixed.  The command line goes in
' as the first line, so a failure can be pasted straight into a console and
' watched again.
Private Function cs_RunAndWait(ByVal sCmd As String, ByVal sLog As String) As Long

    Dim fso   As Object
    Dim ts    As Object
    Dim wsh   As Object
    Dim sFull As String

    On Error GoTo Failure

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(sLog, True)            ' True = overwrite
    ts.WriteLine "> " & sCmd
    ts.WriteLine String(60, "-")
    ts.Close

    Set wsh = CreateObject("WScript.Shell")

    ' cmd /S /C "<command> >> "<log>" 2>&1"
    ' /S strips only the outermost pair of quotes and takes the rest verbatim,
    ' which is what keeps every quoted path inside it intact.
    sFull = "cmd.exe /S /C """ & sCmd & " >> " & cs_Quote(sLog) & " 2>&1"""

    cs_RunAndWait = wsh.Run(sFull, 0, True)            ' 0 = hidden, True = wait
    Exit Function

Failure:
    cs_RunAndWait = -1

End Function


' Last nLines of the log, for the message box.  Without this a failed run is
' an exit code and nothing else; with it you get Python's traceback in the
' dialog, and a successful one shows what came out of the drawing.
Private Function cs_Tail(ByVal sPath As String, ByVal nLines As Long) As String

    Dim fso    As Object
    Dim ts     As Object
    Dim vAll   As Variant
    Dim i      As Long
    Dim nStart As Long
    Dim sOut   As String

    If Not cs_FileExists(sPath) Then
        cs_Tail = "(no log file was written)"
        Exit Function
    End If

    On Error GoTo Failure

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(sPath, 1)                ' 1 = ForReading

    If ts.AtEndOfStream Then
        ts.Close
        cs_Tail = "(the log file is empty)"
        Exit Function
    End If

    vAll = Split(ts.ReadAll, vbCrLf)
    ts.Close

    nStart = UBound(vAll) - nLines + 1
    If nStart < LBound(vAll) Then nStart = LBound(vAll)

    For i = nStart To UBound(vAll)
        If Len(Trim$(CStr(vAll(i)))) > 0 Then
            sOut = sOut & CStr(vAll(i)) & vbCrLf
        End If
    Next i

    cs_Tail = sOut
    Exit Function

Failure:
    cs_Tail = "(the log could not be read)"

End Function


'==========================  SMALL STUFF  =================================

Private Function cs_Quote(ByVal s As String) As String
    cs_Quote = Chr$(34) & s & Chr$(34)
End Function


Private Function cs_FileExists(ByVal sPath As String) As Boolean
    If Len(sPath) = 0 Then Exit Function
    On Error Resume Next
    cs_FileExists = (Len(Dir$(sPath, vbNormal)) > 0)
End Function


' Drop a stale output so a failed run cannot leave yesterday's file looking
' like today's.  SetAttr first - a read-only leftover quietly defeats both
' Kill and SaveAs.
Private Sub cs_ClearTarget(ByVal sTarget As String)
    If Not cs_FileExists(sTarget) Then Exit Sub
    On Error Resume Next
    SetAttr sTarget, vbNormal
    Kill sTarget
    On Error GoTo 0
End Sub


' Create a folder and every parent it needs, and report against the disk
' rather than against whatever the last CreateFolder returned.
Private Function cs_MakeFolder(ByVal sPath As String) As Boolean

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sPath) Then
        cs_MakeFolder = True
        Exit Function
    End If

    On Error Resume Next
    cs_MakeParent fso, fso.GetParentFolderName(cs_NoTrailSlash(sPath))
    fso.CreateFolder cs_NoTrailSlash(sPath)
    On Error GoTo 0

    cs_MakeFolder = fso.FolderExists(sPath)

End Function


Private Sub cs_MakeParent(ByVal fso As Object, ByVal sPath As String)
    If Len(sPath) = 0 Then Exit Sub
    If fso.FolderExists(sPath) Then Exit Sub
    cs_MakeParent fso, fso.GetParentFolderName(sPath)
    fso.CreateFolder sPath
End Sub


Private Function cs_Slash(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    If Right$(s, 1) <> "\" Then s = s & "\"
    cs_Slash = s
End Function


Private Function cs_NoTrailSlash(ByVal s As String) As String
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    cs_NoTrailSlash = s
End Function


' "M:\Jobs\16022.SLDDRW" -> "M:\Jobs\"
Private Function cs_Folder(ByVal sPath As String) As String
    Dim n As Long
    n = InStrRev(sPath, "\")
    If n > 0 Then cs_Folder = Left$(sPath, n)
End Function


' "M:\Jobs\16022.SLDDRW" -> "16022"
Private Function cs_Title(ByVal sPath As String) As String

    Dim s As String
    Dim n As Long

    s = Mid$(sPath, InStrRev(sPath, "\") + 1)
    n = InStrRev(s, ".")
    If n > 1 Then s = Left$(s, n - 1)

    cs_Title = s

End Function
