Option Explicit

' ======= CONFIG =======
Const NAME_COL As Long = 1           ' Column A = production names
Const PARAM_START_COL As Long = 2    ' Parameters start at column B
Const PATH_COL As Long = 6           ' Column F = save path
Const OPEN_RETRY_MS As Long = 150
Const OPEN_MAX_RETRIES As Long = 10
' ======================

Public Sub GenerateInventorParts_SaveAs()
    Dim invApp As Object                     ' Inventor.Application
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim masterDoc As Object, masterPath As String, masterFolder As String
    Dim fso As Object
    Dim partName As String, targetPath As String, ext As String
    Dim paramName As String, paramValue As Variant
    Dim newDoc As Object, params As Object

    ' --- Connect to Inventor, require master to be open ---
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    On Error GoTo 0
    If invApp Is Nothing Or invApp.ActiveDocument Is Nothing Then
        MsgBox "Open Inventor and load the MASTER document first.", vbExclamation
        Exit Sub
    End If
    invApp.Visible = True

    ' Excel
    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Master info
    Set masterDoc = invApp.ActiveDocument
    masterPath = masterDoc.FullFileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    masterFolder = fso.GetParentFolderName(masterPath)

    ' ---------------------------------------
    ' LOOP 1 — SAVEAS ALL COPIES & WRITE PATH
    ' ---------------------------------------
    Application.StatusBar = "Phase 1: Creating SaveAs copies..."
    For r = 2 To lastRow
        partName = Trim$(ws.Cells(r, NAME_COL).Value)
        If Len(partName) = 0 Then GoTo NextRowCopy

        ' If no extension supplied, default to part
        ext = LCase$(Right$(partName, 4))
        If ext <> ".ipt" And ext <> ".iam" Then
            targetPath = masterFolder & "\" & partName & ".ipt"
        Else
            targetPath = masterFolder & "\" & partName
        End If

        ' Open master fresh (writable), SaveAs new name, close the new doc
        On Error GoTo CopyErr
        Dim tempDoc As Object
        Set tempDoc = invApp.Documents.Open(masterPath, False)
        tempDoc.SaveAs targetPath, True           ' overwrite if exists
        ' IMPORTANT: after SaveAs, tempDoc now represents the NEW file on disk.
        tempDoc.Close False                       ' close so we can open all later
        ws.Cells(r, PATH_COL).Value = targetPath
        ws.Rows(r).Interior.Color = RGB(200, 200, 255) ' blue = copied
        Application.StatusBar = "Saved: " & targetPath
        DoEvents
        On Error GoTo 0
        GoTo NextRowCopy

CopyErr:
        ws.Cells(r, PATH_COL).Value = "SAVEAS FAILED: " & Err.Description
        ws.Rows(r).Interior.Color = RGB(255, 204, 204)
        Err.Clear
        On Error GoTo 0
NextRowCopy:
    Next r

    ' ----------------------------------------------------
    ' LOOP 2 — OPEN EACH COPY, UPDATE PARAMS, SAVE, LEAVE OPEN
    ' ----------------------------------------------------
    Application.StatusBar = "Phase 2: Opening and updating parameters (mm)..."
    For r = 2 To lastRow
        targetPath = CStr(ws.Cells(r, PATH_COL).Value)
        If targetPath = "" Or InStr(1, targetPath, "FAILED", vbTextCompare) > 0 Then GoTo NextRowUpdate
        If Not fso.FileExists(targetPath) Then
            ws.Cells(r, PATH_COL).Value = "MISSING: " & targetPath
            ws.Rows(r).Interior.Color = RGB(255, 204, 204)
            GoTo NextRowUpdate
        End If

        ' Open (no Activate — not required and can fail)
        Set newDoc = invApp.Documents.Open(targetPath, False)
        If newDoc Is Nothing Then
            ws.Cells(r, PATH_COL).Value = "OPEN FAILED"
            GoTo NextRowUpdate
        End If
        If Not WaitForDocReady(newDoc, OPEN_MAX_RETRIES, OPEN_RETRY_MS) Then
            ws.Cells(r, PATH_COL).Value = "DOC NOT READY"
            GoTo NextRowUpdate
        End If

        ' Only part/assembly have Parameters on ComponentDefinition
        On Error Resume Next
        Set params = newDoc.ComponentDefinition.Parameters
        On Error GoTo 0
        If params Is Nothing Then
            ws.Cells(r, PATH_COL).Value = "NO PARAMETERS"
            GoTo NextRowUpdate
        End If

        ' Highlight current row while updating
        ws.Rows(r).Interior.Color = RGB(255, 255, 153) ' yellow

        ' Update by headers (B1..), writing values in mm
        For c = PARAM_START_COL To lastCol
            paramName = Trim$(ws.Cells(1, c).Value)
            paramValue = ws.Cells(r, c).Value
            If Len(paramName) = 0 Or Len(CStr(paramValue)) = 0 Then GoTo NextParam

            ' clear previous formatting/comments
            ws.Cells(r, c).Interior.ColorIndex = xlNone
            ClearNote ws.Cells(r, c)

            Application.StatusBar = "Updating " & targetPath & " | " & paramName
            DoEvents

            If SetParameterInMM(params, paramName, paramValue) Then
                ws.Cells(r, c).Interior.Color = RGB(144, 238, 144) ' green
                SafeNote ws.Cells(r, c), "Updated (" & paramName & " = " & CStr(paramValue) & " mm)"
            Else
                ws.Cells(r, c).Interior.Color = RGB(255, 102, 102) ' red
                SafeNote ws.Cells(r, c), "NOT FOUND: " & paramName
            End If

            On Error Resume Next
            newDoc.Update
            invApp.ActiveView.Update
            On Error GoTo 0
            DoEvents
NextParam:
        Next c

        ' Save (leave open so you can see results)
        On Error Resume Next
        newDoc.Save
        If Err.Number <> 0 Then
            ' If Save fails due to unresolved references, let user know in the sheet
            ws.Cells(r, PATH_COL).Value = "SAVE FAILED: resolve refs (Vault?)"
            Err.Clear
        End If
        On Error GoTo 0

        ws.Rows(r).Interior.Color = RGB(204, 255, 204) ' light green = done
        Application.StatusBar = "Done: " & targetPath
        DoEvents

NextRowUpdate:
        Set params = Nothing
        ' intentionally not closing newDoc
    Next r

    Application.StatusBar = False
    MsgBox "Finished: copies made with SaveAs, parameters set in mm, parts left open in Inventor.", vbInformation
End Sub

' ===== Helpers =====

Private Function WaitForDocReady(doc As Object, maxRetries As Long, waitMs As Long) As Boolean
    Dim i As Long
    WaitForDocReady = True
    For i = 1 To maxRetries
        On Error Resume Next
        Dim t$: t = doc.DisplayName
        If Err.Number = 0 Then Exit Function
        Err.Clear
        On Error GoTo 0
        SleepMs waitMs
    Next i
    WaitForDocReady = False
End Function

Private Sub SleepMs(ms As Long)
    Dim t As Single: t = Timer
    Do While Timer - t < (ms / 1000#)
        DoEvents
    Loop
End Sub

' Set parameter named `name` to valueInMM (always written as an Expression in mm)
Private Function SetParameterInMM(ByVal params As Object, ByVal name As String, ByVal valueInMM As Variant) As Boolean
    Dim p As Object
    SetParameterInMM = False

    On Error Resume Next
    Set p = params.Item(name)
    If Err.Number <> 0 Or p Is Nothing Then
        Err.Clear
        Exit Function                     ' parameter not found
    End If

    Err.Clear
    p.Expression = CStr(valueInMM) & " mm"
    If Err.Number = 0 Then
        SetParameterInMM = True
    Else
        ' fallback for non-length (text/bool) parameters
        Err.Clear
        p.Expression = CStr(valueInMM)
        If Err.Number = 0 Then SetParameterInMM = True
    End If
    On Error GoTo 0
End Function

' Comments that work across Excel versions
Private Sub SafeNote(ByVal cell As Range, ByVal txt As String)
    On Error Resume Next
    cell.ClearComments
    cell.AddComment txt
    If Err.Number <> 0 Then
        Err.Clear
        cell.ClearComments
        cell.AddCommentThreaded txt
    End If
    On Error GoTo 0
End Sub

Private Sub ClearNote(ByVal cell As Range)
    On Error Resume Next
    cell.ClearComments
    cell.CommentThreaded.Delete
    On Error GoTo 0
End Sub

