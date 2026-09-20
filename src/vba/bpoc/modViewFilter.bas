Attribute VB_Name = "modViewFilter"
Option Explicit

' ==========================================================================
' BPOC V6 - "Active only" view toggle
'
' Separated cadets cannot be SORTED to the bottom: row 12 has to be the same
' cadet on Cadets, Spelling, Writing, PT, sysGrades and every other sheet, so
' re-ordering rows would break the whole engine. Filtering achieves the same
' thing for data entry without moving anything - the row keeps its place and
' its data, it just stops being in your way.
'
' These two macros apply (or clear) that filter on EVERY cadet grid at once,
' so it is one click instead of nine.
' ==========================================================================

Private Const PW As String = "TPDAcademy"
Private Const HDR As Long = 5

' every sheet where one row = one cadet
Private Function CadetGrids() As Variant
    CadetGrids = Array("Cadets", "Spelling", "Writing", "PT", _
                       "Certifications", "SkillsCheck", "StateExam", _
                       "ScoresGrid", "GradChecklist")
End Function

' The status column is found by its HEADER, never by a hard-coded letter:
' columns get appended to these sheets over time, and a literal letter would
' silently start filtering the wrong column. On Cadets the cadet's own
' "Status" column is authoritative; everywhere else it is the mirrored
' "Cadet Status" (StateExam has a "Status" column too, but it holds the exam
' result, so matching the bare word there would filter the wrong thing).
Private Function StatusCol(ws As Worksheet) As Long
    Dim c As Long, h As String
    For c = 2 To 80
        h = Trim$(CStr(ws.Cells(HDR, c).Value))
        If h = "Cadet Status" Then StatusCol = c: Exit Function
        If h = "Status" And ws.Name = "Cadets" Then StatusCol = c: Exit Function
    Next c
    StatusCol = 0
End Function

Public Sub ShowActiveOnly()
    ApplyView True
End Sub

Public Sub ShowAllCadets()
    ApplyView False
End Sub

Private Sub ApplyView(activeOnly As Boolean)
    Dim names As Variant: names = CadetGrids()
    Dim i As Long, done As Long, skipped As String
    Dim ws As Worksheet, col As Long, wasProt As Boolean

    Application.ScreenUpdating = False
    On Error GoTo Cleanup
    For i = LBound(names) To UBound(names)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(names(i))
        On Error GoTo Cleanup
        If ws Is Nothing Then
            skipped = skipped & vbCrLf & "  " & names(i) & " (sheet missing)"
        Else
            col = StatusCol(ws)
            If col = 0 Then
                skipped = skipped & vbCrLf & "  " & ws.Name & " (no status column)"
            Else
                ' protection blocks changing a filter even when it allows the
                ' user to USE one, so unprotect around the change and put the
                ' sheet back exactly as it was
                wasProt = ws.ProtectContents
                ws.Unprotect PW
                If ws.AutoFilterMode = False Then
                    ws.Range(ws.Cells(HDR, 2), _
                             ws.Cells(HDR, LastHeaderCol(ws))).AutoFilter
                End If
                ' field number is relative to the filter range, which starts
                ' at column B
                If activeOnly Then
                    ws.Range("B" & HDR).AutoFilter Field:=col - 1, _
                        Criteria1:="Active"
                Else
                    ws.Range("B" & HDR).AutoFilter Field:=col - 1
                End If
                If wasProt Then ws.Protect PW
                ' re-assert: Protect resets these to "blocked"
                ws.EnableAutoFilter = True
                ws.Protect Password:=PW, AllowFiltering:=True
                done = done + 1
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    If activeOnly Then
        MsgBox "Showing ACTIVE cadets only on " & done & " sheet(s)." & vbCrLf & _
               "Separated cadets keep their rows and their data - they are " & _
               "just hidden from view." & vbCrLf & _
               "Use 'Show All Cadets' to bring them back." & _
               IIf(skipped = "", "", vbCrLf & vbCrLf & "Skipped:" & skipped), _
               vbInformation, "Active only"
    Else
        MsgBox "Showing ALL cadets on " & done & " sheet(s)." & _
               IIf(skipped = "", "", vbCrLf & vbCrLf & "Skipped:" & skipped), _
               vbInformation, "Show all"
    End If
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    MsgBox "Could not finish changing the view: " & Err.Description & _
           vbCrLf & vbCrLf & "Any sheet already changed is fine; run " & _
           "'Show All Cadets' to clear the filters.", vbExclamation
End Sub

Private Function LastHeaderCol(ws As Worksheet) As Long
    Dim c As Long, lastC As Long
    lastC = 2
    For c = 2 To 80
        If Trim$(CStr(ws.Cells(HDR, c).Value)) <> "" Then lastC = c
    Next c
    LastHeaderCol = lastC
End Function
