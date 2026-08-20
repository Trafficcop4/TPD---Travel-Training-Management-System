Attribute VB_Name = "modNewAcademy"
Option Explicit

' ==========================================================================
' BPOC V6 - New Academy reset
' Save-As the workbook first, then run this on the COPY. Clears all cadet
' and daily-operations data; keeps masters (chapters, exams, skills,
' spelling words, writing prompts, agencies, instructors) and Settings.
' Optionally clears the Schedule.
' ==========================================================================

Private Const PW As String = "TPDAcademy"
Private Const ROW1 As Long = 6

Public Sub NewAcademyReset()
    If MsgBox("This clears ALL cadet data, scores, attendance, incidents, " & _
              "counseling, PT, medical, reviews and email logs from THIS file." & _
              vbCrLf & vbCrLf & "Have you saved a copy of the finished academy " & _
              "first?", vbYesNo + vbExclamation, "New Academy") <> vbYes Then Exit Sub

    Dim clearSched As Boolean
    clearSched = (MsgBox("Also clear the Schedule (class-day time blocks)?", _
                  vbYesNo + vbQuestion, "New Academy") = vbYes)

    Application.ScreenUpdating = False
    ClearRange "Cadets", "B6:E55": ClearRange "Cadets", "G6:G55"
    ClearRange "Cadets", "I6:I55": ClearRange "Cadets", "K6:M55"
    ClearRange "ExamScores", "C6:C1505": ClearRange "ExamScores", "F6:F1505"
    ClearRange "ExamScores", "K6:L1505": ClearRange "ExamScores", "S6:S1505"
    ClearRange "ExamScores", "V6:W1505"
    ClearRange "Spelling", "D6:O55"
    ClearRange "Attendance", "C6:D805": ClearRange "Attendance", "G6:J805"
    ClearRange "Attendance", "L6:M805": ClearRange "Attendance", "O6:O805"
    ClearRange "Makeup", "C6:D505": ClearRange "Makeup", "F6:K505"
    ClearRange "Makeup", "M6:M505"
    ClearRange "Skills", "C6:C605": ClearRange "Skills", "E6:E605"
    ClearRange "Skills", "I6:M605": ClearRange "Skills", "Q6:R605"
    ClearRange "Writing", "D6:AQ55"
    ClearRange "Incidents", "C6:D405": ClearRange "Incidents", "F6:K405"
    ClearRange "Incidents", "M6:O405"
    ClearRange "Memos", "C6:D305": ClearRange "Memos", "F6:G305"
    ClearRange "Memos", "I6:I305": ClearRange "Memos", "K6:K305"
    ClearRange "Memos", "M6:N305"
    ClearRange "DailyLog", "B6:B175": ClearRange "DailyLog", "E6:H175"
    ClearRange "DailyLog", "L6:N175"
    ' board meeting list persists across academies; only this academy's
    ' alignment answers reset (re-asked by the startup review below)
    ClearRange "AdvisoryBoard", "D7:D9"
    ClearRange "Counseling", "C6:D405": ClearRange "Counseling", "F6:L405"
    ClearRange "Counseling", "M6:N405"
    ClearRange "PT", "D6:Z55"
    ClearRange "Medical", "C6:D205": ClearRange "Medical", "F6:L205"
    ClearRange "Medical", "N6:N205"
    ClearRange "Certifications", "D6:S55"
    ClearRange "StateExam", "D6:I55": ClearRange "StateExam", "L6:L55"
    ' per-academy instructor picks cleared; the certified banks persist
    ClearRange "InstructorBanks", "M6:T105"
    ClearRange "DismissalLog", "C6:C105": ClearRange "DismissalLog", "E6:O105"
    ClearRange "EmailLog", "B6:I505"
    ClearAuditDocs
    If clearSched Then
        ClearRange "Schedule", "B6:B905": ClearRange "Schedule", "D6:E905"
        ClearRange "Schedule", "G6:G905": ClearRange "Schedule", "I6:J905"
        ClearRange "Schedule", "M6:M905"
    End If
    SetName "cfgCurrentExamNum", 1
    SetName "cfgCurrentSpellingNum", 1
    Application.ScreenUpdating = True
    MsgBox "Reset complete. Update Settings (class label, dates) and enter " & _
           "the new roster.", vbInformation, "New Academy"
    AcademyStartupReview
End Sub

' ==========================================================================
' Academy Startup Review - governance alignment prompts. Runs automatically
' after New Academy Reset; also available from its own button any time.
' Records the latest advisory-board meeting (running list), whether rules
' changed, and confirms this workbook was aligned to the current policies.
' ==========================================================================
Public Sub AcademyStartupReview()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("AdvisoryBoard")
    On Error Resume Next
    ws.Unprotect PW
    On Error GoTo 0

    ' 1) latest advisory board meeting -> running list (if new)
    Dim dTxt As String, d As Date, haveDate As Boolean
    dTxt = InputBox("Date of the MOST RECENT Advisory Board meeting " & _
                    "(mm/dd/yyyy)?" & vbCrLf & "(Leave blank to skip)", _
                    "Academy Startup Review")
    If IsDate(dTxt) Then
        d = CDate(dTxt): haveDate = True
        Dim r As Long, exists As Boolean, firstEmpty As Long
        For r = 12 To 41
            If IsDate(ws.Cells(r, "B").Value) Then
                If CDate(ws.Cells(r, "B").Value) = d Then exists = True
            ElseIf firstEmpty = 0 Then
                firstEmpty = r
            End If
        Next r
        If Not exists And firstEmpty > 0 Then
            ws.Cells(firstEmpty, "B").Value = d
            ws.Cells(firstEmpty, "C").Value = _
                InputBox("Server folder where those minutes live?", _
                         "Academy Startup Review")
            If MsgBox("Did that meeting change any rules or procedures " & _
                      "affecting this academy?", vbYesNo + vbQuestion, _
                      "Academy Startup Review") = vbYes Then
                ws.Cells(firstEmpty, "E").Value = "Yes"
                ws.Cells(firstEmpty, "F").Value = _
                    InputBox("Briefly: what changed, and what was updated " & _
                             "in this workbook to match?", _
                             "Academy Startup Review")
            Else
                ws.Cells(firstEmpty, "E").Value = "No"
            End If
            ws.Cells(firstEmpty, "H").Value = Environ$("USERNAME")
            ws.Cells(firstEmpty, "I").Value = Date
        End If
    End If

    ' 2) policy manual version for THIS academy
    Dim pv As String
    pv = InputBox("Policy manual version/date in effect for THIS academy?" & _
                  vbCrLf & "(e.g. 'May 2026' - rules updated between " & _
                  "academies must be reflected in Settings, ChapterMaster, " & _
                  "WritingMaster and SpellingMaster)", _
                  "Academy Startup Review", ws.Range("D6").Value)
    If pv <> "" Then ws.Range("D6").Value = pv

    ' 3) alignment confirmations (drive the audit checks)
    ws.Range("D7").Value = IIf(haveDate, "Yes", "No")
    If MsgBox("Confirm: rule/procedure changes (board + academy policy) " & _
              "have been reviewed and this workbook's settings, hours, " & _
              "prompts and lists match the current rules?", _
              vbYesNo + vbQuestion, "Academy Startup Review") = vbYes Then
        ws.Range("D8").Value = "Yes"
    Else
        ws.Range("D8").Value = "No"
        MsgBox "The Audit sheet will show 'Governance alignment' as CHECK " & _
               "until the workbook is aligned and this review is re-run.", _
               vbExclamation, "Academy Startup Review"
    End If
    ws.Range("D9").Value = Environ$("USERNAME") & " " & Format$(Date, "mm/dd/yyyy")
    MsgBox "Startup review recorded on the AdvisoryBoard sheet.", _
           vbInformation, "Academy Startup Review"
End Sub

Private Sub ClearAuditDocs()
    ' per-cadet enrollment docs grid + program checklist answers
    Dim ws As Worksheet, c As Range, firstDoc As Long, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Audit")
    If ws Is Nothing Then Exit Sub
    ws.Unprotect PW
    ' find the docs header by its "Enroll App" label, clear the 50-row grid
    For r = 1 To 120
        If InStr(1, CStr(ws.Cells(r, 3).Value), "Enroll App", vbTextCompare) > 0 Then
            ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 50, 9)).ClearContents
            Exit For
        End If
    Next r
    On Error GoTo 0
End Sub

Private Sub ClearRange(sheetName As String, addr As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Unprotect PW
    ws.Range(addr).ClearContents
    On Error GoTo 0
End Sub

Private Sub SetName(nm As String, v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
    On Error GoTo 0
End Sub
