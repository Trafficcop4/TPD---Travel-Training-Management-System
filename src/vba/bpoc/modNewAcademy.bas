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
    ClearRange "ExamScores", "Y6:Y1505"
    ClearRange "Spelling", "D6:O55"
    ClearRange "Attendance", "C6:D805": ClearRange "Attendance", "G6:J805"
    ClearRange "Attendance", "L6:M805": ClearRange "Attendance", "O6:O805"
    ClearRange "Makeup", "C6:D505": ClearRange "Makeup", "F6:K505"
    ClearRange "Makeup", "M6:M505"
    ClearRange "Skills", "C6:C605": ClearRange "Skills", "E6:E605"
    ClearRange "Skills", "I6:M605": ClearRange "Skills", "Q6:R605"
    ' SkillsCheck: the coordinator's Pass/Fail marks (D..M). B/C mirror the
    ' roster and N/O/P are formulas, so only the ten input columns are cleared.
    ClearRange "SkillsCheck", "D6:M55"
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
    ' per-academy instructor picks cleared; the certified banks persist.
    ' M..V = the SEL_SLOTS "Teach" columns in build/bpoc/sheets_config.py -
    ' keep this range in step with that constant.
    ClearRange "InstructorBanks", "M6:V105"
    ' per-chapter TRAINING FILE evidence (sign-in sheets, assessments, grade
    ' sheets, evals, special-req met) is this academy's audit record - it
    ' MUST reset or the Audit sheet reports the new class's chapter files
    ' complete with zero evidence collected. K = default instructor and
    ' S = the TCOLE special-requirement text are master data; G/H/I/J/U are
    ' formulas. Rows 6..49 = the 44 TCOLE chapters (nrCHfileOK's span).
    ClearRange "ChapterMaster", "L6:R49"
    ClearRange "ChapterMaster", "T6:T49"
    ' WritingMaster I/J are LAST academy's hand-entered handout dates, and
    ' K = IF(I<>"",I,G) / L = IF(J<>"",J,H) means a surviving override
    ' permanently beats the new class's computed dates on rngWMassigned /
    ' rngWMdue - every cadet would read 40 overdue writing assignments from
    ' day one and every agency email would say so. Rows 6..45 = the 40
    ' assignments (rngWMnum's span); G/H/K/L are formulas - never clear
    ' G6:L45 as one block.
    ClearRange "WritingMaster", "I6:J45"
    ' Control F = this academy's Extra Closure Dates. A stale one silently
    ' deletes a class day from the NEXT academy's generated calendar whenever
    ' the two date ranges overlap. F6:F20 = the 15-row input block; G is the
    ' Closure Check formula.
    ClearRange "Control", "F6:F20"
    ' E:P - P is "Closes Trigger", added when the DismissalLog became the
    ' authoritative way to CLOSE an engine-raised dismissal review
    ClearRange "DismissalLog", "C6:C105": ClearRange "DismissalLog", "E6:P105"
    ClearRange "EmailLog", "B6:I505"
    ' award overrides + notes are LAST academy's decisions, and F = IF(E<>"",
    ' E, C) means a surviving override permanently beats the new class's
    ' computed winner on nrAWfinal and the printed transcript. Column F is a
    ' formula - never clear E6:G9 as one block.
    ClearRange "sysAwards", "E6:E9": ClearRange "sysAwards", "G6:G9"
    ClearAuditDocs
    If clearSched Then
        ClearRange "Schedule", "B6:B905": ClearRange "Schedule", "D6:E905"
        ClearRange "Schedule", "G6:G905": ClearRange "Schedule", "I6:J905"
        ClearRange "Schedule", "M6:M905"
    End If
    SetName "cfgCurrentExamNum", 1
    SetName "cfgCurrentSpellingNum", 1
    Application.ScreenUpdating = True
    ' cfgTotalScheduledMinutes is deliberately NOT zeroed here. It no longer
    ' drives any cap - there is no classroom attendance allowance, per TCOLE
    ' ("there is no 10% attendance rule") and Academy policy 400 - but it is
    ' still the academy-length reference, and a zero would read as a real
    ' answer. Settings F/G flags a stale value on sight, and the message
    ' below sends the coordinator there.
    MsgBox "Reset complete. Update Settings — class label, start/end dates, " & _
           "and Total Scheduled Minutes (this academy's length; the Check " & _
           "column beside it will say CHECK until it matches the new " & _
           "Schedule) — then enter the new roster.", _
           vbInformation, "New Academy"
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
    Dim abWasProt As Boolean
    On Error Resume Next
    abWasProt = ws.ProtectContents
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
        If Not exists And firstEmpty = 0 Then
            MsgBox "The AdvisoryBoard meeting list is FULL (30 rows) - " & _
                   "this meeting was NOT recorded. Archive old rows and " & _
                   "re-run the Startup Review.", vbExclamation, _
                   "Academy Startup Review"
        End If
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
    ' the prompt above says "leave blank to skip", so a blank answer must not
    ' overwrite a "Yes" recorded on an earlier run - that flipped the
    ' governance-alignment audit check to CHECK even though the board minutes
    ' had been reviewed. Only write when we actually have a date, or when
    ' nothing has been recorded yet.
    If haveDate Then
        ws.Range("D7").Value = "Yes"
    ElseIf Trim$(CStr(ws.Range("D7").Value)) = "" Then
        ws.Range("D7").Value = "No"
    End If
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
    ' restore whatever protection state this sub found
    If abWasProt Then
        On Error Resume Next
        ws.Protect PW
        On Error GoTo 0
    End If
    MsgBox "Startup review recorded on the AdvisoryBoard sheet.", _
           vbInformation, "Academy Startup Review"
End Sub

Private Sub ClearAuditDocs()
    ' per-cadet enrollment docs grid + program checklist answers
    Dim ws As Worksheet, c As Range, firstDoc As Long, r As Long, rr As Long
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Audit")
    If ws Is Nothing Then Exit Sub
    ' Audit is a protected sheet with unlocked input cells - re-lock it when
    ' we are done, or the reset ships an editable audit packet
    wasProt = ws.ProtectContents
    ws.Unprotect PW
    ' find the docs header by its "Enroll App" label, clear the 50-row grid
    For r = 1 To 120
        If InStr(1, CStr(ws.Cells(r, 3).Value), "Enroll App", vbTextCompare) > 0 Then
            ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 50, 9)).ClearContents
            Exit For
        End If
    Next r
    ' find the program-requirements header by its "Met?" label and clear
    ' the previous academy's Met?/Notes answers (F/G) under it
    For r = 1 To 120
        If StrComp(Trim$(CStr(ws.Cells(r, 6).Value)), "Met?", vbTextCompare) = 0 Then
            rr = r + 1
            Do While Trim$(CStr(ws.Cells(rr, 2).Value)) <> ""
                ws.Range(ws.Cells(rr, 6), ws.Cells(rr, 7)).ClearContents
                rr = rr + 1
            Loop
            Exit For
        End If
    Next r
    If wasProt Then ws.Protect PW
    On Error GoTo 0
End Sub

Private Sub ClearRange(sheetName As String, addr As String)
    ' The sheet's protection state must SURVIVE the reset. This used to
    ' unprotect and never re-protect, so every New Academy Reset stripped
    ' sysAwards permanently: the computed-winner and FINAL formulas that feed
    ' the printed transcript were left freely editable, and that state was
    ' saved with the file. (Audit, PrintCenter and every other protected
    ' sheet this touches now come back locked too.)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    Dim wasProt As Boolean: wasProt = ws.ProtectContents
    ws.Unprotect PW
    ws.Range(addr).ClearContents
    If wasProt Then ws.Protect PW
    On Error GoTo 0
End Sub

Private Sub SetName(nm As String, v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
    On Error GoTo 0
End Sub
