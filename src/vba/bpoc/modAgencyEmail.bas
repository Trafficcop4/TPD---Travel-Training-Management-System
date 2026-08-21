Attribute VB_Name = "modAgencyEmail"
Option Explicit

' ==========================================================================
' BPOC V6 - Agency Score Emails (Outlook drafts)
' Same mechanics as V5:
'   * Reports the CURRENT exam + CURRENT spelling test per cadet.
'   * Spelling INCLUDED only when Current Spelling # >= Current Exam #.
'   * Outside agencies get only their own ACTIVE cadets; the home agency
'     (cfgHomeAgency) gets ONE draft with ALL active cadets by agency.
'   * Drafts are DISPLAYED for review - never auto-sent. Runs are logged.
' New in V6 (policy 300.6 reporting of deficiencies):
'   * Class average beside each score.
'   * RETEST note when the recorded score is a capped retake.
'   * Spelling early-intervention note when average < threshold (300.4.B).
'   * Attendance tier + writing-current status line per cadet.
'   * "Since your last report" digest of negative incidents & counseling
'     entries newer than the agency's last logged email.
' ==========================================================================

Private Const ROW1 As Long = 6          ' first data row everywhere
Private Const ROWN As Long = 55         ' last cadet row
Private Const AVG_ROW As Long = 57      ' ScoresGrid class-average row

Public Sub GenerateAgencyEmails()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsCad As Worksheet, wsAg As Worksheet, wsGrid As Worksheet
    Dim wsPlan As Worksheet, wsSpell As Worksheet
    Set wsCad = wb.Worksheets("Cadets")
    Set wsAg = wb.Worksheets("Agencies")
    Set wsGrid = wb.Worksheets("ScoresGrid")
    Set wsPlan = wb.Worksheets("ExamPlan")
    Set wsSpell = wb.Worksheets("Spelling")

    Dim academyClass As String, homeAgency As String
    Dim examNum As Long, spellNum As Long
    academyClass = CStr(NameVal(wb, "cfgAcademyClass"))
    homeAgency = Trim$(CStr(NameVal(wb, "cfgHomeAgency")))
    examNum = CLngSafe(NameVal(wb, "cfgCurrentExamNum"))
    spellNum = CLngSafe(NameVal(wb, "cfgCurrentSpellingNum"))
    If examNum <= 0 Then
        MsgBox "Set 'Current Exam #' on the Settings sheet first.", vbExclamation
        Exit Sub
    End If

    Dim omitSpelling As Boolean
    omitSpelling = (examNum > spellNum) Or (spellNum <= 0) Or (spellNum > 12)

    Dim examName As String
    Dim examCode As String: examCode = ExamCodeBySeq(wsPlan, examNum, examName)
    If examCode = "" Then
        MsgBox "No exam on the Exam Plan has Seq = " & examNum & ".", vbExclamation
        Exit Sub
    End If
    Dim examCol As Long: examCol = GridCol(wsGrid, examCode)
    If examCol = 0 Then
        MsgBox "Exam '" & examCode & "' was not found on ScoresGrid.", vbExclamation
        Exit Sub
    End If
    Dim spellCol As Long: spellCol = 0
    If Not omitSpelling Then spellCol = 3 + spellNum   ' Spelling: D=S01

    Dim olApp As Object: Set olApp = GetOutlookApp()
    If olApp Is Nothing Then
        MsgBox "Outlook could not be opened. Open Outlook and try again.", vbCritical
        Exit Sub
    End If

    Dim classAvg As Variant: classAvg = wsGrid.Cells(AVG_ROW, examCol).Value

    Dim agLast As Long: agLast = wsAg.Cells(wsAg.Rows.Count, "B").End(xlUp).Row
    Dim made As Long, agRow As Long, homeRow As Long
    Dim agID As String, agName As String, agEmail As String, active As String

    For agRow = ROW1 To agLast
        agID = Trim$(SafeStr(wsAg.Cells(agRow, "B").Value))
        agName = SafeStr(wsAg.Cells(agRow, "C").Value)
        agEmail = Trim$(SafeStr(wsAg.Cells(agRow, "E").Value))
        active = SafeStr(wsAg.Cells(agRow, "H").Value)
        If agID = "" Then GoTo NextAg
        If StrComp(agID, homeAgency, vbTextCompare) = 0 Then homeRow = agRow: GoTo NextAg
        If StrComp(Trim$(active), "Yes", vbTextCompare) <> 0 Then GoTo NextAg

        Dim body As String, n As Long, cutoff As Date
        cutoff = LastEmailDate(wb, agID)
        body = MailHead(academyClass, agName, examName, omitSpelling, spellNum, classAvg)
        body = body & AgencyTable(wb, agID, examCol, spellCol, omitSpelling, n)
        body = body & SinceLastSection(wb, agID, cutoff, False)
        body = body & MailFoot()
        If n > 0 Then
            MakeDraft olApp, agEmail, academyClass & " - Results (" & examName & _
                SpellSuffix(omitSpelling, spellNum) & ")", body
            made = made + 1
            LogRun wb, agID, examNum, spellNum, n, cutoff
        End If
NextAg:
    Next agRow

    If homeRow > 0 Then
        Dim hEmail As String, hName As String, allBody As String
        Dim total As Long, m As Long
        hEmail = Trim$(SafeStr(wsAg.Cells(homeRow, "E").Value))
        hName = SafeStr(wsAg.Cells(homeRow, "C").Value)
        Dim hCut As Date: hCut = LastEmailDate(wb, homeAgency)
        allBody = MailHead(academyClass, hName & " (all cadets, grouped by agency)", _
                           examName, omitSpelling, spellNum, classAvg)
        For agRow = ROW1 To agLast
            agID = Trim$(SafeStr(wsAg.Cells(agRow, "B").Value))
            agName = SafeStr(wsAg.Cells(agRow, "C").Value)
            ' NO ActiveYN filter here: the per-agency loop above already
            ' skips inactive agencies, so filtering again meant a cadet whose
            ' agency is not marked Active appeared in no draft at all - not
            ' even in the one titled "All Cadets".
            Dim sect As String
            If agID <> "" Then
                sect = AgencyTable(wb, agID, examCol, spellCol, omitSpelling, m)
                If m > 0 Then
                    allBody = allBody & "<h3 style='margin:16px 0 4px 0'>" & _
                              EscapeHtml(agName) & "</h3>" & sect
                    total = total + m
                End If
            End If
        Next agRow
        ' count active cadets whose agency produced no section at all
        Dim wsCadX As Worksheet: Set wsCadX = wb.Worksheets("Cadets")
        Dim rc As Long, orphan As Long, agOfCadet As String
        For rc = ROW1 To ROWN
            If SafeStr(wsCadX.Cells(rc, "B").Value) <> "" And _
               StrComp(Trim$(SafeStr(wsCadX.Cells(rc, "I").Value)), "Active", _
                       vbTextCompare) = 0 Then
                agOfCadet = Trim$(SafeStr(wsCadX.Cells(rc, "G").Value))
                If agOfCadet = "" Then orphan = orphan + 1
            End If
        Next rc
        allBody = allBody & SinceLastSection(wb, homeAgency, hCut, True)
        allBody = allBody & MailFoot()
        If total > 0 Then
            MakeDraft olApp, hEmail, academyClass & " - Results (" & examName & _
                SpellSuffix(omitSpelling, spellNum) & ") - All Cadets", allBody
            made = made + 1
            LogRun wb, homeAgency, examNum, spellNum, total, hCut
        End If
    End If

    MsgBox made & " draft email(s) opened in Outlook for review." & vbCrLf & _
           IIf(omitSpelling, "Spelling was OMITTED (spelling # < exam #).", _
               "Spelling #" & spellNum & " included.") & _
           IIf(orphan > 0, vbCrLf & orphan & " active cadet(s) have no " & _
               "AgencyID on Cadets and were reported to nobody.", ""), _
           vbInformation, "Agency Score Emails"
End Sub

' ---------- per-agency cadet table ----------
Private Function AgencyTable(wb As Workbook, agID As String, examCol As Long, _
        spellCol As Long, omitSpelling As Boolean, ByRef cnt As Long) As String
    Dim wsCad As Worksheet, wsGrid As Worksheet, wsSpell As Worksheet
    Dim wsAtt As Worksheet, wsWr As Worksheet, wsES As Worksheet
    Set wsCad = wb.Worksheets("Cadets"): Set wsGrid = wb.Worksheets("ScoresGrid")
    Set wsSpell = wb.Worksheets("Spelling")
    Set wsAtt = wb.Worksheets("sysAttendance"): Set wsWr = wb.Worksheets("Writing")
    Set wsES = wb.Worksheets("ExamScores")

    Dim spellThresh As Double
    spellThresh = CDblSafe(NameVal(wb, "cfgSpellInterventionAvg"), 75)
    Dim classAvg As Variant: classAvg = wsGrid.Cells(AVG_ROW, examCol).Value
    Dim r As Long, s As String
    cnt = 0
    s = "<table border='1' cellspacing='0' cellpadding='5' " & _
        "style='border-collapse:collapse;font-size:10pt'>" & _
        "<tr style='background:#1F3B5C;color:#fff'><th align='left'>Cadet</th>" & _
        "<th>Exam</th><th>Class Avg</th>"
    If Not omitSpelling Then s = s & "<th>Spelling</th>"
    s = s & "<th align='left'>Status</th></tr>"
    For r = ROW1 To ROWN
        If Trim$(SafeStr(wsCad.Cells(r, "B").Value)) = "" Then GoTo NextR
        If StrComp(Trim$(SafeStr(wsCad.Cells(r, "I").Value)), "Active", vbTextCompare) <> 0 Then GoTo NextR
        If StrComp(Trim$(SafeStr(wsCad.Cells(r, "G").Value)), agID, vbTextCompare) <> 0 Then GoTo NextR

        Dim nm As String, pid As String, ex As Variant, sp As Variant
        nm = SafeStr(wsCad.Cells(r, "F").Value)
        pid = SafeStr(wsCad.Cells(r, "B").Value)
        ex = wsGrid.Cells(r, examCol).Value
        s = s & "<tr><td>" & EscapeHtml(nm) & "</td>"
        s = s & "<td align='center' style='background:" & ScoreBg(ex) & "'>" & _
                EscapeHtml(FormatScore(ex)) & RetakeNote(wb, wsES, pid, examCol) & "</td>"
        s = s & "<td align='center'>" & EscapeHtml(FormatScore(classAvg)) & "</td>"
        If Not omitSpelling Then
            sp = wsSpell.Cells(r, spellCol).Value
            Dim spAvg As Variant: spAvg = wsSpell.Cells(r, "P").Value
            Dim spTxt As String: spTxt = FormatScore(sp)
            If IsNumeric(spAvg) Then
                If CDbl(spAvg) < spellThresh Then _
                    spTxt = spTxt & " (avg " & FormatScore(spAvg) & _
                            " - below " & Format$(spellThresh, "0") & _
                            ", early intervention per academy policy)"
            End If
            s = s & "<td align='center'>" & EscapeHtml(spTxt) & "</td>"
        End If
        s = s & "<td>" & EscapeHtml(StatusLine(wsAtt, wsWr, r)) & "</td></tr>"
        cnt = cnt + 1
NextR:
    Next r
    s = s & "</table>"
    AgencyTable = s
End Function

Private Function StatusLine(wsAtt As Worksheet, wsWr As Worksheet, r As Long) As String
    Dim t As String
    t = "Attendance " & SafeStr(wsAtt.Cells(r, "J").Value) & _
        " / PT " & SafeStr(wsAtt.Cells(r, "R").Value)
    Dim cur As String: cur = SafeStr(wsWr.Cells(r, "AT").Value)
    Dim od As Variant: od = wsWr.Cells(r, "AS").Value
    If cur = "Yes" Then
        t = t & "; writing current"
    ElseIf cur = "No" Then
        t = t & "; " & FormatScore(od) & " writing assignment(s) overdue"
    End If
    StatusLine = t
End Function

Private Function RetakeNote(wb As Workbook, wsES As Worksheet, pid As String, _
                            examCol As Long) As String
    ' the recorded grid score reflects the retake cap when attempt 2 exists
    Dim code As String
    code = SafeStr(wb.Worksheets("ScoresGrid").Cells(5, examCol).Value)
    Dim cnt As Double
    On Error Resume Next
    cnt = Application.WorksheetFunction.CountIfs( _
        wsES.Range("D6:D1505"), pid, wsES.Range("F6:F1505"), code, _
        wsES.Range("K6:K1505"), 2)
    On Error GoTo 0
    If cnt > 0 Then RetakeNote = " <i>(retest - recorded at cap)</i>"
End Function

' ---------- since-last-email digest ----------
Private Function SinceLastSection(wb As Workbook, agID As String, _
        cutoff As Date, isHome As Boolean) As String
    Dim wsIn As Worksheet, wsCo As Worksheet, wsCad As Worksheet
    Set wsIn = wb.Worksheets("Incidents"): Set wsCo = wb.Worksheets("Counseling")
    Set wsCad = wb.Worksheets("Cadets")
    Dim s As String, r As Long, cnt As Long
    Dim d As Variant, pid As String
    Dim wsMe As Worksheet: Set wsMe = wb.Worksheets("Memos")
    ' selective reporting: only rows the coordinator marked for the agency
    ' (Incidents "Report to Agency?" col O, Counseling "Agency Notified?"
    ' col J, Memos "Report to Agency?" col M). Everything else stays an
    ' academy teaching moment.
    s = "<h3 style='margin:16px 0 4px 0'>Reported items since your last report" & _
        IIf(cutoff > 0, " (" & Format$(cutoff, "mm/dd/yyyy") & ")", "") & "</h3>" & _
        "<table border='1' cellspacing='0' cellpadding='4' " & _
        "style='border-collapse:collapse;font-size:10pt'>" & _
        "<tr style='background:#44607E;color:#fff'><th>Date</th><th align='left'>Cadet</th>" & _
        "<th>Type</th><th align='left'>Description</th></tr>"
    For r = ROW1 To 405
        d = wsIn.Cells(r, "C").Value
        pid = SafeStr(wsIn.Cells(r, "E").Value)
        If IsDate(d) And pid <> "" Then
            If CDate(d) >= cutoff And _
               StrComp(SafeStr(wsIn.Cells(r, "O").Value), "Yes", vbTextCompare) = 0 And _
               (isHome Or AgencyOfPID(wsCad, pid) = agID) Then
                s = s & "<tr><td>" & Format$(CDate(d), "mm/dd") & "</td><td>" & _
                    EscapeHtml(NameOfPID(wsCad, pid)) & "</td><td>Incident (" & _
                    EscapeHtml(SafeStr(wsIn.Cells(r, "G").Value)) & ")</td><td>" & _
                    EscapeHtml(SafeStr(wsIn.Cells(r, "I").Value)) & "</td></tr>"
                cnt = cnt + 1
            End If
        End If
    Next r
    For r = ROW1 To 405
        d = wsCo.Cells(r, "C").Value
        pid = SafeStr(wsCo.Cells(r, "E").Value)
        If IsDate(d) And pid <> "" Then
            If CDate(d) >= cutoff And _
               StrComp(SafeStr(wsCo.Cells(r, "J").Value), "Yes", vbTextCompare) = 0 And _
               (isHome Or AgencyOfPID(wsCad, pid) = agID) Then
                s = s & "<tr><td>" & Format$(CDate(d), "mm/dd") & "</td><td>" & _
                    EscapeHtml(NameOfPID(wsCad, pid)) & "</td><td>" & _
                    EscapeHtml(SafeStr(wsCo.Cells(r, "F").Value)) & "</td><td>" & _
                    EscapeHtml(SafeStr(wsCo.Cells(r, "H").Value)) & "</td></tr>"
                cnt = cnt + 1
            End If
        End If
    Next r
    For r = ROW1 To 305
        d = wsMe.Cells(r, "C").Value
        pid = SafeStr(wsMe.Cells(r, "E").Value)
        If IsDate(d) And pid <> "" Then
            If CDate(d) >= cutoff And _
               StrComp(SafeStr(wsMe.Cells(r, "M").Value), "Yes", vbTextCompare) = 0 And _
               (isHome Or AgencyOfPID(wsCad, pid) = agID) Then
                s = s & "<tr><td>" & Format$(CDate(d), "mm/dd") & "</td><td>" & _
                    EscapeHtml(NameOfPID(wsCad, pid)) & "</td><td>Memo (" & _
                    EscapeHtml(SafeStr(wsMe.Cells(r, "L").Value)) & ")</td><td>" & _
                    EscapeHtml(SafeStr(wsMe.Cells(r, "G").Value)) & "</td></tr>"
                cnt = cnt + 1
            End If
        End If
    Next r
    s = s & "</table>"
    If cnt = 0 Then
        SinceLastSection = "<p style='font-size:10pt'><i>No items were marked " & _
            "for reporting since your last report.</i></p>"
    Else
        SinceLastSection = s
    End If
End Function

Private Function AgencyOfPID(wsCad As Worksheet, pid As String) As String
    Dim r As Long
    For r = ROW1 To ROWN
        If SafeStr(wsCad.Cells(r, "B").Value) = pid Then
            AgencyOfPID = Trim$(SafeStr(wsCad.Cells(r, "G").Value)): Exit Function
        End If
    Next r
End Function

Private Function NameOfPID(wsCad As Worksheet, pid As String) As String
    Dim r As Long
    For r = ROW1 To ROWN
        If SafeStr(wsCad.Cells(r, "B").Value) = pid Then
            NameOfPID = SafeStr(wsCad.Cells(r, "F").Value): Exit Function
        End If
    Next r
    NameOfPID = pid
End Function

Private Function LastEmailDate(wb As Workbook, agID As String) As Date
    Dim ws As Worksheet: Set ws = wb.Worksheets("EmailLog")
    Dim r As Long, d As Variant, best As Date
    For r = ROW1 To 505
        If StrComp(Trim$(SafeStr(ws.Cells(r, "C").Value)), agID, vbTextCompare) = 0 Then
            d = ws.Cells(r, "B").Value
            If IsDate(d) Then If CDate(d) > best Then best = CDate(d)
        End If
    Next r
    ' EmailLog!B is stamped with Now (date AND time), but Incidents,
    ' Counseling and Memos all hold DATE-ONLY values. Comparing a date-only
    ' entry against a mid-afternoon timestamp with ">" dropped every item
    ' dated on a run day but entered (or flagged "Report to Agency") after
    ' that run - permanently, from that digest and every later one. Int()
    ' makes the cutoff day-granular and the filters use ">=" to match; worst
    ' case an item is listed once more, instead of being lost silently.
    ' EmailPreview!B62 applies the identical INT()/>= rule - keep them in
    ' lockstep or the preview stops matching the draft it claims to show.
    LastEmailDate = Int(best)
End Function

' ---------- lookups ----------
Private Function ExamCodeBySeq(wsPlan As Worksheet, seq As Long, _
                               ByRef examName As String) As String
    Dim r As Long
    For r = ROW1 To 30
        If CLngSafe(wsPlan.Cells(r, "G").Value) = seq And _
           Trim$(SafeStr(wsPlan.Cells(r, "C").Value)) <> "" Then
            examName = SafeStr(wsPlan.Cells(r, "D").Value)
            ExamCodeBySeq = Trim$(SafeStr(wsPlan.Cells(r, "C").Value))
            Exit Function
        End If
    Next r
End Function

Private Function GridCol(wsGrid As Worksheet, code As String) As Long
    Dim c As Long
    For c = 4 To 28
        If StrComp(Trim$(SafeStr(wsGrid.Cells(5, c).Value)), code, vbTextCompare) = 0 Then
            GridCol = c: Exit Function
        End If
    Next c
End Function

' ---------- html helpers ----------
Private Function MailHead(cls As String, agName As String, examName As String, _
        omit As Boolean, spellNum As Long, classAvg As Variant) As String
    Dim h As String
    h = "<html><body style='font-family:Calibri,Arial;font-size:11pt'>" & _
        "<p>To the training liaison at <b>" & EscapeHtml(agName) & "</b>,</p>" & _
        "<p>Below are the results for <b>" & EscapeHtml(examName) & "</b> in <b>" & _
        EscapeHtml(cls) & "</b>"
    If Not omit Then h = h & " (with Spelling Test #" & spellNum & ")"
    h = h & ". Class average: <b>" & EscapeHtml(FormatScore(classAvg)) & _
        "</b>. Recorded scores reflect academy retest rules (passed retests " & _
        "record at 70).</p>"
    If omit Then h = h & "<p><i>Spelling scores are omitted this run because " & _
        "the spelling test number is behind the exam number.</i></p>"
    MailHead = h
End Function

Private Function MailFoot() As String
    MailFoot = "<p style='color:#666;font-size:10pt'>Draft generated for " & _
        "review. Please contact the Training Coordinator with any questions.</p>" & _
        "</body></html>"
End Function

Private Function SpellSuffix(omit As Boolean, n As Long) As String
    If Not omit Then SpellSuffix = " / Spelling #" & n
End Function

Private Function ScoreBg(v As Variant) As String
    If IsNumeric(v) Then
        If CDbl(v) < 70 Then ScoreBg = "#FFC7CE" Else ScoreBg = "#C6EFCE"
    Else
        ScoreBg = "#FFFFFF"
    End If
End Function

Private Function FormatScore(v As Variant) As String
    If IsError(v) Then
        FormatScore = ""
    ElseIf IsNumeric(v) Then
        FormatScore = Format$(CDbl(v), "0.##")
    Else
        FormatScore = Trim$(SafeStr(v))
    End If
End Function

Private Function EscapeHtml(ByVal t As String) As String
    t = Replace(t, "&", "&amp;"): t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;"): t = Replace(t, """", "&quot;")
    EscapeHtml = t
End Function

' ---------- misc ----------
Private Function NameVal(wb As Workbook, nm As String) As Variant
    On Error Resume Next
    NameVal = wb.Names(nm).RefersToRange.Value
    On Error GoTo 0
End Function

Private Function CLngSafe(v As Variant) As Long
    If IsNumeric(v) Then CLngSafe = CLng(v)
End Function

Private Function CDblSafe(v As Variant, dflt As Double) As Double
    If IsNumeric(v) Then CDblSafe = CDbl(v) Else CDblSafe = dflt
End Function

Private Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    If GetOutlookApp Is Nothing Then Set GetOutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

Private Sub MakeDraft(olApp As Object, toAddr As String, subj As String, html As String)
    Dim m As Object: Set m = olApp.CreateItem(0)   ' olMailItem
    m.To = toAddr
    m.Subject = subj
    m.HTMLBody = html
    m.Display                                       ' review only - never .Send
End Sub

Private Sub LogRun(wb As Workbook, agID As String, examNum As Long, _
                   spellNum As Long, cnt As Long, cutoff As Date)
    Dim ws As Worksheet: Set ws = wb.Worksheets("EmailLog")
    Dim r As Long
    For r = ROW1 To 505
        If Trim$(SafeStr(ws.Cells(r, "B").Value)) = "" Then Exit For
    Next r
    If r > 505 Then
        MsgBox "EmailLog is full (rows 6-505) - this run was NOT logged. " & _
               "Archive or clear old rows so future digests use the right " & _
               "'since last email' cutoff.", vbExclamation, "Agency Score Emails"
        Exit Sub
    End If
    ws.Cells(r, "B").Value = Now
    ws.Cells(r, "C").Value = agID
    ws.Cells(r, "D").Value = examNum
    ws.Cells(r, "E").Value = spellNum
    ws.Cells(r, "F").Value = cnt
    ws.Cells(r, "G").Value = IIf(cutoff > 0, Format$(cutoff, "mm/dd/yyyy"), "start of academy")
    ws.Cells(r, "H").Value = Environ$("USERNAME")
End Sub

Private Function SafeStr(v As Variant) As String
    On Error Resume Next
    If IsError(v) Then SafeStr = "" Else SafeStr = CStr(v)
    On Error GoTo 0
End Function
