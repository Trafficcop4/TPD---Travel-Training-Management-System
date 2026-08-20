Attribute VB_Name = "modPrint"
Option Explicit

' ==========================================================================
' BPOC V6 - Print Center
' Each button prints its sheet's defined print area. "Stack" variants loop.
' All printing goes through PrintPreview=False PrintOut to the default
' printer; use Ctrl+P on the sheet for printer/PDF choices.
' ==========================================================================

Private Const ROW1 As Long = 6
Private Const ROWN As Long = 55

Public Sub btnPrintSignIn()
    PrintSheet "SignIn"
End Sub

' prints the next 5 class days' daily report/roster forms, each with its
' own date, day # and schedule pre-filled (from the Control class-day
' calendar, starting at the SignIn sheet's date or today)
Public Sub btnPrintSignInWeek()
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Worksheets("SignIn")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets("Control")
    Dim startDate As Date
    If IsDate(wsS.Range("C5").Value) Then
        startDate = CDate(wsS.Range("C5").Value)
    Else
        startDate = Date
    End If
    Dim keep As Variant: keep = wsS.Range("C5").Value
    wsS.Unprotect "TPDAcademy"
    Dim r As Long, printed As Long, d As Variant
    For r = 6 To 175
        d = wsC.Cells(r, "I").Value
        If IsDate(d) Then
            If CDate(d) >= startDate And printed < 5 Then
                If wsC.Cells(r, "K").Value = "Yes" Then   ' in session
                    wsS.Range("C5").Value = CDate(d)
                    Application.Calculate
                    wsS.PrintOut
                    printed = printed + 1
                End If
            End If
        End If
        If printed >= 5 Then Exit For
    Next r
    wsS.Range("C5").Value = keep
    wsS.Protect "TPDAcademy"
    MsgBox printed & " daily report form(s) printed (next class days from " & _
           Format$(startDate, "mm/dd/yyyy") & ").", vbInformation, "Sign-In Week"
End Sub

Public Sub btnPrintSpelling()
    PrintSheet "SpellingPrint"
End Sub

Public Sub btnPrintWriting()
    PrintSheet "WritingHandout"
End Sub

Public Sub btnPrintProfile()
    PrintSheet "CadetProfile"
End Sub

Public Sub btnPrintRanking()
    PrintSheet "Ranking"
End Sub

Public Sub btnPrintGradCheck()
    PrintSheet "GradChecklist"
End Sub

Public Sub btnPrintAudit()
    PrintSheet "Audit"
End Sub

Public Sub btnPrintAddendum()
    PrintSheet "Addendum"
End Sub

Public Sub btnPrintChapterPacket()
    PrintSheet "ChapterPacket"
End Sub

Public Sub btnPrintGradeSheet()
    PrintSheet "ExamSheet"
End Sub

Public Sub btnPrintSchedule()
    PrintSheet "Schedule"
End Sub

' one critique per ACTIVE cadet (same blank form; prints N copies)
Public Sub btnPrintEvals()
    Dim n As Long, r As Long
    Dim wsCad As Worksheet: Set wsCad = ThisWorkbook.Worksheets("Cadets")
    For r = ROW1 To ROWN
        If Trim$(CStr(wsCad.Cells(r, "B").Value)) <> "" And _
           StrComp(CStr(wsCad.Cells(r, "I").Value), "Active", vbTextCompare) = 0 Then n = n + 1
    Next r
    If n = 0 Then MsgBox "No active cadets on the roster.", vbExclamation: Exit Sub
    If MsgBox("Print " & n & " copies of the critique form for the selected " & _
              "chapter?", vbYesNo + vbQuestion, "Eval stack") <> vbYes Then Exit Sub
    ThisWorkbook.Worksheets("EvalSheet").PrintOut Copies:=n
End Sub

' transcript for the selected cadet, or the whole class
Public Sub btnPrintTranscript()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Yes = print transcript for EVERY cadet on the roster." & _
                 vbCrLf & "No = print only the cadet currently selected on " & _
                 "the Transcript sheet.", vbYesNoCancel + vbQuestion, "Transcripts")
    If ans = vbCancel Then Exit Sub
    Dim wsT As Worksheet: Set wsT = ThisWorkbook.Worksheets("Transcript")
    If ans = vbNo Then wsT.PrintOut: Exit Sub

    Dim wsCad As Worksheet: Set wsCad = ThisWorkbook.Worksheets("Cadets")
    Dim r As Long, nm As String, printed As Long
    Dim keep As Variant: keep = wsT.Range("C5").Value
    wsT.Unprotect "TPDAcademy"
    For r = ROW1 To ROWN
        nm = Trim$(CStr(wsCad.Cells(r, "F").Value))
        If nm <> "" Then
            wsT.Range("C5").Value = nm
            Application.Calculate
            wsT.PrintOut
            printed = printed + 1
        End If
    Next r
    wsT.Range("C5").Value = keep
    wsT.Protect "TPDAcademy"
    MsgBox printed & " transcript(s) sent to the printer.", vbInformation
End Sub

' prints the daily report/roster for EVERY in-session class day - the
' class leader's academy book, produced once at the start. Tip: set the
' default printer to "Microsoft Print to PDF" first to get a file for
' the print shop instead of 130 loose pages.
Public Sub btnPrintSignInAcademy()
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Worksheets("SignIn")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets("Control")
    Dim r As Long, total As Long, d As Variant
    For r = 6 To 175
        If IsDate(wsC.Cells(r, "I").Value) And _
           wsC.Cells(r, "K").Value = "Yes" Then total = total + 1
    Next r
    If total = 0 Then
        MsgBox "No class days found - set the Start/End dates on Settings " & _
               "first.", vbExclamation: Exit Sub
    End If
    If MsgBox("Print the daily report form for ALL " & total & " class " & _
              "days (the class leader's academy book)?" & vbCrLf & vbCrLf & _
              "Tip: set your default printer to 'Microsoft Print to PDF' " & _
              "first to get one file for binding.", _
              vbYesNo + vbQuestion, "Academy Book") <> vbYes Then Exit Sub
    Dim keep As Variant: keep = wsS.Range("C5").Value
    wsS.Unprotect "TPDAcademy"
    Application.ScreenUpdating = False
    Dim printed As Long
    For r = 6 To 175
        d = wsC.Cells(r, "I").Value
        If IsDate(d) And wsC.Cells(r, "K").Value = "Yes" Then
            wsS.Range("C5").Value = CDate(d)
            Application.Calculate
            wsS.PrintOut
            printed = printed + 1
        End If
    Next r
    Application.ScreenUpdating = True
    wsS.Range("C5").Value = keep
    wsS.Protect "TPDAcademy"
    MsgBox printed & " forms printed - the academy book is ready to bind." & _
           vbCrLf & "After a mid-academy separation, strike the name by " & _
           "hand on remaining pages or reprint the rest of the book.", _
           vbInformation, "Academy Book"
End Sub

Private Sub PrintSheet(nm As String)
    On Error GoTo Oops
    ThisWorkbook.Worksheets(nm).PrintOut
    Exit Sub
Oops:
    MsgBox "Could not print '" & nm & "': " & Err.Description, vbExclamation
End Sub
