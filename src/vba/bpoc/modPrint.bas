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

Private Sub PrintSheet(nm As String)
    On Error GoTo Oops
    ThisWorkbook.Worksheets(nm).PrintOut
    Exit Sub
Oops:
    MsgBox "Could not print '" & nm & "': " & Err.Description, vbExclamation
End Sub
