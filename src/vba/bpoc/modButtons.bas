Attribute VB_Name = "modButtons"
Option Explicit

' ==========================================================================
' BPOC V6 - one-time button installer
' Run AddAllButtons once after installing the VBA modules. Safe to re-run
' (existing buttons are removed first).
' ==========================================================================

Public Sub AddAllButtons()
    AddButtons "PrintCenter", Array( _
        Array("Sign-In Sheet", "modPrint.btnPrintSignIn"), _
        Array("Sign-In Week (5 days)", "modPrint.btnPrintSignInWeek"), _
        Array("Academy Book (all days)", "modPrint.btnPrintSignInAcademy"), _
        Array("Eval Stack", "modPrint.btnPrintEvals"), _
        Array("Spelling Test/Key", "modPrint.btnPrintSpelling"), _
        Array("Writing Handout", "modPrint.btnPrintWriting"), _
        Array("Cadet Profile", "modPrint.btnPrintProfile"), _
        Array("Transcript(s)", "modPrint.btnPrintTranscript"), _
        Array("Ranking", "modPrint.btnPrintRanking"), _
        Array("Grad Checklist", "modPrint.btnPrintGradCheck"), _
        Array("Audit Packet", "modPrint.btnPrintAudit"), _
        Array("Chapter Packet", "modPrint.btnPrintChapterPacket"), _
        Array("Exam Grade Sheet", "modPrint.btnPrintGradeSheet"), _
        Array("Addendum Report", "modPrint.btnPrintAddendum"), _
        Array("Schedule", "modPrint.btnPrintSchedule"))
    AddButtons "Dashboard", Array( _
        Array("Agency Email Drafts", "modAgencyEmail.GenerateAgencyEmails"))
    AddButtons "EmailPreview", Array( _
        Array("Build Outlook Drafts", "modAgencyEmail.GenerateAgencyEmails"))
    AddButtons "StartHere", Array( _
        Array("New Academy Reset", "modNewAcademy.NewAcademyReset"))
    MsgBox "Buttons installed.", vbInformation
End Sub

Private Sub AddButtons(sheetName As String, defs As Variant)
    Dim ws As Worksheet, b As Object, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    ws.Unprotect "TPDAcademy"
    For Each b In ws.Buttons
        b.Delete
    Next b
    On Error GoTo 0
    Dim topPos As Double: topPos = 6
    For i = LBound(defs) To UBound(defs)
        Dim btn As Object
        Set btn = ws.Buttons.Add(ws.Columns("L").Left + 6, topPos, 130, 24)
        btn.Caption = defs(i)(0)
        btn.OnAction = defs(i)(1)
        topPos = topPos + 30
    Next i
End Sub
