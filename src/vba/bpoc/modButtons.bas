Attribute VB_Name = "modButtons"
Option Explicit

' ==========================================================================
' BPOC V6 - button installer (FALLBACK ONLY)
'
' tools\Install-BPOC-VBA.ps1 already places every button when it builds the
' .xlsm, and that script is the authority on where they go. Run AddAllButtons
' ONLY if you are wiring the macros into a workbook by hand (no PowerShell,
' no "Trust access to the VBA project object model"), or if the buttons were
' deleted from a live file.
'
' This module used to be a SECOND, CONTRADICTORY installer: it deleted the
' buttons the PowerShell script had placed and restacked all 15 Print Center
' buttons in a single column at column L, six rows above the table they label
' - exactly the misalignment the installer's own comment says was fixed. The
' anchors below are now the SAME anchors the installer uses; keep the two
' lists in step (Install-BPOC-VBA.ps1 -> Add-Button calls).
' ==========================================================================

Private Const PW As String = "TPDAcademy"
Private Const BTN_W As Double = 130     ' matches Add-Button's minimum width
Private Const BTN_H As Double = 22

Public Sub AddAllButtons()
    If MsgBox("Place the macro buttons on PrintCenter, Dashboard, " & _
              "EmailPreview and StartHere?" & vbCrLf & vbCrLf & _
              "Normally tools\Install-BPOC-VBA.ps1 does this when it builds " & _
              "the .xlsm - use this only when wiring the macros up by hand.", _
              vbYesNo + vbQuestion, "BPOC buttons") <> vbYes Then Exit Sub

    ' PrintCenter: ONE button per table row (the printables table is rows
    ' 8..22), same order as `rows` in build_printcenter (sheets_outputs.py)
    ClearButtons "PrintCenter"
    AddButtonAt "PrintCenter", "G8", "Sign-In Sheet", "modPrint.btnPrintSignIn"
    AddButtonAt "PrintCenter", "G9", "Sign-In Week", "modPrint.btnPrintSignInWeek"
    AddButtonAt "PrintCenter", "G10", "Academy Book", "modPrint.btnPrintSignInAcademy"
    AddButtonAt "PrintCenter", "G11", "Eval Stack", "modPrint.btnPrintEvals"
    AddButtonAt "PrintCenter", "G12", "Spelling Test/Key", "modPrint.btnPrintSpelling"
    AddButtonAt "PrintCenter", "G13", "Writing Handout", "modPrint.btnPrintWriting"
    AddButtonAt "PrintCenter", "G14", "Cadet Profile", "modPrint.btnPrintProfile"
    AddButtonAt "PrintCenter", "G15", "Transcript(s)", "modPrint.btnPrintTranscript"
    AddButtonAt "PrintCenter", "G16", "Ranking", "modPrint.btnPrintRanking"
    AddButtonAt "PrintCenter", "G17", "Grad Checklist", "modPrint.btnPrintGradCheck"
    AddButtonAt "PrintCenter", "G18", "Audit Packet", "modPrint.btnPrintAudit"
    AddButtonAt "PrintCenter", "G19", "Chapter Packet", "modPrint.btnPrintChapterPacket"
    AddButtonAt "PrintCenter", "G20", "Exam Grade Sheet", "modPrint.btnPrintGradeSheet"
    AddButtonAt "PrintCenter", "G21", "Addendum Report", "modPrint.btnPrintAddendum"
    AddButtonAt "PrintCenter", "G22", "Schedule", "modPrint.btnPrintSchedule"

    ClearButtons "Dashboard"
    AddButtonAt "Dashboard", "K5", "Agency Email Drafts", _
                "modAgencyEmail.GenerateAgencyEmails"

    ' L5, not H5: the E5 status line ("Exam # ... | Last emailed: ...") spills
    ' across F:J, and a button anchored at H5 sat on top of the cutoff date
    ' the whole preview sheet exists to show.
    ClearButtons "EmailPreview"
    AddButtonAt "EmailPreview", "L5", "Build Outlook Drafts", _
                "modAgencyEmail.GenerateAgencyEmails"

    ' D16 and H16: buttons are forced to a 130pt minimum width, and D..G on
    ' StartHere are default-width columns, so anchoring the second button at
    ' F16 (96pt to the right) drew it over the right ~34pt of the first.
    ClearButtons "StartHere"
    AddButtonAt "StartHere", "D16", "New Academy Reset", _
                "modNewAcademy.NewAcademyReset"
    AddButtonAt "StartHere", "H16", "Startup Review", _
                "modNewAcademy.AcademyStartupReview"

    MsgBox "Buttons installed.", vbInformation, "BPOC buttons"
End Sub

Private Sub ClearButtons(sheetName As String)
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    wasProt = ws.ProtectContents
    ws.Unprotect PW
    ' iterate DOWNWARD, never For Each: deleting members of the live
    ' Buttons collection while enumerating it skips alternate items, so a
    ' second AddAllButtons left about half the old buttons behind and
    ' stacked a full new set on top of them.
    Dim bi As Long
    For bi = ws.Buttons.Count To 1 Step -1
        ws.Buttons(bi).Delete
    Next bi
    ' restore the protection state this sub found - Dashboard, EmailPreview
    ' and PrintCenter must not be left open because buttons were refreshed
    If wasProt Then ws.Protect PW
    On Error GoTo 0
End Sub

Private Sub AddButtonAt(sheetName As String, anchorCell As String, _
                        caption As String, macroName As String)
    Dim ws As Worksheet, btn As Object, anchor As Range
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    wasProt = ws.ProtectContents
    ws.Unprotect PW
    Set anchor = ws.Range(anchorCell)
    Set btn = ws.Buttons.Add(anchor.Left, anchor.Top + 2, BTN_W, BTN_H)
    btn.Name = "btn_" & macroName
    btn.Caption = caption
    btn.OnAction = macroName
    If wasProt Then ws.Protect PW
    On Error GoTo 0
End Sub
