Attribute VB_Name = "modEmails"
Option Explicit

' =====================================================================
'  Email drafting. Every email OPENS FOR REVIEW (.Display) - this
'  system never sends mail on its own. Wording lives on Settings
'  (Email Templates section); {PLACEHOLDERS} fill from the record.
' =====================================================================

Public Sub DraftApprovalEmail()
    ' To the supervisor, with the printed application packet attached.
    Dim r As Long, pdf As String
    r = RequireActiveRow(): If r = 0 Then Exit Sub
    pdf = ExportPacketPDF(r, False)   ' pre-trip packet, no Post 2-90
    DraftMail _
        toAddr:=CStr(RequestsSheet().Cells(r, COL_SUPEMAIL).Value), _
        subj:=FillTemplate(CStr(Setting("ET_ApprSubj")), r), _
        body:=FillTemplate(CStr(Setting("ET_ApprBody")), r), _
        attachPath:=pdf
End Sub

Public Sub DraftOfficerNotification()
    ' To the traveling officer, with the notification sheet attached.
    Dim r As Long, pdf As String
    r = RequireActiveRow(): If r = 0 Then Exit Sub
    pdf = ExportSheetPDF(r, "Notification", " Notification")
    DraftMail _
        toAddr:=CStr(RequestsSheet().Cells(r, COL_EMAIL).Value), _
        subj:=FillTemplate(CStr(Setting("ET_NotifSubj")), r), _
        body:=FillTemplate(CStr(Setting("ET_NotifBody")), r), _
        attachPath:=pdf
End Sub

Public Sub DraftSettlementEmail()
    ' To the traveler after reconciliation: what is owed, either way.
    Dim r As Long
    r = RequireActiveRow(): If r = 0 Then Exit Sub
    If Len(Trim$(CStr(RequestsSheet().Cells(r, COL_TOTALACT).Value))) = 0 Then
        MsgBox "No actuals entered yet for " & CStr(Setting("ActiveTrk")) & _
               " - fill the Actual columns on Requests first.", _
               vbExclamation, "Settlement"
        Exit Sub
    End If
    DraftMail _
        toAddr:=CStr(RequestsSheet().Cells(r, COL_EMAIL).Value), _
        subj:=FillTemplate(CStr(Setting("ET_ReimbSubj")), r), _
        body:=FillTemplate(CStr(Setting("ET_ReimbBody")), r), _
        attachPath:=""
End Sub

Private Sub DraftMail(ByVal toAddr As String, ByVal subj As String, _
                      ByVal body As String, ByVal attachPath As String)
    Dim ol As Object, mail As Object
    On Error GoTo Fail
    Set ol = OutlookApp()
    Set mail = ol.CreateItem(0)          ' olMailItem
    mail.To = toAddr
    mail.Subject = subj
    mail.Body = body
    If Len(attachPath) > 0 Then
        If Dir(attachPath) <> "" Then mail.Attachments.Add attachPath
    End If
    mail.Display                          ' review, then YOU press Send
    Exit Sub
Fail:
    MsgBox "Could not open an Outlook draft (" & Err.Description & ")." & _
           vbCrLf & "Is Outlook installed and signed in?", vbExclamation, "Email"
End Sub

' ---------------------------------------------------------------------
'  PDF helpers (also used by modPrint)
' ---------------------------------------------------------------------

Public Function ExportPacketPDF(ByVal r As Long, ByVal includePost As Boolean) As String
    ' Application + Pre 2-90 (+ Post 2-90 when reconciling) + Notification,
    ' exported as one PDF named for the tracking number.
    Dim path As String, trk As String
    trk = CStr(RequestsSheet().Cells(r, COL_TRK).Value)
    path = PacketPath(trk, IIf(includePost, " Packet (final)", " Packet"))
    If path = "" Then Exit Function
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    If includePost Then
        ThisWorkbook.Sheets(Array("Application", "Pre 2-90", "Post 2-90", _
                                  "Notification")).Select
    Else
        ThisWorkbook.Sheets(Array("Application", "Pre 2-90", "Notification")).Select
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, _
        Quality:=xlQualityStandard, OpenAfterPublish:=False
    ThisWorkbook.Worksheets("Dashboard").Select
    Application.ScreenUpdating = True
    ExportPacketPDF = path
End Function

Public Function ExportSheetPDF(ByVal r As Long, ByVal sheetName As String, _
                               ByVal suffix As String) As String
    Dim path As String
    path = PacketPath(CStr(RequestsSheet().Cells(r, COL_TRK).Value), suffix)
    If path = "" Then Exit Function
    ThisWorkbook.Worksheets(sheetName).ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=path, Quality:=xlQualityStandard, OpenAfterPublish:=False
    ExportSheetPDF = path
End Function

Private Function PacketPath(ByVal trk As String, ByVal suffix As String) As String
    Dim folder As String
    folder = FolderSetting("PacketFolder")
    If Dir(folder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
        If Dir(folder, vbDirectory) = "" Then
            MsgBox "Packet folder not found and could not be created:" & vbCrLf & _
                   folder & vbCrLf & "Fix the path on Settings.", _
                   vbExclamation, "Packets"
            Exit Function
        End If
    End If
    PacketPath = folder & "\" & trk & suffix & ".pdf"
End Function
