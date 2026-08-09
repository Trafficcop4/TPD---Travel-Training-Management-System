Attribute VB_Name = "modPrint"
Option Explicit

' =====================================================================
'  PrintPacket - one PDF of the active record's paperwork:
'  Application, Pre 2-90, Notification - plus Post 2-90 automatically
'  once actuals have been entered (reconciliation stage).
'  Saved to the PacketFolder (Settings), then opened for printing.
' =====================================================================

Public Sub PrintPacket()
    Dim r As Long, path As String, hasActuals As Boolean
    r = RequireActiveRow(): If r = 0 Then Exit Sub
    hasActuals = Len(Trim$(CStr(RequestsSheet().Cells(r, COL_TOTALACT).Value))) > 0
    path = ExportPacketPDF(r, hasActuals)
    If Len(path) = 0 Then Exit Sub
    On Error Resume Next
    ThisWorkbook.FollowHyperlink path      ' open in the PDF viewer to print
    On Error GoTo 0
    MsgBox "Packet saved:" & vbCrLf & path, vbInformation, "Print Packet"
End Sub
