Attribute VB_Name = "modCalendar"
Option Explicit

' =====================================================================
'  PushDeadlinesToOutlook - puts each record's two deadlines on the
'  Outlook calendar:
'    - Finance packet due: the Tuesday, at NOON (checks are cut that week)
'    - Receipts due: (return + ReceiptDays), 9:00 AM
'  Only NEW items are pushed: a record is skipped once its
'  "Cal Pushed?" column reads Y, so re-running never duplicates.
' =====================================================================

Private Const olAppointmentItem As Long = 1

Public Sub PushDeadlinesToOutlook()
    Dim ws As Worksheet, ol As Object, r As Long
    Dim pushed As Long, status As String

    Set ws = RequestsSheet()
    On Error GoTo NoOutlook
    Set ol = OutlookApp()
    On Error GoTo 0

    For r = FIRST_ROW To LAST_ROW
        If Len(Trim$(CStr(ws.Cells(r, COL_TRK).Value))) > 0 Then
            status = CStr(ws.Cells(r, COL_STATUS).Value)
            If ws.Cells(r, COL_CALPUSHED).Value <> "Y" _
               And status <> "Denied" And status <> "Cancelled" Then
                If PushOne(ol, ws, r) Then
                    ws.Cells(r, COL_CALPUSHED).Value = "Y"
                    pushed = pushed + 1
                End If
            End If
        End If
    Next r

    MsgBox pushed & " record(s) pushed to the Outlook calendar." & vbCrLf & _
           "Records already marked Y in 'Cal Pushed?' were left alone.", _
           vbInformation, "Calendar"
    Exit Sub

NoOutlook:
    MsgBox "Could not reach Outlook (" & Err.Description & ").", _
           vbExclamation, "Calendar"
End Sub

Private Function PushOne(ByVal ol As Object, ByVal ws As Worksheet, _
                         ByVal r As Long) As Boolean
    Dim trk As String, who As String, made As Boolean
    trk = CStr(ws.Cells(r, COL_TRK).Value)
    who = CStr(ws.Cells(r, COL_OFFICER).Value)

    If IsDate(ws.Cells(r, COL_FINANCEDUE).Value) Then
        MakeAppt ol, CDate(ws.Cells(r, COL_FINANCEDUE).Value) + TimeSerial(12, 0, 0), _
                 "FINANCE PACKET DUE (Tue noon): " & trk & " " & who, _
                 "Packet must reach Finance by NOON today for a check this week. " & _
                 "Record " & trk & " - " & CStr(ws.Cells(r, COL_COURSE).Value)
        made = True
    End If
    If IsDate(ws.Cells(r, COL_RECEIPTSDUE).Value) Then
        MakeAppt ol, CDate(ws.Cells(r, COL_RECEIPTSDUE).Value) + TimeSerial(9, 0, 0), _
                 "Receipts due: " & trk & " " & who, _
                 "Itemized receipts due from " & who & " for " & _
                 CStr(ws.Cells(r, COL_COURSE).Value) & " (" & trk & ")."
        made = True
    End If
    PushOne = made
End Function

Private Sub MakeAppt(ByVal ol As Object, ByVal startAt As Date, _
                     ByVal subj As String, ByVal body As String)
    Dim appt As Object
    Set appt = ol.CreateItem(olAppointmentItem)
    appt.Subject = subj
    appt.Start = startAt
    appt.Duration = 30
    appt.Body = body
    appt.ReminderSet = True
    appt.ReminderMinutesBeforeStart = 120
    appt.BusyStatus = 0                   ' olFree - it's a reminder, not a meeting
    appt.Save
End Sub
