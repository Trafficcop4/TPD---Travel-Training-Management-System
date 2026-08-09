Attribute VB_Name = "modUtil"
Option Explicit

' =====================================================================
'  TPD Training Master - shared helpers
'  Everything configurable is read from the Settings sheet BY NAME,
'  never by cell address, so Settings can be rearranged safely.
' =====================================================================

' Requests sheet column numbers. Must match the Requests sheet layout.
Public Const COL_TRK As Long = 1        ' A  Tracking #
Public Const COL_STATUS As Long = 2     ' B  Status
Public Const COL_OFFICER As Long = 3    ' C
Public Const COL_RANK As Long = 4       ' D
Public Const COL_EMPID As Long = 5      ' E
Public Const COL_DIVISION As Long = 6   ' F
Public Const COL_EMAIL As Long = 7      ' G
Public Const COL_PHONE As Long = 8      ' H
Public Const COL_SUPNAME As Long = 9    ' I
Public Const COL_SUPEMAIL As Long = 10  ' J
Public Const COL_COURSE As Long = 11    ' K
Public Const COL_PROVIDER As Long = 12  ' L
Public Const COL_LOCATION As Long = 13  ' M
Public Const COL_CLASSSTART As Long = 14 ' N
Public Const COL_CLASSEND As Long = 15  ' O
Public Const COL_DEPDATE As Long = 16   ' P
Public Const COL_DEPTIME As Long = 17   ' Q
Public Const COL_RETDATE As Long = 18   ' R
Public Const COL_RETTIME As Long = 19   ' S
Public Const COL_REGEST As Long = 20    ' T
Public Const COL_NIGHTS As Long = 21    ' U
Public Const COL_LODGERATE As Long = 22 ' V
Public Const COL_MODE As Long = 24      ' X
Public Const COL_MILES As Long = 25     ' Y
Public Const COL_TRAVELOTHER As Long = 26 ' Z
Public Const COL_PERDIEM As Long = 28   ' AB (formula)
Public Const COL_OTHEREST As Long = 29  ' AC
Public Const COL_TOTALEST As Long = 30  ' AD (formula)
Public Const COL_TOTALACT As Long = 36  ' AJ (formula)
Public Const COL_BALANCE As Long = 39   ' AM (formula)
Public Const COL_SUBMITTED As Long = 40 ' AN
Public Const COL_RECEIPTSDUE As Long = 44 ' AR (formula)
Public Const COL_FINANCEDUE As Long = 46  ' AT (formula)
Public Const COL_CALPUSHED As Long = 48   ' AV
Public Const COL_TCOLE As Long = 58     ' BF
Public Const COL_WEBSITE As Long = 59   ' BG
Public Const COL_OTHERDESC As Long = 60 ' BH
Public Const COL_JUSTIFY As Long = 61   ' BI

Public Const FIRST_ROW As Long = 2
Public Const LAST_ROW As Long = 251

Public Function Setting(ByVal settingName As String) As Variant
    ' A named cell on the Settings sheet (or Dashboard, e.g. ActiveTrk).
    Setting = ThisWorkbook.Names(settingName).RefersToRange.Value
End Function

Public Function RequestsSheet() As Worksheet
    Set RequestsSheet = ThisWorkbook.Worksheets("Requests")
End Function

Public Function ActiveTrkRow() As Long
    ' Row (on Requests) of the Dashboard's ACTIVE record; 0 if none picked.
    Dim trk As Variant, hit As Variant
    trk = Setting("ActiveTrk")
    If Len(Trim$(CStr(trk))) = 0 Then ActiveTrkRow = 0: Exit Function
    hit = Application.Match(trk, RequestsSheet().Range("A:A"), 0)
    If IsError(hit) Then ActiveTrkRow = 0 Else ActiveTrkRow = CLng(hit)
End Function

Public Function RequireActiveRow() As Long
    RequireActiveRow = ActiveTrkRow()
    If RequireActiveRow = 0 Then
        MsgBox "Pick a record first: Dashboard > ACTIVE (the yellow box)." & vbCrLf & _
               "Every email, form and packet works on that record.", _
               vbExclamation, "No active record"
    End If
End Function

Public Function NextRequestRow() As Long
    ' First empty row in the Requests table (Tracking # column blank).
    Dim ws As Worksheet, r As Long
    Set ws = RequestsSheet()
    For r = FIRST_ROW To LAST_ROW
        If Len(Trim$(CStr(ws.Cells(r, COL_TRK).Value))) = 0 Then
            NextRequestRow = r
            Exit Function
        End If
    Next r
    NextRequestRow = 0   ' table full - see Instructions sheet to extend
End Function

Public Function NextTrackingNumber() As String
    ' TrkPrefix + next 3-digit sequence, scanning existing tracking numbers.
    Dim ws As Worksheet, r As Long, prefix As String, s As String, n As Long, hi As Long
    Set ws = RequestsSheet()
    prefix = CStr(Setting("TrkPrefix"))
    hi = 0
    For r = FIRST_ROW To LAST_ROW
        s = Trim$(CStr(ws.Cells(r, COL_TRK).Value))
        If Len(s) > Len(prefix) Then
            If Left$(s, Len(prefix)) = prefix And IsNumeric(Mid$(s, Len(prefix) + 1)) Then
                n = CLng(Mid$(s, Len(prefix) + 1))
                If n > hi Then hi = n
            End If
        End If
    Next r
    NextTrackingNumber = prefix & Format$(hi + 1, "000")
End Function

Public Function FillTemplate(ByVal template As String, ByVal r As Long) As String
    ' Replace {PLACEHOLDERS} in an email template with record r's values.
    Dim ws As Worksheet, s As String
    Set ws = RequestsSheet()
    s = template
    s = Replace(s, "{TRK}", CStr(ws.Cells(r, COL_TRK).Value))
    s = Replace(s, "{OFFICER}", CStr(ws.Cells(r, COL_OFFICER).Value))
    s = Replace(s, "{RANK}", CStr(ws.Cells(r, COL_RANK).Value))
    s = Replace(s, "{SUPERVISOR}", CStr(ws.Cells(r, COL_SUPNAME).Value))
    s = Replace(s, "{COURSE}", CStr(ws.Cells(r, COL_COURSE).Value))
    s = Replace(s, "{PROVIDER}", CStr(ws.Cells(r, COL_PROVIDER).Value))
    s = Replace(s, "{LOCATION}", CStr(ws.Cells(r, COL_LOCATION).Value))
    s = Replace(s, "{DATES}", FmtDate(ws.Cells(r, COL_CLASSSTART).Value) & _
                              " - " & FmtDate(ws.Cells(r, COL_CLASSEND).Value))
    s = Replace(s, "{DEPART}", FmtDate(ws.Cells(r, COL_DEPDATE).Value) & " " & _
                               FmtTime(ws.Cells(r, COL_DEPTIME).Value))
    s = Replace(s, "{RETURN}", FmtDate(ws.Cells(r, COL_RETDATE).Value) & " " & _
                               FmtTime(ws.Cells(r, COL_RETTIME).Value))
    s = Replace(s, "{TOTAL}", FmtMoney(ws.Cells(r, COL_TOTALEST).Value))
    s = Replace(s, "{BALANCE}", FmtMoney(ws.Cells(r, COL_BALANCE).Value))
    s = Replace(s, "{RECEIPTS_DUE}", FmtDate(ws.Cells(r, COL_RECEIPTSDUE).Value))
    s = Replace(s, "{COORD}", CStr(Setting("CoordName")))
    s = Replace(s, "{PHONE}", CStr(Setting("CoordPhone")))
    FillTemplate = s
End Function

Public Function FmtDate(ByVal v As Variant) As String
    If IsDate(v) Then FmtDate = Format$(CDate(v), "mm/dd/yyyy") Else FmtDate = ""
End Function

Public Function FmtTime(ByVal v As Variant) As String
    If IsDate(v) Or IsNumeric(v) Then
        If Len(Trim$(CStr(v))) > 0 Then FmtTime = Format$(CDate(v), "h:mm AM/PM")
    End If
End Function

Public Function FmtMoney(ByVal v As Variant) As String
    If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
        FmtMoney = Format$(CDbl(v), "$#,##0.00")
    Else
        FmtMoney = "(not yet calculated)"
    End If
End Function

Public Function OutlookApp() As Object
    ' Running Outlook if there is one, else start it.
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
End Function

Public Function FolderSetting(ByVal settingName As String) As String
    ' A folder path from Settings, without trailing backslash; "" if unset.
    Dim p As String
    p = Trim$(CStr(Setting(settingName)))
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    FolderSetting = p
End Function
