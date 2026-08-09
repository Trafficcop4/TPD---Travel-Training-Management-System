Attribute VB_Name = "modImport"
Option Explicit

' =====================================================================
'  ImportRequests - one click pulls in every officer request form
'  sitting in the ImportFolder (Settings), assigns each a permanent
'  tracking number, writes it to the Requests table, and moves the
'  file to the ProcessedFolder renamed to its tracking number.
'
'  Officer forms are read by FIELD NAME (the rq_* named ranges on the
'  request form), never by cell address, so the form layout can evolve
'  without breaking the import.
' =====================================================================

Public Sub ImportRequests()
    Dim inbox As String, done As String, f As String
    Dim imported As Long, skipped As String

    inbox = FolderSetting("ImportFolder")
    done = FolderSetting("ProcessedFolder")
    If Dir(inbox, vbDirectory) = "" Then
        MsgBox "Import folder not found:" & vbCrLf & inbox & vbCrLf & vbCrLf & _
               "Fix the path on the Settings sheet (Folders section).", _
               vbExclamation, "Import Requests"
        Exit Sub
    End If
    If Dir(done, vbDirectory) = "" Then MkDir done

    Application.ScreenUpdating = False
    f = Dir(inbox & "\*.xls*")
    Do While Len(f) > 0
        If ImportOneForm(inbox & "\" & f, done) Then
            imported = imported + 1
            f = Dir(inbox & "\*.xls*")   ' restart scan: the file moved
        Else
            skipped = skipped & vbCrLf & "  - " & f
            f = Dir()                    ' leave it in place, continue
        End If
    Loop
    Application.ScreenUpdating = True

    Dim msg As String
    msg = imported & " request(s) imported and numbered."
    If Len(skipped) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Skipped (not a readable request form " & _
              "or the table is full):" & skipped & vbCrLf & vbCrLf & _
              "Skipped files stay in the inbox folder."
    End If
    MsgBox msg, vbInformation, "Import Requests"
End Sub

Private Function ImportOneForm(ByVal path As String, ByVal doneFolder As String) As Boolean
    Dim src As Workbook, ws As Worksheet, r As Long, trk As String
    Dim destName As String

    On Error GoTo Fail
    r = NextRequestRow()
    If r = 0 Then Exit Function        ' table full - see Instructions sheet

    Set src = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    If Not HasName(src, "rq_Name") Then GoTo Fail   ' not one of our forms

    Set ws = RequestsSheet()
    trk = NextTrackingNumber()
    ws.Cells(r, COL_TRK).Value = trk
    ws.Cells(r, COL_STATUS).Value = "Awaiting Approval"
    ws.Cells(r, COL_SUBMITTED).Value = Date

    ws.Cells(r, COL_OFFICER).Value = NamedVal(src, "rq_Name")
    ws.Cells(r, COL_RANK).Value = NamedVal(src, "rq_Rank")
    ws.Cells(r, COL_EMPID).Value = NamedVal(src, "rq_EmpID")
    ws.Cells(r, COL_DIVISION).Value = NamedVal(src, "rq_Division")
    ws.Cells(r, COL_EMAIL).Value = NamedVal(src, "rq_Email")
    ws.Cells(r, COL_PHONE).Value = NamedVal(src, "rq_Phone")
    ws.Cells(r, COL_SUPNAME).Value = NamedVal(src, "rq_SupName")
    ws.Cells(r, COL_SUPEMAIL).Value = NamedVal(src, "rq_SupEmail")
    ws.Cells(r, COL_COURSE).Value = NamedVal(src, "rq_Course")
    ws.Cells(r, COL_PROVIDER).Value = NamedVal(src, "rq_Provider")
    ws.Cells(r, COL_TCOLE).Value = NamedVal(src, "rq_TCOLE")
    ws.Cells(r, COL_LOCATION).Value = NamedVal(src, "rq_Location")
    ws.Cells(r, COL_WEBSITE).Value = NamedVal(src, "rq_Website")
    ws.Cells(r, COL_CLASSSTART).Value = NamedVal(src, "rq_Start")
    ws.Cells(r, COL_CLASSEND).Value = NamedVal(src, "rq_End")
    ws.Cells(r, COL_DEPDATE).Value = NamedVal(src, "rq_DepDate")
    ws.Cells(r, COL_DEPTIME).Value = NamedVal(src, "rq_DepTime")
    ws.Cells(r, COL_RETDATE).Value = NamedVal(src, "rq_RetDate")
    ws.Cells(r, COL_RETTIME).Value = NamedVal(src, "rq_RetTime")
    ws.Cells(r, COL_REGEST).Value = NamedVal(src, "rq_Reg")
    ws.Cells(r, COL_NIGHTS).Value = NamedVal(src, "rq_Nights")
    ws.Cells(r, COL_LODGERATE).Value = NamedVal(src, "rq_LodgeRate")
    ws.Cells(r, COL_MODE).Value = NamedVal(src, "rq_Mode")
    ws.Cells(r, COL_MILES).Value = NamedVal(src, "rq_Miles")
    ws.Cells(r, COL_TRAVELOTHER).Value = NamedVal(src, "rq_TravelOther")
    ws.Cells(r, COL_OTHEREST).Value = NamedVal(src, "rq_Other")
    ws.Cells(r, COL_OTHERDESC).Value = NamedVal(src, "rq_OtherDesc")
    ws.Cells(r, COL_JUSTIFY).Value = NamedVal(src, "rq_Notes")

    src.Close SaveChanges:=False
    Set src = Nothing

    ' Keep the original, renamed to its tracking number, in Processed.
    destName = doneFolder & "\" & trk & " - " & Mid$(path, InStrRev(path, "\") + 1)
    If Dir(destName) <> "" Then Kill destName
    Name path As destName

    ImportOneForm = True
    Exit Function

Fail:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    ImportOneForm = False
End Function

Private Function HasName(ByVal wb As Workbook, ByVal n As String) As Boolean
    Dim t As String
    On Error Resume Next
    t = wb.Names(n).Name
    HasName = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function NamedVal(ByVal wb As Workbook, ByVal n As String) As Variant
    ' Value of a named cell in the officer's form; "" when absent/blank.
    On Error Resume Next
    NamedVal = wb.Names(n).RefersToRange.Cells(1, 1).Value
    If Err.Number <> 0 Then NamedVal = ""
    On Error GoTo 0
    If IsError(NamedVal) Then NamedVal = ""
End Function
