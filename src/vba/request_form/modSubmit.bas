Attribute VB_Name = "modSubmit"
Option Explicit

' =====================================================================
'  TPD Travel/Training Request Form - the officer's one button.
'  Validates the form, then opens a pre-addressed Outlook email to the
'  Training Unit with this workbook attached. The officer reviews and
'  presses Send - nothing sends automatically.
'  Keyboard fallback: Ctrl+Shift+S (wired by Auto_Open).
' =====================================================================

Public Sub Auto_Open()
    Application.OnKey "^+S", "SubmitToTrainingUnit"
End Sub

Public Sub SubmitToTrainingUnit()
    Dim missing As String, ol As Object, mail As Object
    Dim attach As String, officer As String, course As String

    missing = MissingFields()
    If Len(missing) > 0 Then
        MsgBox "Please fill in:" & vbCrLf & missing, vbExclamation, _
               "A few boxes are still empty"
        Exit Sub
    End If

    officer = CStr(NamedVal("rq_Name"))
    course = CStr(NamedVal("rq_Course"))

    ' Attach a saved copy (works even if the officer never saved the file).
    attach = Environ$("TEMP") & "\Travel Request - " & CleanFileName(officer) & _
             " - " & Format$(Now, "yyyymmdd-hhnn") & _
             Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    ThisWorkbook.SaveCopyAs attach

    On Error GoTo Fail
    Set ol = GetOutlook()
    Set mail = ol.CreateItem(0)          ' olMailItem
    mail.To = CStr(NamedVal("rq_TrainingEmail"))
    mail.Subject = CStr(NamedVal("rq_SubjectPrefix")) & " - " & officer & _
                   " - " & course
    mail.Body = "Training Unit," & vbCrLf & vbCrLf & _
                "Attached is my travel/training request:" & vbCrLf & _
                "    Course:  " & course & vbCrLf & _
                "    Dates:   " & Format$(NamedVal("rq_Start"), "mm/dd/yyyy") & _
                " - " & Format$(NamedVal("rq_End"), "mm/dd/yyyy") & vbCrLf & vbCrLf & _
                "Thank you," & vbCrLf & officer
    mail.Attachments.Add attach
    mail.Display                          ' review, then press Send
    Exit Sub

Fail:
    MsgBox "Could not open an Outlook draft (" & Err.Description & ")." & vbCrLf & _
           "If Outlook is not available, email the saved copy yourself:" & _
           vbCrLf & attach, vbExclamation, "Submit"
End Sub

Private Function MissingFields() As String
    ' Field name -> friendly label. Costs may legitimately be zero/blank.
    Dim req As Variant, i As Long, s As String
    req = Array( _
        "rq_Name", "your name", _
        "rq_Rank", "your rank", _
        "rq_EmpID", "your employee ID", _
        "rq_Email", "your email", _
        "rq_SupName", "your supervisor's name", _
        "rq_SupEmail", "your supervisor's email", _
        "rq_Course", "the course/event title", _
        "rq_Location", "the location", _
        "rq_Start", "the class start date", _
        "rq_End", "the class end date", _
        "rq_DepDate", "your departure date", _
        "rq_DepTime", "your departure time", _
        "rq_RetDate", "your return date", _
        "rq_RetTime", "your return time")
    For i = LBound(req) To UBound(req) Step 2
        If Len(Trim$(CStr(NamedVal(CStr(req(i)))))) = 0 Then
            s = s & "  - " & req(i + 1) & vbCrLf
        End If
    Next i
    MissingFields = s
End Function

Private Function NamedVal(ByVal n As String) As Variant
    On Error Resume Next
    NamedVal = ThisWorkbook.Names(n).RefersToRange.Cells(1, 1).Value
    If Err.Number <> 0 Then NamedVal = ""
    On Error GoTo 0
    If IsError(NamedVal) Then NamedVal = ""
End Function

Private Function GetOutlook() As Object
    On Error Resume Next
    Set GetOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If GetOutlook Is Nothing Then Set GetOutlook = CreateObject("Outlook.Application")
End Function

Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, b As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each b In bad
        s = Replace(s, CStr(b), "-")
    Next b
    CleanFileName = Trim$(s)
End Function
