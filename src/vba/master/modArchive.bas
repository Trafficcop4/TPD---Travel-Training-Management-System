Attribute VB_Name = "modArchive"
Option Explicit

' =====================================================================
'  ArchiveYear - step 2 of the 10-minute year-end rollover
'  (the full procedure is on the Instructions sheet).
'  Saves a complete, self-contained copy of this workbook to the
'  ArchiveFolder named for the closing fiscal year. That copy is the
'  permanent record; this workbook then rolls forward.
' =====================================================================

Public Sub ArchiveYear()
    Dim folder As String, fy As String, ext As String, path As String

    folder = FolderSetting("ArchiveFolder")
    fy = CStr(Setting("FY_Label"))
    If Dir(folder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
        If Dir(folder, vbDirectory) = "" Then
            MsgBox "Archive folder not found and could not be created:" & _
                   vbCrLf & folder & vbCrLf & "Fix the path on Settings.", _
                   vbExclamation, "Archive Year"
            Exit Sub
        End If
    End If

    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    path = folder & "\TPD_Training_Master_" & fy & ext
    If Dir(path) <> "" Then
        If MsgBox("An archive for " & fy & " already exists:" & vbCrLf & path & _
                  vbCrLf & vbCrLf & "Overwrite it?", vbYesNo + vbExclamation, _
                  "Archive Year") <> vbYes Then Exit Sub
    End If

    ThisWorkbook.Save
    ThisWorkbook.SaveCopyAs path

    MsgBox "Archived: " & path & vbCrLf & vbCrLf & _
           "Now finish the rollover (Instructions sheet, steps 3-5):" & vbCrLf & _
           "  3. Settings - new FY label, FY start, tracking prefix, budget." & vbCrLf & _
           "  4. Requests - clear the typed values in all data rows" & vbCrLf & _
           "     (formula and helper columns stay)." & vbCrLf & _
           "  5. Check the Dashboard reads zero, then Save.", _
           vbInformation, "Archive Year"
End Sub
