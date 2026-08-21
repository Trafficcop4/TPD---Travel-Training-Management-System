<#
.SYNOPSIS
    One-time build step, run on a Windows PC with Excel and this repo:
    turns workbooks\BPOC_Academy_Management_V6.xlsx into the finished
    .xlsm by importing the VBA source (src\vba\bpoc\*.bas) and adding
    the macro buttons.

.DESCRIPTION
    Produces:  workbooks\BPOC_Academy_Management_V6.xlsm

    Prerequisite (one time, then it can be turned back off):
      Excel > File > Options > Trust Center > Trust Center Settings >
      Macro Settings > check "Trust access to the VBA project object model".

    Safe to re-run: starts from the .xlsx source each time and overwrites
    the .xlsm output.

.NOTES
    Run from the repo's tools folder:
        powershell -ExecutionPolicy Bypass -File .\Install-BPOC-VBA.ps1
#>
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$workbooks = Join-Path $RepoRoot 'workbooks'
$vbaBpoc   = Join-Path $RepoRoot 'src\vba\bpoc'
$xlsmFmt   = 52   # xlOpenXMLWorkbookMacroEnabled
$btnCtrl   = 0    # xlButtonControl
$pw        = 'TPDAcademy'

function Import-Modules($wb, $folder) {
    Get-ChildItem -Path $folder -Filter '*.bas' | Sort-Object Name | ForEach-Object {
        [void]$wb.VBProject.VBComponents.Import($_.FullName)
        Write-Host "    module  $($_.Name)"
    }
}

function Add-Button($ws, $anchorCell, $widthCells, $caption, $macro) {
    $anchor = $ws.Range($anchorCell)
    $left   = $anchor.Left
    $top    = $anchor.Top + 2
    $width  = 0
    for ($i = 0; $i -lt $widthCells; $i++) { $width += $anchor.Offset(0, $i).Width }
    # never narrower than the caption needs: anchors on default-width columns
    # produced ~44pt buttons whose text was clipped to a few characters
    # NOTE: this minimum is why anchors must be spaced at least 130pt apart.
    # Two buttons anchored closer than $w WILL overlap - see the StartHere
    # anchors below. Warn rather than silently drawing one over the other.
    $w = [Math]::Max($width - 4, 130)
    if ($width -gt 0 -and $width -lt 130) {
        Write-Host ("    note    '{0}' widened to 130pt (anchor span {1:N0}pt) - " +
                    'keep the next anchor at least 130pt to the right') `
                   -f $caption, $width
    }
    $shape = $ws.Shapes.AddFormControl($btnCtrl, $left, $top, $w, 22)
    $shape.Name = "btn_$macro"
    $shape.OnAction = $macro
    $shape.TextFrame.Characters().Text = $caption
}

function Remove-OldButtons($ws) {
    @($ws.Shapes) | Where-Object { $_.Name -like 'btn_*' } |
        ForEach-Object { $_.Delete() }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $probe = $excel.Workbooks.Add()
    try {
        [void]$probe.VBProject.Name
    } catch {
        Write-Host ''
        Write-Host 'BLOCKED: Excel is not allowing scripts to write VBA.' -ForegroundColor Red
        Write-Host 'Enable it once, then re-run this script:'
        Write-Host '  Excel > File > Options > Trust Center > Trust Center Settings'
        Write-Host '  > Macro Settings > [x] Trust access to the VBA project object model'
        exit 1
    } finally {
        $probe.Close($false)
    }

    Write-Host 'Building BPOC_Academy_Management_V6.xlsm'
    $wb = $excel.Workbooks.Open((Join-Path $workbooks 'BPOC_Academy_Management_V6.xlsx'))
    Import-Modules $wb $vbaBpoc

    # workbook event: auto-capitalize x -> X on the Writing grid
    $eventCode = @'
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' any runtime error must never leave EnableEvents off
    On Error GoTo Cleanup
    ' Writing grid: auto-capitalize x -> X
    If Sh.Name = "Writing" Then
        Dim rng As Range
        Set rng = Application.Intersect(Target, Sh.Range("D6:AQ55"))
        If Not rng Is Nothing Then
            Dim c As Range
            Application.EnableEvents = False
            For Each c In rng
                If LCase$(Trim$(CStr(c.Value))) = "x" Then c.Value = "X"
            Next c
            Application.EnableEvents = True
        End If
        Exit Sub
    End If
    ' Schedule instructors: dropdown multi-select (pick again to remove)
    If Sh.Name = "Schedule" Then
        If Target.Cells.Count <> 1 Then Exit Sub
        If Application.Intersect(Target, Sh.Range("I6:I905")) Is Nothing Then Exit Sub
        Dim newVal As String, oldVal As String
        newVal = Trim$(CStr(Target.Value))
        If newVal = "" Or InStr(newVal, ",") > 0 Then Exit Sub
        Application.EnableEvents = False
        Application.Undo
        oldVal = Trim$(CStr(Target.Value))
        ' Rebuild the list token by token. The old code used
        ' Replace(", " & oldVal, ", " & newVal, "") with NO occurrence count:
        '   * it removed EVERY match, not the one that was re-picked, and
        '   * because it matched on a prefix, re-picking "Smith" out of
        '     "Smith, Smithson" deleted both and left the fragment "son".
        ' It also could not remove the LAST remaining name - "oldVal = newVal"
        ' rewrote the cell instead of clearing it, contradicting the
        ' "pick again to remove" contract this same block advertises.
        ' Instructor names never contain a comma (see DL.INSTRUCTORS /
        ' GUEST_ENTITIES), so splitting on "," is safe.
        Dim parts() As String, i As Long, rebuilt As String, removed As Boolean
        Dim tok As String
        parts = Split(oldVal, ",")
        For i = LBound(parts) To UBound(parts)
            tok = Trim$(parts(i))
            If tok <> "" Then
                If Not removed And StrComp(tok, newVal, vbTextCompare) = 0 Then
                    removed = True          ' drop exactly one occurrence
                ElseIf rebuilt = "" Then
                    rebuilt = tok
                Else
                    rebuilt = rebuilt & ", " & tok
                End If
            End If
        Next i
        If oldVal = "" Then
            Target.Value = newVal
        ElseIf removed Then
            If rebuilt = "" Then
                Target.ClearContents        ' removing the last name clears it
            Else
                Target.Value = rebuilt
            End If
        Else
            Target.Value = oldVal & ", " & newVal
        End If
        Application.EnableEvents = True
    End If
    Exit Sub
Cleanup:
    Application.EnableEvents = True
End Sub
'@
    $twb = $wb.VBProject.VBComponents.Item('ThisWorkbook')
    $twb.CodeModule.AddFromString($eventCode)
    Write-Host '    event   Workbook_SheetChange (Writing x -> X)'

    # buttons: Print Center strip + email + reset
    $pc = $wb.Worksheets.Item('PrintCenter')
    $pc.Unprotect($pw) 2>$null
    Remove-OldButtons $pc
    # ONE button per table row. The printables table is rows 8..22 (one row
    # each); the old every-other-row grid starting at row 6 put 14 of the 15
    # buttons beside a row naming a different macro, and the last four at or
    # below the end of the sheet. Keep this list in the same order as `rows`
    # in build_printcenter (build/bpoc/sheets_outputs.py).
    Add-Button $pc 'G8'  2 'Sign-In Sheet'     'modPrint.btnPrintSignIn'
    Add-Button $pc 'G9'  2 'Sign-In Week'      'modPrint.btnPrintSignInWeek'
    Add-Button $pc 'G10' 2 'Academy Book'      'modPrint.btnPrintSignInAcademy'
    Add-Button $pc 'G11' 2 'Eval Stack'        'modPrint.btnPrintEvals'
    Add-Button $pc 'G12' 2 'Spelling Test/Key' 'modPrint.btnPrintSpelling'
    Add-Button $pc 'G13' 2 'Writing Handout'   'modPrint.btnPrintWriting'
    Add-Button $pc 'G14' 2 'Cadet Profile'     'modPrint.btnPrintProfile'
    Add-Button $pc 'G15' 2 'Transcript(s)'     'modPrint.btnPrintTranscript'
    Add-Button $pc 'G16' 2 'Ranking'           'modPrint.btnPrintRanking'
    Add-Button $pc 'G17' 2 'Grad Checklist'    'modPrint.btnPrintGradCheck'
    Add-Button $pc 'G18' 2 'Audit Packet'      'modPrint.btnPrintAudit'
    Add-Button $pc 'G19' 2 'Chapter Packet'    'modPrint.btnPrintChapterPacket'
    Add-Button $pc 'G20' 2 'Exam Grade Sheet'  'modPrint.btnPrintGradeSheet'
    Add-Button $pc 'G21' 2 'Addendum Report'   'modPrint.btnPrintAddendum'
    Add-Button $pc 'G22' 2 'Schedule'          'modPrint.btnPrintSchedule'

    # Dashboard is now a protected pure-output sheet (zero input cells), so
    # it has to be unprotected while its button is (re)placed, exactly like
    # PrintCenter and EmailPreview.
    $dash = $wb.Worksheets.Item('Dashboard')
    $dash.Unprotect($pw) 2>$null
    Remove-OldButtons $dash
    Add-Button $dash 'K5' 1 'Agency Email Drafts' 'modAgencyEmail.GenerateAgencyEmails'
    $dash.Protect($pw) 2>$null

    $ep = $wb.Worksheets.Item('EmailPreview')
    $ep.Unprotect($pw) 2>$null
    Remove-OldButtons $ep
    # L5, not H5: the E5 status line ("Exam # ... | Last emailed: <date>")
    # spills across F:J, and a button anchored at H5 covered the cutoff date
    # this whole preview sheet exists to show.
    Add-Button $ep 'L5' 2 'Build Outlook Drafts' 'modAgencyEmail.GenerateAgencyEmails'
    $ep.Protect($pw) 2>$null

    $sh = $wb.Worksheets.Item('StartHere')
    Remove-OldButtons $sh
    # D16 / H16, not D16 / F16. Add-Button forces a 130pt minimum width (the
    # two-default-width span D:E is only ~96pt), so anchoring the second
    # button two columns right drew 'Startup Review' over the right ~34pt of
    # 'New Academy Reset'. Four columns (~192pt) clears it with room to spare.
    Add-Button $sh 'D16' 2 'New Academy Reset' 'modNewAcademy.NewAcademyReset'
    Add-Button $sh 'H16' 2 'Startup Review'    'modNewAcademy.AcademyStartupReview'

    $out = Join-Path $workbooks 'BPOC_Academy_Management_V6.xlsm'
    $wb.SaveAs($out, $xlsmFmt)
    $wb.Close($false)
    Write-Host "    saved   $out"

    Write-Host ''
    Write-Host 'Done. Distribute the .xlsm; keep the .xlsx source for rebuilds.'
    Write-Host 'First open: enable macros, then check Settings and the PT rubric block.'
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
