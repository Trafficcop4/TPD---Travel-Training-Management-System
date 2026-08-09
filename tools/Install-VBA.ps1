<#
.SYNOPSIS
    One-time build step, run on a Windows PC with Excel and this repo:
    turns the two .xlsx workbooks into the finished .xlsm files by
    importing the VBA source (src/vba/**) and adding the macro buttons.

.DESCRIPTION
    Produces:
      workbooks\TPD_Training_Master.xlsm      (coordinator's workbook)
      workbooks\TPD_Travel_Request_Form.xlsm  (officer form)

    Prerequisite (one time, then it can be turned back off):
      Excel > File > Options > Trust Center > Trust Center Settings >
      Macro Settings > check "Trust access to the VBA project object model".

    Safe to re-run: it starts from the .xlsx sources each time and
    overwrites the .xlsm outputs.

.NOTES
    No parameters needed when run from the repo's tools folder:
        powershell -ExecutionPolicy Bypass -File .\Install-VBA.ps1
#>
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$workbooks = Join-Path $RepoRoot 'workbooks'
$vbaMaster = Join-Path $RepoRoot 'src\vba\master'
$vbaForm   = Join-Path $RepoRoot 'src\vba\request_form'
$xlsmFmt   = 52   # xlOpenXMLWorkbookMacroEnabled
$btnCtrl   = 0    # xlButtonControl

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
    $shape = $ws.Shapes.AddFormControl($btnCtrl, $left, $top, $width - 4, 22)
    $shape.Name = "btn_$macro"
    $shape.OnAction = $macro
    $shape.TextFrame.Characters().Text = $caption
}

function Remove-OldButtons($ws) {
    # re-run safety: drop any buttons a previous install added
    @($ws.Shapes) | Where-Object { $_.Name -like 'btn_*' } |
        ForEach-Object { $_.Delete() }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # --- verify VBA project model access before doing anything ---------
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

    # --- master workbook ----------------------------------------------
    Write-Host 'Building TPD_Training_Master.xlsm'
    $wb = $excel.Workbooks.Open((Join-Path $workbooks 'TPD_Training_Master.xlsx'))
    Import-Modules $wb $vbaMaster
    $dash = $wb.Worksheets.Item('Dashboard')
    Remove-OldButtons $dash
    # a strip of buttons under the BUTTONS banner (row 34, cols B..O)
    Add-Button $dash 'B34' 2 'Import Requests'      'ImportRequests'
    Add-Button $dash 'D34' 2 'Approval Email'       'DraftApprovalEmail'
    Add-Button $dash 'G34' 2 'Officer Notification' 'DraftOfficerNotification'
    Add-Button $dash 'I34' 2 'Settlement Email'     'DraftSettlementEmail'
    Add-Button $dash 'L34' 1 'Print Packet'         'PrintPacket'
    Add-Button $dash 'M34' 2 'Push to Outlook'      'PushDeadlinesToOutlook'
    Add-Button $dash 'O34' 1 'Archive Year'         'ArchiveYear'
    $out = Join-Path $workbooks 'TPD_Training_Master.xlsm'
    $wb.SaveAs($out, $xlsmFmt)
    $wb.Close($false)
    Write-Host "    saved   $out"

    # --- officer request form -----------------------------------------
    Write-Host 'Building TPD_Travel_Request_Form.xlsm'
    $wb = $excel.Workbooks.Open((Join-Path $workbooks 'TPD_Travel_Request_Form.xlsx'))
    Import-Modules $wb $vbaForm
    $form = $wb.Worksheets.Item('Request Form')
    Remove-OldButtons $form
    Add-Button $form 'C30' 2 'Submit to Training Unit' 'SubmitToTrainingUnit'
    $out = Join-Path $workbooks 'TPD_Travel_Request_Form.xlsm'
    $wb.SaveAs($out, $xlsmFmt)
    $wb.Close($false)
    Write-Host "    saved   $out"

    Write-Host ''
    Write-Host 'Done. Distribute the .xlsm files; keep the .xlsx sources for rebuilds.'
    Write-Host 'Remember to set the yellow cells on the master Settings sheet and the'
    Write-Host 'Training Unit address on the form''s hidden "Form Settings" sheet.'
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
