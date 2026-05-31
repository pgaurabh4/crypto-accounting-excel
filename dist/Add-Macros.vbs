' Add-Macros.vbs — build GravityLedger.xlsm from GravityLedger.xlsx
' Double-click on Windows with Excel installed.
' Prerequisite (one-time): Excel > File > Options > Trust Center > Trust Center
' Settings > Macro Settings > tick "Trust access to the VBA project object model".
Option Explicit
Dim fso, here, xlsx, xlsm, vbaDir, xl, wb, proj, f, ws, btn, file, lastRow
Set fso = CreateObject("Scripting.FileSystemObject")
here   = fso.GetParentFolderName(WScript.ScriptFullName)
xlsx   = fso.BuildPath(here, "GravityLedger.xlsx")
xlsm   = fso.BuildPath(here, "GravityLedger.xlsm")
vbaDir = fso.BuildPath(here, "vba")

If Not fso.FileExists(xlsx) Then
    MsgBox "GravityLedger.xlsx not found next to this script.", vbCritical : WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(xlsx)

' Save as macro-enabled (xlOpenXMLWorkbookMacroEnabled = 52) BEFORE injecting VBA.
wb.SaveAs xlsm, 52

On Error Resume Next
Set proj = wb.VBProject
On Error GoTo 0
If proj Is Nothing Then
    MsgBox "Cannot access the VBA project." & vbCrLf & _
        "Enable: Trust Center > Macro Settings > 'Trust access to the VBA project object model', then re-run.", _
        vbCritical
    wb.Close False : xl.Quit : WScript.Quit 1
End If

' Import every .bas module.
Set f = fso.GetFolder(vbaDir)
For Each file In f.Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        proj.VBComponents.Import file.Path
    End If
Next

' Place the RUN ALL button on the README sheet.
Set ws = wb.Sheets("README")
lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   ' xlUp
Set btn = ws.Buttons.Add(ws.Cells(lastRow, 2).Left, ws.Cells(lastRow, 2).Top, 120, 28)
btn.OnAction = "RunAll"
btn.Caption = "RUN ALL"
btn.Name = "btnRunAll"

' Build the ledger once so reports are populated on first open.
On Error Resume Next
xl.Run "RunAllSilent"
On Error GoTo 0

wb.Save
wb.Close True
xl.Quit
MsgBox "Built " & xlsm & vbCrLf & "Open it, enable macros, and click RUN ALL anytime to rebuild.", vbInformation, "Gravity Ledger"
