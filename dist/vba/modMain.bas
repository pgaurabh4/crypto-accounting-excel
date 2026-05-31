Attribute VB_Name = "modMain"
'====================================================================
' Gravity Ledger - master crypto accounting workbook
' Mirrors the Rust double-entry engine (pgaurabh4/crypto-accounting).
' INPUTS you edit:  OpeningBalances, Trades, Transfers, BankStatement
' Everything else is DERIVED on RunAll: facts by macro, reports by SQL.
'====================================================================
Option Explicit

Public Const ASSET As String = "Asset"
Public Const LIAB As String = "Liability"
Public Const EQUITY As String = "Equity"
Public Const INCOME As String = "Income"
Public Const EXPENSE As String = "Expense"

Public gSilent As Boolean       ' True suppresses dialogs (used by the builder)

' RUN ALL button target. Rebuilds the entire ledger from the 4 input sheets.
Public Sub RunAll()
    Dim t As Single: t = Timer
    On Error GoTo fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Engine_Reset
    Engine_PostOpening
    Engine_PostHistory      ' trades + transfers + bank, merged by timestamp
    Engine_RevalueAll       ' period-end mark-to-market (last price or Marks override)
    Engine_DumpFacts        ' write ChartOfAccounts/Journal/JournalLines/Lots

    Reports_BuildAll        ' pure-VBA aggregation -> every report sheet
    Recon_Build             ' match BankStatement lines to posted cash journals
    ThisWorkbook.Save       ' persist the rebuilt workbook

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Not gSilent Then _
        MsgBox "Ledger rebuilt in " & Format(Timer - t, "0.0") & "s." & vbCrLf & _
           "Global imbalance = " & Engine_Imbalance() & " (must be 0)." & vbCrLf & _
           "Reporting engine: " & gReportEngine, vbInformation, "Gravity Ledger"
    Exit Sub
fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Not gSilent Then MsgBox "RunAll failed: " & Err.Description, vbCritical, "Gravity Ledger"
End Sub

' Entry point the .vbs builder calls so no dialog blocks the headless build.
Public Sub RunAllSilent()
    gSilent = True
    RunAll
    gSilent = False
End Sub
