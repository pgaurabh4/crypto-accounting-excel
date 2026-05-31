//! The VBA engine source, mirroring the Rust ledger in `app/src/*`.
//!
//! Returned as (module_name, source) pairs. Two layers:
//!   * modEngine  — procedural ingestion + double-entry posting + FIFO lots,
//!                  writing the "fact" sheets (ChartOfAccounts/Journal/
//!                  JournalLines/Lots). Sequential by nature.
//!   * modReports — SQL (ADODB + ACE OLEDB) GROUP-BY over those fact sheets to
//!                  build every report sheet; a VBA fallback runs the same
//!                  aggregations from memory if the SQL provider is absent.
//!   * modMain    — RunAll orchestration + the RUN ALL button handler.

pub fn modules() -> Vec<(&'static str, &'static str)> {
    vec![
        ("modMain", MOD_MAIN),
        ("modEngine", MOD_ENGINE),
        ("modReports", MOD_REPORTS),
    ]
}

const MOD_MAIN: &str = r#"Attribute VB_Name = "modMain"
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
"#;

const MOD_ENGINE: &str = r#"Attribute VB_Name = "modEngine"
Option Explicit
'--------------------------------------------------------------------
' In-memory ledger state (rebuilt each RunAll). Money is VBA Decimal
' (CDec) so 1e10 IDR notionals and fractional-crypto bases stay exact.
'--------------------------------------------------------------------
Private accById As Object      ' id -> Array(entity,code,name,type,ccy)
Private accByKey As Object      ' "entity|code" -> id
Private jrnTs As Object         ' journal_id -> ts/kind/memo/entity (parallel arrays via collections)
Private jId As Long
Private aId As Long
Private lotId As Long
Private jr As Collection        ' each: Array(id, ts, kind, memo, entity)
Private jl As Collection        ' each: Array(journal_id, account_id, amount)
Private lots As Collection      ' each: dict-like Array(id,entity,asset,exchange,acct,acq,qty,unit,srcJ)
Private fvadj As Object         ' "entity|asset" -> current adj (Decimal)

Public Sub Engine_Reset()
    Set accById = CreateObject("Scripting.Dictionary")
    Set accByKey = CreateObject("Scripting.Dictionary")
    Set jr = New Collection
    Set jl = New Collection
    Set lots = New Collection
    Set fvadj = CreateObject("Scripting.Dictionary")
    jId = 0: aId = 0: lotId = 0
End Sub

Public Function D(v As Variant) As Variant
    ' Coerce to Decimal; blanks -> 0. Check string-ness BEFORE comparing to ""
    ' so a numeric cell is never mistaken for empty.
    If IsNull(v) Or IsEmpty(v) Then D = CDec(0): Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then D = CDec(0): Exit Function
    End If
    D = CDec(v)
End Function

'----- accounts ------------------------------------------------------
Private Function AcctType(ByVal code As String) As String
    If Left$(code, 5) = "CASH:" Then
        AcctType = ASSET
    ElseIf Left$(code, 7) = "CRYPTO:" Then
        AcctType = ASSET
    ElseIf Left$(code, 5) = "LIAB:" Then
        AcctType = LIAB
    ElseIf Left$(code, 3) = "EQ:" Then
        AcctType = EQUITY
    ElseIf Left$(code, 4) = "PNL:" Then
        AcctType = INCOME
    ElseIf Left$(code, 4) = "INC:" Then
        AcctType = INCOME
    ElseIf Left$(code, 4) = "EXP:" Then
        AcctType = EXPENSE
    Else
        AcctType = EQUITY
    End If
End Function

Public Function IsDebitNormal(ByVal t As String) As Boolean
    IsDebitNormal = (t = ASSET Or t = EXPENSE)
End Function

Private Function GetAcct(entity As String, code As String, Optional ccy As String = "USD", Optional nm As String = "") As Long
    Dim k As String: k = entity & "|" & code
    If accByKey.Exists(k) Then GetAcct = accByKey(k): Exit Function
    aId = aId + 1
    If nm = "" Then nm = code
    accById(aId) = Array(entity, code, nm, AcctType(code), ccy)
    accByKey(k) = aId
    GetAcct = aId
End Function

'----- posting -------------------------------------------------------
' lines() = 2D-ish: pass arrays of account_id and Decimal amounts.
Private Function Post(ts As String, kind As String, memo As String, entity As String, ids() As Long, amts() As Variant, ByVal nlines As Long) As Long
    Dim s As Variant: s = CDec(0)
    Dim i As Long
    For i = 0 To nlines - 1: s = s + amts(i): Next i
    If s <> CDec(0) Then Err.Raise vbObjectError + 1, , "unbalanced journal (" & kind & "/" & memo & "): sum=" & CStr(s)
    jId = jId + 1
    jr.Add Array(jId, ts, kind, memo, entity)
    For i = 0 To nlines - 1
        jl.Add Array(jId, ids(i), amts(i))
    Next i
    Post = jId
End Function

Private Sub AddLot(entity As String, asset As String, exch As String, acct As Long, acq As String, qty As Variant, unit As Variant, srcJ As Long)
    lotId = lotId + 1
    lots.Add Array(lotId, entity, asset, exch, acct, acq, qty, unit, srcJ)
End Sub

' FIFO consume `qty` from oldest lots of (entity, acct). Returns cost basis.
Private Function ConsumeFifo(entity As String, acct As Long, qty As Variant) As Variant
    Dim q As Variant: q = D(qty)
    If q < CDec(0) Then Err.Raise vbObjectError + 2, , "consume negative qty"
    If q = CDec(0) Then ConsumeFifo = CDec(0): Exit Function
    ' availability
    Dim avail As Variant: avail = CDec(0)
    Dim i As Long, L As Variant
    For i = 1 To lots.Count
        L = lots(i)
        If L(1) = entity And L(4) = acct Then avail = avail + L(6)
    Next i
    If avail < q Then Err.Raise vbObjectError + 3, , "insufficient inventory acct " & acct & " (need " & CStr(q) & ", have " & CStr(avail) & ") - spot only"

    ' oldest first: build sorted index by (acq, id)
    Dim idx() As Long, cnt As Long: cnt = 0
    ReDim idx(1 To lots.Count)
    For i = 1 To lots.Count
        L = lots(i)
        If L(1) = entity And L(4) = acct And L(6) > CDec(0) Then cnt = cnt + 1: idx(cnt) = i
    Next i
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To cnt - 1
        For b = a + 1 To cnt
            If lots(idx(b))(5) < lots(idx(a))(5) Or (lots(idx(b))(5) = lots(idx(a))(5) And lots(idx(b))(0) < lots(idx(a))(0)) Then
                tmp = idx(a): idx(a) = idx(b): idx(b) = tmp
            End If
        Next b
    Next a

    Dim remaining As Variant: remaining = q
    Dim cost As Variant: cost = CDec(0)
    Dim j As Long
    For j = 1 To cnt
        If remaining = CDec(0) Then Exit For
        L = lots(idx(j))
        Dim lq As Variant: lq = L(6)
        Dim uc As Variant: uc = L(7)
        If lq <= remaining Then
            cost = cost + lq * uc
            remaining = remaining - lq
            L(6) = CDec(0)                ' empty
            lots.Remove idx(j): lots.Add L  ' move to tail w/ qty 0 (kept out by qty filter)
            ' note: removing shifts indices; re-find below avoided by single pass using qty=0 marker
            ReindexAfterRemoval idx, idx(j), cnt
        Else
            cost = cost + remaining * uc
            L(6) = lq - remaining
            remaining = CDec(0)
            lots.Remove idx(j): lots.Add L
            ReindexAfterRemoval idx, idx(j), cnt
        End If
    Next j
    If remaining <> CDec(0) Then Err.Raise vbObjectError + 4, , "fifo underflow"
    ConsumeFifo = cost
End Function

' Collection.Remove shifts every later index down by one; fix our idx map.
Private Sub ReindexAfterRemoval(ByRef idx() As Long, ByVal removed As Long, ByVal cnt As Long)
    Dim k As Long
    For k = 1 To cnt
        If idx(k) > removed Then idx(k) = idx(k) - 1
    Next k
End Sub

'----- event handlers (mirror trades.rs / ops.rs) --------------------
Private Sub DoTrade(entity As String, side As String, base As String, quote As String, qty As Variant, price As Variant, fee As Variant, ts As String)
    Dim cryptoA As Long, cashA As Long, feeA As Long, realA As Long
    cryptoA = GetAcct(entity, "CRYPTO:" & base & ":MAIN", base, base & " inventory @MAIN")
    cashA = GetAcct(entity, "CASH:" & quote, quote, "Cash " & quote)
    feeA = GetAcct(entity, "EXP:FEES:TRADING", quote, "Trading fees")
    Dim notional As Variant: notional = D(qty) * D(price)
    Dim ids(0 To 3) As Long, amts(0 To 3) As Variant
    If LCase$(side) = "buy" Then
        ids(0) = cryptoA: amts(0) = notional
        ids(1) = feeA: amts(1) = D(fee)
        ids(2) = cashA: amts(2) = -(notional + D(fee))
        Post ts, "trade", "buy " & CStr(qty) & " " & base & "/" & quote, entity, ids, amts, 3
        AddLot entity, base, "MAIN", cryptoA, ts, D(qty), D(price), jId
    Else
        realA = GetAcct(entity, "PNL:REALIZED", "USD", "Realized PnL")
        Dim cb As Variant: cb = ConsumeFifo(entity, cryptoA, D(qty))
        Dim proceeds As Variant: proceeds = notional
        Dim realized As Variant: realized = proceeds - cb
        ids(0) = cashA: amts(0) = proceeds - D(fee)
        ids(1) = feeA: amts(1) = D(fee)
        ids(2) = cryptoA: amts(2) = -cb
        ids(3) = realA: amts(3) = -realized
        Post ts, "trade", "sell " & CStr(qty) & " " & base & "/" & quote, entity, ids, amts, 4
    End If
End Sub

Private Sub DoTransfer(entity As String, asset As String, fromX As String, toX As String, qty As Variant, fee As Variant, ts As String)
    Dim fromA As Long, toA As Long, feeNet As Long
    fromA = GetAcct(entity, "CRYPTO:" & asset & ":" & fromX, asset)
    toA = GetAcct(entity, "CRYPTO:" & asset & ":" & toX, asset)
    feeNet = GetAcct(entity, "EXP:FEES:NETWORK", "USD", "Network fees")
    Dim costMoved As Variant: costMoved = ConsumeFifo(entity, fromA, D(qty))
    Dim costFee As Variant: costFee = CDec(0)
    If D(fee) > CDec(0) Then costFee = ConsumeFifo(entity, fromA, D(fee))
    Dim ids(0 To 2) As Long, amts(0 To 2) As Variant
    ids(0) = toA: amts(0) = costMoved
    ids(1) = fromA: amts(1) = -(costMoved + costFee)
    ids(2) = feeNet: amts(2) = costFee
    Post ts, "transfer", asset & " " & fromX & "->" & toX, entity, ids, amts, 3
    If D(qty) > CDec(0) Then AddLot entity, asset, toX, toA, ts, D(qty), costMoved / D(qty), jId
End Sub

Private Sub DoBank(entity As String, ts As String, ccy As String, amount As Variant, counter As String, memo As String)
    Dim cashA As Long, ctrA As Long
    cashA = GetAcct(entity, "CASH:" & ccy, ccy, "Cash " & ccy)
    If counter = "" Then counter = "EQ:CAPITAL"
    ctrA = GetAcct(entity, counter, ccy, counter)
    Dim ids(0 To 1) As Long, amts(0 To 1) As Variant
    ids(0) = cashA: amts(0) = D(amount)            ' + deposit (debit) / - withdrawal (credit)
    ids(1) = ctrA: amts(1) = -D(amount)
    Post ts, "bank", memo, entity, ids, amts, 2
End Sub

Private Sub DoOpenCash(entity As String, ccy As String, amount As Variant, asof As String)
    Dim cashA As Long, opEq As Long
    cashA = GetAcct(entity, "CASH:" & ccy, ccy, "Cash " & ccy)
    opEq = GetAcct(entity, "EQ:OPENING", ccy, "Opening balance equity")
    Dim ids(0 To 1) As Long, amts(0 To 1) As Variant
    ids(0) = cashA: amts(0) = D(amount)
    ids(1) = opEq: amts(1) = -D(amount)
    Post asof, "opening", "opening cash " & ccy, entity, ids, amts, 2
End Sub

Private Sub DoOpenCrypto(entity As String, asset As String, exch As String, qty As Variant, unit As Variant, asof As String)
    Dim cA As Long, opEq As Long
    cA = GetAcct(entity, "CRYPTO:" & asset & ":" & exch, asset)
    opEq = GetAcct(entity, "EQ:OPENING", "USD", "Opening balance equity")
    Dim basis As Variant: basis = D(qty) * D(unit)
    Dim ids(0 To 1) As Long, amts(0 To 1) As Variant
    ids(0) = cA: amts(0) = basis
    ids(1) = opEq: amts(1) = -basis
    Post asof, "opening", "opening " & asset & " @" & exch, entity, ids, amts, 2
    AddLot entity, asset, exch, cA, asof, D(qty), D(unit), jId
End Sub

Private Sub DoOpenLoan(entity As String, ccy As String, amount As Variant, counterparty As String, asof As String)
    Dim loanA As Long, opEq As Long
    loanA = GetAcct(entity, "LIAB:LOAN", ccy, "Loan payable")
    opEq = GetAcct(entity, "EQ:OPENING", ccy, "Opening balance equity")
    Dim ids(0 To 1) As Long, amts(0 To 1) As Variant
    ids(0) = loanA: amts(0) = -D(amount)           ' credit liability
    ids(1) = opEq: amts(1) = D(amount)             ' debit equity (reduces it)
    Post asof, "opening", "opening loan " & counterparty, entity, ids, amts, 2
End Sub

Private Sub DoRevalue(entity As String, asset As String, mark As Variant, ts As String)
    Dim qty As Variant: qty = CDec(0)
    Dim cost As Variant: cost = CDec(0)
    Dim i As Long, L As Variant
    For i = 1 To lots.Count
        L = lots(i)
        If L(1) = entity And L(2) = asset Then
            qty = qty + L(6): cost = cost + L(6) * L(7)
        End If
    Next i
    If qty = CDec(0) Then Exit Sub
    Dim target As Variant: target = qty * D(mark) - cost
    Dim k As String: k = entity & "|" & asset
    Dim cur As Variant: cur = CDec(0)
    If fvadj.Exists(k) Then cur = fvadj(k)
    Dim delta As Variant: delta = target - cur
    Dim fvA As Long, unA As Long
    fvA = GetAcct(entity, "CRYPTO:FVADJ", "USD", "Fair-value adjustment")
    unA = GetAcct(entity, "PNL:UNREALIZED", "USD", "Unrealized PnL")
    Dim ids(0 To 1) As Long, amts(0 To 1) As Variant
    ids(0) = fvA: amts(0) = delta
    ids(1) = unA: amts(1) = -delta
    Post ts, "revalue", "mark-to-market " & asset, entity, ids, amts, 2
    fvadj(k) = target
End Sub

'----- input reading -------------------------------------------------
Private Function LastRow(ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Public Sub Engine_PostOpening()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("OpeningBalances")
    Dim r As Long, lr As Long: lr = LastRow(ws)
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then GoTo nxt
        Dim entity As String, kind As String
        entity = CStr(ws.Cells(r, 1).Value)
        kind = UCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        ' cols: entity, kind, ccy_or_asset, exchange, qty, unit_cost, amount, counterparty, as_of, memo
        Dim ca As String, ex As String, asof As String, cp As String
        ca = CStr(ws.Cells(r, 3).Value)
        ex = CStr(ws.Cells(r, 4).Value)
        asof = CStr(ws.Cells(r, 9).Value)
        cp = CStr(ws.Cells(r, 8).Value)
        Select Case kind
            Case "CASH": DoOpenCash entity, ca, ws.Cells(r, 7).Value, asof
            Case "CRYPTO": DoOpenCrypto entity, ca, ex, ws.Cells(r, 5).Value, ws.Cells(r, 6).Value, asof
            Case "LOAN": DoOpenLoan entity, ca, ws.Cells(r, 7).Value, cp, asof
        End Select
nxt:
    Next r
End Sub

' Merge Trades+Transfers+Bank into one timestamp-sorted event list, then post.
Public Sub Engine_PostHistory()
    Dim ev As Collection: Set ev = New Collection
    Dim ws As Worksheet, r As Long, lr As Long
    ' Trades: entity, side, base, quote, qty, price, fee, ts
    Set ws = ThisWorkbook.Sheets("Trades"): lr = LastRow(ws)
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            ev.Add Array(CStr(ws.Cells(r, 8).Value), "T", _
                CStr(ws.Cells(r, 1).Value), CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), _
                CStr(ws.Cells(r, 4).Value), ws.Cells(r, 5).Value, ws.Cells(r, 6).Value, ws.Cells(r, 7).Value)
        End If
    Next r
    ' Transfers: entity, asset, from, to, qty, fee, ts
    Set ws = ThisWorkbook.Sheets("Transfers"): lr = LastRow(ws)
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            ev.Add Array(CStr(ws.Cells(r, 7).Value), "X", _
                CStr(ws.Cells(r, 1).Value), CStr(ws.Cells(r, 2).Value), CStr(ws.Cells(r, 3).Value), _
                CStr(ws.Cells(r, 4).Value), ws.Cells(r, 5).Value, ws.Cells(r, 6).Value, "")
        End If
    Next r
    ' Bank: entity, ts, ccy, amount, counter_account, memo
    Set ws = ThisWorkbook.Sheets("BankStatement"): lr = LastRow(ws)
    For r = 2 To lr
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            ev.Add Array(CStr(ws.Cells(r, 2).Value), "B", _
                CStr(ws.Cells(r, 1).Value), CStr(ws.Cells(r, 3).Value), ws.Cells(r, 4).Value, _
                CStr(ws.Cells(r, 5).Value), CStr(ws.Cells(r, 6).Value), "", "")
        End If
    Next r

    ' stable insertion sort by ts (ISO8601 sorts lexicographically).
    Dim arr() As Variant, nE As Long: nE = ev.Count
    If nE = 0 Then Exit Sub
    ReDim arr(1 To nE)
    Dim i As Long: For i = 1 To nE: arr(i) = ev(i): Next i
    Dim j As Long, key As Variant
    For i = 2 To nE
        key = arr(i): j = i - 1
        Do While j >= 1
            If CStr(arr(j)(0)) > CStr(key(0)) Then arr(j + 1) = arr(j): j = j - 1 Else Exit Do
        Loop
        arr(j + 1) = key
    Next i

    For i = 1 To nE
        Dim a As Variant: a = arr(i)
        Select Case a(1)
            Case "T": DoTrade CStr(a(2)), CStr(a(3)), CStr(a(4)), CStr(a(5)), a(6), a(7), a(8), CStr(a(0))
            Case "X": DoTransfer CStr(a(2)), CStr(a(3)), CStr(a(4)), CStr(a(5)), a(6), a(7), CStr(a(0))
            Case "B": DoBank CStr(a(2)), CStr(a(0)), CStr(a(3)), a(4), CStr(a(5)), CStr(a(6))
        End Select
    Next i
End Sub

' Period-end mark-to-market for every (entity, asset) still held. Mark =
' Marks-sheet override if present, else the last trade price for that asset.
Public Sub Engine_RevalueAll()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, L As Variant
    For i = 1 To lots.Count
        L = lots(i)
        Dim k As String: k = L(1) & "|" & L(2)
        If Not seen.Exists(k) And L(6) > CDec(0) Then
            seen(k) = 1
            DoRevalue CStr(L(1)), CStr(L(2)), MarkForPub(CStr(L(1)), CStr(L(2))), "9999-12-31T00:00:00Z"
        End If
    Next i
End Sub

Public Function MarkForPub(entity As String, asset As String) As Variant
    ' 1) explicit override on Marks sheet (asset, price[, entity])
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Marks")
    On Error GoTo 0
    If Not ws Is Nothing Then
        ' cols: asset, price, entity(optional). Entity-specific row wins over a
        ' blank-entity global default; keeps a USD mark off an IDR-quoted entity.
        Dim r As Long, lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim glob As Variant: Dim haveGlob As Boolean: haveGlob = False
        For r = 2 To lr
            If CStr(ws.Cells(r, 1).Value) = asset Then
                Dim me_ As String: me_ = Trim$(CStr(ws.Cells(r, 3).Value))
                If me_ = entity Then MarkForPub = D(ws.Cells(r, 2).Value): Exit Function
                If me_ = "" Then glob = D(ws.Cells(r, 2).Value): haveGlob = True
            End If
        Next r
        If haveGlob Then MarkForPub = glob: Exit Function
    End If
    ' 2) last trade price for this asset in this entity
    Dim last As Variant: last = CDec(0): Dim found As Boolean: found = False
    Dim t As Worksheet: Set t = ThisWorkbook.Sheets("Trades")
    Dim rr As Long, lrr As Long: lrr = t.Cells(t.Rows.Count, 1).End(xlUp).Row
    For rr = 2 To lrr
        If CStr(t.Cells(rr, 1).Value) = entity And CStr(t.Cells(rr, 3).Value) = asset Then
            last = D(t.Cells(rr, 6).Value): found = True   ' price col
        End If
    Next rr
    If found Then MarkForPub = last Else MarkForPub = CDec(0)
End Function

'----- dump facts to sheets -----------------------------------------
Public Sub Engine_DumpFacts()
    Dim ws As Worksheet
    Set ws = Sheet_Reset("ChartOfAccounts", Array("id", "entity", "code", "name", "acct_type", "currency"))
    Dim k As Variant, r As Long: r = 2
    Dim keys() As Variant: keys = accById.keys
    Dim i As Long
    For i = 0 To accById.Count - 1
        Dim id As Long: id = keys(i)
        Dim a As Variant: a = accById(id)
        ws.Cells(r, 1) = id: ws.Cells(r, 2) = a(0): ws.Cells(r, 3) = a(1)
        ws.Cells(r, 4) = a(2): ws.Cells(r, 5) = a(3): ws.Cells(r, 6) = a(4)
        r = r + 1
    Next i

    Set ws = Sheet_Reset("Journal", Array("id", "ts", "kind", "memo", "entity"))
    r = 2
    For i = 1 To jr.Count
        Dim jjj As Variant: jjj = jr(i)
        ws.Cells(r, 1) = jjj(0): ws.Cells(r, 2) = jjj(1): ws.Cells(r, 3) = jjj(2)
        ws.Cells(r, 4) = jjj(3): ws.Cells(r, 5) = jjj(4): r = r + 1
    Next i

    Set ws = Sheet_Reset("JournalLines", Array("id", "journal_id", "account_id", "amount"))
    r = 2
    For i = 1 To jl.Count
        Dim x As Variant: x = jl(i)
        ws.Cells(r, 1) = i: ws.Cells(r, 2) = x(0): ws.Cells(r, 3) = x(1)
        ws.Cells(r, 4) = CDbl(x(2)): r = r + 1
    Next i

    Set ws = Sheet_Reset("Lots", Array("id", "entity", "asset", "exchange", "account_id", "acquired_ts", "qty_remaining", "unit_cost", "source_journal"))
    r = 2
    For i = 1 To lots.Count
        Dim L As Variant: L = lots(i)
        If L(6) > CDec(0) Then
            ws.Cells(r, 1) = L(0): ws.Cells(r, 2) = L(1): ws.Cells(r, 3) = L(2)
            ws.Cells(r, 4) = L(3): ws.Cells(r, 5) = L(4): ws.Cells(r, 6) = L(5)
            ws.Cells(r, 7) = CDbl(L(6)): ws.Cells(r, 8) = CDbl(L(7)): ws.Cells(r, 9) = L(8)
            r = r + 1
        End If
    Next i
End Sub

Public Function Engine_Imbalance() As Double
    Dim s As Variant: s = CDec(0): Dim i As Long
    For i = 1 To jl.Count: s = s + jl(i)(2): Next i
    Engine_Imbalance = CDbl(s)
End Function

' Clear a sheet (create if missing) and write a header row. Returns the sheet.
Public Function Sheet_Reset(nm As String, headers As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
    End If
    ws.Cells.Clear
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
        ws.Cells(1, c + 1).Font.Bold = True
    Next c
    Set Sheet_Reset = ws
End Function
"#;

const MOD_REPORTS: &str = r#"Attribute VB_Name = "modReports"
Option Explicit
Public gReportEngine As String   ' "SQL (ACE OLEDB)" or "VBA fallback"

' Build every report sheet. PURE VBA by default — the whole engine runs end to
' end inside Excel with no external provider, database, or add-in required.
'
' (Optional: an equivalent SQL implementation lives in Reports_Sql below. To use
'  it instead, set USE_SQL=True; it needs the Microsoft.ACE.OLEDB provider. It is
'  off by default precisely so the workbook is 100% self-contained VBA.)
Public Const USE_SQL As Boolean = False

Public Sub Reports_BuildAll()
    If USE_SQL Then
        On Error GoTo useFallback
        Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
                ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        gReportEngine = "SQL (ACE OLEDB)"
        Reports_Sql cn
        cn.Close
        Exit Sub
    End If
useFallback:
    gReportEngine = "VBA (pure — no external provider)"
    Reports_Vba
End Sub

'==================== SQL PATH ======================================
Private Sub Reports_Sql(cn As Object)
    ' Net signed balance per account (the one heavy GROUP BY + JOIN).
    Dim sql As String
    sql = "SELECT a.entity, a.code, a.acct_type, a.currency, " & _
          "SUM(jl.amount) AS net " & _
          "FROM ([JournalLines$] AS jl INNER JOIN [ChartOfAccounts$] AS a " & _
          "ON jl.account_id = a.id) GROUP BY a.entity, a.code, a.acct_type, a.currency"
    Dim rs As Object: Set rs = cn.Execute(sql)

    Dim bws As Worksheet: Set bws = Sheet_Reset("Balances", Array("entity", "code", "acct_type", "currency", "net_signed", "display"))
    Dim tws As Worksheet: Set tws = Sheet_Reset("TrialBalance", Array("entity", "code", "acct_type", "debit", "credit"))
    Dim br As Long: br = 2: Dim tr As Long: tr = 2
    Dim tD As Double, tC As Double
    ' aggregates for income statement / balance sheet keyed by entity
    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        Dim ent As String, code As String, ty As String, ccy As String, net As Double
        ent = CStr(rs.Fields(0).Value): code = CStr(rs.Fields(1).Value)
        ty = CStr(rs.Fields(2).Value): ccy = CStr(rs.Fields(3).Value)
        net = NzD(rs.Fields(4).Value)
        Dim disp As Double: disp = IIf(IsDebitNormal(ty), net, -net)
        bws.Cells(br, 1) = ent: bws.Cells(br, 2) = code: bws.Cells(br, 3) = ty
        bws.Cells(br, 4) = ccy: bws.Cells(br, 5) = net: bws.Cells(br, 6) = disp: br = br + 1
        Dim dr As Double, cr As Double
        If net > 0 Then dr = net Else dr = 0
        If net < 0 Then cr = -net Else cr = 0
        tws.Cells(tr, 1) = ent: tws.Cells(tr, 2) = code: tws.Cells(tr, 3) = ty
        tws.Cells(tr, 4) = dr: tws.Cells(tr, 5) = cr: tr = tr + 1
        tD = tD + dr: tC = tC + cr
        AccumEntity agg, ent, ty, code, net, disp
        rs.MoveNext
    Loop
    rs.Close
    tws.Cells(tr, 2) = "TOTAL": tws.Cells(tr, 4) = tD: tws.Cells(tr, 5) = tC
    tws.Cells(tr + 1, 2) = "balanced": tws.Cells(tr + 1, 4) = IIf(Abs(tD - tC) <= 0.005 + 0.000000001 * (Abs(tD) + Abs(tC)), "TRUE", "FALSE")

    ' Exchange balances by coin, and positions by asset (GROUP BY on Lots).
    LotsReport cn, "ExchangeBalances", "entity, exchange, asset", Array("entity", "exchange", "asset")
    LotsReport cn, "Positions", "entity, asset", Array("entity", "asset")

    EmitDerived agg          ' Loans, PnL, IncomeStatement, BalanceSheet
End Sub

Private Sub LotsReport(cn As Object, nm As String, grp As String, hdr As Variant)
    Dim sql As String
    sql = "SELECT " & grp & ", SUM(qty_remaining) AS qty, " & _
          "SUM(qty_remaining*unit_cost) AS cost FROM [Lots$] GROUP BY " & grp
    Dim rs As Object
    On Error Resume Next
    Set rs = cn.Execute(sql)
    On Error GoTo 0
    Dim full As Variant
    If UBound(hdr) = 2 Then
        full = Array(hdr(0), hdr(1), hdr(2), "qty", "cost_basis", "avg_cost", "mark", "market_value", "unrealized")
    Else
        full = Array(hdr(0), hdr(1), "qty", "cost_basis", "avg_cost", "mark", "market_value", "unrealized")
    End If
    Dim ws As Worksheet: Set ws = Sheet_Reset(nm, full)
    Dim r As Long: r = 2
    If rs Is Nothing Then Exit Sub
    Do While Not rs.EOF
        Dim base As Long: base = UBound(hdr) + 1   ' number of group cols
        Dim ent As String, asset As String
        Dim c As Long
        For c = 0 To UBound(hdr)
            ws.Cells(r, c + 1) = CStr(rs.Fields(c).Value)
        Next c
        ent = CStr(rs.Fields(0).Value)
        asset = CStr(rs.Fields(UBound(hdr)).Value)
        Dim qty As Double, cost As Double
        qty = NzD(rs.Fields(base).Value): cost = NzD(rs.Fields(base + 1).Value)
        Dim avg As Double: If qty <> 0 Then avg = cost / qty
        Dim mk As Double: mk = CDbl(MarkForPub(ent, asset))
        Dim mv As Double: mv = qty * mk
        ws.Cells(r, base + 1) = qty: ws.Cells(r, base + 2) = cost: ws.Cells(r, base + 3) = avg
        ws.Cells(r, base + 4) = mk: ws.Cells(r, base + 5) = mv: ws.Cells(r, base + 6) = mv - cost
        r = r + 1: rs.MoveNext
    Loop
    rs.Close
End Sub

'==================== VBA FALLBACK ==================================
Private Sub Reports_Vba()
    ' Re-derive net-per-account straight from the JournalLines/ChartOfAccounts
    ' sheets we just wrote (no SQL provider available).
    Dim coa As Worksheet: Set coa = ThisWorkbook.Sheets("ChartOfAccounts")
    Dim jlw As Worksheet: Set jlw = ThisWorkbook.Sheets("JournalLines")
    Dim id2 As Object: Set id2 = CreateObject("Scripting.Dictionary") ' id -> Array(entity,code,type,ccy)
    Dim lr As Long, r As Long
    lr = coa.Cells(coa.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        id2(CLng(coa.Cells(r, 1).Value)) = Array(CStr(coa.Cells(r, 2).Value), CStr(coa.Cells(r, 3).Value), CStr(coa.Cells(r, 5).Value), CStr(coa.Cells(r, 6).Value))
    Next r
    Dim net As Object: Set net = CreateObject("Scripting.Dictionary")
    lr = jlw.Cells(jlw.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        Dim acc As Long: acc = CLng(jlw.Cells(r, 3).Value)
        net(acc) = NzD(net.Item(acc)) + CDbl(jlw.Cells(r, 4).Value)
    Next r

    Dim bws As Worksheet: Set bws = Sheet_Reset("Balances", Array("entity", "code", "acct_type", "currency", "net_signed", "display"))
    Dim tws As Worksheet: Set tws = Sheet_Reset("TrialBalance", Array("entity", "code", "acct_type", "debit", "credit"))
    Dim br As Long: br = 2: Dim tr As Long: tr = 2: Dim tD As Double, tC As Double
    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In id2.keys
        Dim a As Variant: a = id2(k)
        Dim nv As Double: nv = NzD(net.Item(k))
        Dim disp As Double: disp = IIf(IsDebitNormal(a(2)), nv, -nv)
        bws.Cells(br, 1) = a(0): bws.Cells(br, 2) = a(1): bws.Cells(br, 3) = a(2)
        bws.Cells(br, 4) = a(3): bws.Cells(br, 5) = nv: bws.Cells(br, 6) = disp: br = br + 1
        Dim dr As Double, cr As Double: dr = 0: cr = 0
        If nv > 0 Then dr = nv
        If nv < 0 Then cr = -nv
        tws.Cells(tr, 1) = a(0): tws.Cells(tr, 2) = a(1): tws.Cells(tr, 3) = a(2)
        tws.Cells(tr, 4) = dr: tws.Cells(tr, 5) = cr: tr = tr + 1: tD = tD + dr: tC = tC + cr
        AccumEntity agg, CStr(a(0)), CStr(a(2)), CStr(a(1)), nv, disp
    Next k
    tws.Cells(tr, 2) = "TOTAL": tws.Cells(tr, 4) = tD: tws.Cells(tr, 5) = tC
    tws.Cells(tr + 1, 2) = "balanced": tws.Cells(tr + 1, 4) = IIf(Abs(tD - tC) <= 0.005 + 0.000000001 * (Abs(tD) + Abs(tC)), "TRUE", "FALSE")

    LotsReportVba "ExchangeBalances", True
    LotsReportVba "Positions", False
    EmitDerived agg
End Sub

Private Sub LotsReportVba(nm As String, byExchange As Boolean)
    Dim lw As Worksheet: Set lw = ThisWorkbook.Sheets("Lots")
    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim lr As Long, r As Long: lr = lw.Cells(lw.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        Dim ent As String, ex As String, asset As String, key As String
        ent = CStr(lw.Cells(r, 2).Value): asset = CStr(lw.Cells(r, 3).Value): ex = CStr(lw.Cells(r, 4).Value)
        If byExchange Then key = ent & "|" & ex & "|" & asset Else key = ent & "|" & asset
        Dim q As Double, c As Double
        q = CDbl(lw.Cells(r, 7).Value): c = CDbl(lw.Cells(r, 7).Value) * CDbl(lw.Cells(r, 8).Value)
        Dim cur As Variant
        If agg.Exists(key) Then cur = agg(key) Else cur = Array(0#, 0#)
        agg(key) = Array(cur(0) + q, cur(1) + c)
    Next r
    Dim full As Variant
    If byExchange Then full = Array("entity", "exchange", "asset", "qty", "cost_basis", "avg_cost", "mark", "market_value", "unrealized") _
    Else full = Array("entity", "asset", "qty", "cost_basis", "avg_cost", "mark", "market_value", "unrealized")
    Dim ws As Worksheet: Set ws = Sheet_Reset(nm, full)
    Dim rr As Long: rr = 2
    Dim k As Variant
    For Each k In agg.keys
        Dim parts() As String: parts = Split(CStr(k), "|")
        Dim v As Variant: v = agg(k)
        Dim base As Long
        If byExchange Then
            ws.Cells(rr, 1) = parts(0): ws.Cells(rr, 2) = parts(1): ws.Cells(rr, 3) = parts(2): base = 3
        Else
            ws.Cells(rr, 1) = parts(0): ws.Cells(rr, 2) = parts(1): base = 2
        End If
        Dim ent As String, asset As String: ent = parts(0): asset = parts(UBound(parts))
        Dim qty As Double, cost As Double: qty = v(0): cost = v(1)
        Dim avg As Double: If qty <> 0 Then avg = cost / qty
        Dim mk As Double: mk = CDbl(MarkForPub(ent, asset))
        ws.Cells(rr, base + 1) = qty: ws.Cells(rr, base + 2) = cost: ws.Cells(rr, base + 3) = avg
        ws.Cells(rr, base + 4) = mk: ws.Cells(rr, base + 5) = qty * mk: ws.Cells(rr, base + 6) = qty * mk - cost
        rr = rr + 1
    Next k
End Sub

'==================== shared derivation =============================
' agg(entity) = Array(income_display_sum, expense_display_sum, asset, liab, equity, loanNet)
Private Sub AccumEntity(agg As Object, ent As String, ty As String, code As String, net As Double, disp As Double)
    Dim a As Variant
    If agg.Exists(ent) Then a = agg(ent) Else a = Array(0#, 0#, 0#, 0#, 0#, 0#)
    Select Case ty
        Case "Income": a(0) = a(0) + disp
        Case "Expense": a(1) = a(1) + disp
        Case "Asset": a(2) = a(2) + disp
        Case "Liability": a(3) = a(3) + disp
        Case "Equity": a(4) = a(4) + disp
    End Select
    If code = "LIAB:LOAN" Then a(5) = a(5) + (-net)   ' outstanding principal (credit-normal)
    agg(ent) = a
End Sub

Private Sub EmitDerived(agg As Object)
    Dim ln As Worksheet: Set ln = Sheet_Reset("Loans", Array("entity", "outstanding"))
    Dim pn As Worksheet: Set pn = Sheet_Reset("PnL", Array("entity", "realized", "unrealized", "fees"))
    Dim is_ As Worksheet: Set is_ = Sheet_Reset("IncomeStatement", Array("entity", "realized", "unrealized", "fees", "net"))
    Dim bs As Worksheet: Set bs = Sheet_Reset("BalanceSheet", Array("entity", "assets", "liabilities", "equity", "retained", "balanced"))
    Dim r As Long: r = 2
    Dim k As Variant
    For Each k In agg.keys
        Dim a As Variant: a = agg(k)
        Dim income As Double, expense As Double, assets As Double, liab As Double, equity As Double, loan As Double
        income = a(0): expense = a(1): assets = a(2): liab = a(3): equity = a(4): loan = a(5)
        Dim retained As Double: retained = income - expense
        ' PnL: income accounts hold realized+unrealized as display (credit-normal -> positive gain)
        ' We split via the dedicated accounts using the Balances sheet lookups.
        Dim realized As Double, unrealized As Double, fees As Double
        realized = LookupDisp(k, "PNL:REALIZED")
        unrealized = LookupDisp(k, "PNL:UNREALIZED")
        fees = LookupDisp(k, "EXP:FEES:TRADING") + LookupDisp(k, "EXP:FEES:NETWORK")
        ln.Cells(r, 1) = k: ln.Cells(r, 2) = loan
        pn.Cells(r, 1) = k: pn.Cells(r, 2) = realized: pn.Cells(r, 3) = unrealized: pn.Cells(r, 4) = fees
        is_.Cells(r, 1) = k: is_.Cells(r, 2) = realized: is_.Cells(r, 3) = unrealized: is_.Cells(r, 4) = fees: is_.Cells(r, 5) = realized + unrealized - fees
        bs.Cells(r, 1) = k: bs.Cells(r, 2) = assets: bs.Cells(r, 3) = liab: bs.Cells(r, 4) = equity: bs.Cells(r, 5) = retained
        Dim rhs As Double: rhs = liab + equity + retained
        bs.Cells(r, 6) = IIf(Abs(assets - rhs) <= 0.005 + 0.000000001 * (Abs(assets) + Abs(rhs)), "TRUE", "FALSE")
        r = r + 1
    Next k
End Sub

' display of one account code for an entity, read from the Balances sheet.
Private Function LookupDisp(entity As String, code As String) As Double
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Balances")
    Dim lr As Long, r As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        If CStr(ws.Cells(r, 1).Value) = entity And CStr(ws.Cells(r, 2).Value) = code Then
            LookupDisp = CDbl(ws.Cells(r, 6).Value): Exit Function
        End If
    Next r
    LookupDisp = 0
End Function

Private Function NzD(v As Variant) As Double
    If IsNull(v) Or IsEmpty(v) Then NzD = 0 Else NzD = CDbl(v)
End Function

'==================== BANK RECONCILIATION ===========================
' Match every BankStatement line to its posted cash journal (kind='bank'),
' then tie out the statement total vs posted cash per entity+currency.
Public Sub Recon_Build()
    Dim coa As Worksheet: Set coa = ThisWorkbook.Sheets("ChartOfAccounts")
    Dim jw As Worksheet: Set jw = ThisWorkbook.Sheets("Journal")
    Dim jlw As Worksheet: Set jlw = ThisWorkbook.Sheets("JournalLines")
    Dim r As Long, lr As Long

    ' account_id -> is a CASH:* account?
    Dim isCash As Object: Set isCash = CreateObject("Scripting.Dictionary")
    lr = coa.Cells(coa.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        If Left$(CStr(coa.Cells(r, 3).Value), 5) = "CASH:" Then isCash(CLng(coa.Cells(r, 1).Value)) = True
    Next r

    ' journal_id -> ts / kind / entity
    Dim jts As Object: Set jts = CreateObject("Scripting.Dictionary")
    Dim jkind As Object: Set jkind = CreateObject("Scripting.Dictionary")
    Dim jent As Object: Set jent = CreateObject("Scripting.Dictionary")
    lr = jw.Cells(jw.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        Dim jid As Long: jid = CLng(jw.Cells(r, 1).Value)
        jts(jid) = CStr(jw.Cells(r, 2).Value): jkind(jid) = CStr(jw.Cells(r, 3).Value): jent(jid) = CStr(jw.Cells(r, 5).Value)
    Next r

    ' posted bank cash movements keyed entity|ts|amount -> journal_id
    Dim posted As Object: Set posted = CreateObject("Scripting.Dictionary")
    lr = jlw.Cells(jlw.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        Dim jj As Long: jj = CLng(jlw.Cells(r, 2).Value)
        Dim acc As Long: acc = CLng(jlw.Cells(r, 3).Value)
        If isCash.Exists(acc) Then
            If jkind(jj) = "bank" Then
                posted(jent(jj) & "|" & jts(jj) & "|" & Format(CDbl(jlw.Cells(r, 4).Value), "0.##########")) = jj
            End If
        End If
    Next r

    Dim ws As Worksheet: Set ws = Sheet_Reset("Reconciliation", Array("entity", "ts", "ccy", "statement_amount", "posted_amount", "status", "journal_id"))
    Dim bk As Worksheet: Set bk = ThisWorkbook.Sheets("BankStatement")
    Dim wr As Long: wr = 2
    Dim totS As Object: Set totS = CreateObject("Scripting.Dictionary")
    Dim totP As Object: Set totP = CreateObject("Scripting.Dictionary")
    Dim unmatched As Long: unmatched = 0
    lr = bk.Cells(bk.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lr
        If Trim$(CStr(bk.Cells(r, 1).Value)) <> "" Then
            Dim e As String, ts As String, ccy As String, sa As Double
            e = CStr(bk.Cells(r, 1).Value): ts = CStr(bk.Cells(r, 2).Value)
            ccy = CStr(bk.Cells(r, 3).Value): sa = CDbl(bk.Cells(r, 4).Value)
            Dim k As String: k = e & "|" & ts & "|" & Format(sa, "0.##########")
            Dim tk As String: tk = e & "|" & ccy
            totS(tk) = NzD(totS.Item(tk)) + sa
            If posted.Exists(k) Then
                ws.Cells(wr, 5) = sa: ws.Cells(wr, 6) = "MATCHED": ws.Cells(wr, 7) = posted(k)
                totP(tk) = NzD(totP.Item(tk)) + sa
            Else
                ws.Cells(wr, 5) = 0: ws.Cells(wr, 6) = "UNMATCHED": unmatched = unmatched + 1
            End If
            ws.Cells(wr, 1) = e: ws.Cells(wr, 2) = ts: ws.Cells(wr, 3) = ccy: ws.Cells(wr, 4) = sa
            wr = wr + 1
        End If
    Next r

    ' tie-out footer per entity+ccy
    wr = wr + 1
    ws.Cells(wr, 1) = "TIE-OUT": ws.Cells(wr, 3) = "ccy": ws.Cells(wr, 4) = "stmt_total": ws.Cells(wr, 5) = "posted_total": ws.Cells(wr, 6) = "status"
    ws.Cells(wr, 1).Font.Bold = True
    wr = wr + 1
    Dim key As Variant
    For Each key In totS.keys
        Dim parts() As String: parts = Split(CStr(key), "|")
        Dim s As Double, p As Double: s = NzD(totS(key)): p = NzD(totP.Item(key))
        ws.Cells(wr, 1) = parts(0): ws.Cells(wr, 3) = parts(1): ws.Cells(wr, 4) = s: ws.Cells(wr, 5) = p
        ws.Cells(wr, 6) = IIf(Abs(s - p) <= 0.005 + 0.000000001 * (Abs(s) + Abs(p)), "TIED", "CHECK")
        wr = wr + 1
    Next key
    ws.Cells(wr + 1, 1) = "Unmatched lines: " & unmatched
End Sub
"#;
