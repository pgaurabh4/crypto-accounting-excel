Attribute VB_Name = "modEngine"
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
