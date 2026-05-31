Attribute VB_Name = "modReports"
Option Explicit
Public gReportEngine As String   ' "SQL (ACE OLEDB)" or "VBA fallback"

' Build every report sheet. Tries SQL via ADODB/ACE over the saved workbook;
' on any provider error, falls back to equivalent in-VBA aggregation so RunAll
' always completes end to end.
Public Sub Reports_BuildAll()
    Dim cn As Object
    On Error GoTo useFallback
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
            ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    gReportEngine = "SQL (ACE OLEDB)"
    Reports_Sql cn
    cn.Close
    Exit Sub
useFallback:
    gReportEngine = "VBA fallback (ACE provider not found)"
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
