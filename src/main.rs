//! Generator for the Gravity Ledger master workbook.
//!
//! Emits into `dist/`:
//!   * GravityLedger.xlsx  — the workbook (4 input sheets pre-filled with the
//!                            Rust engine's demo scenario + all report sheets).
//!   * vba/*.bas           — the VBA engine (SQL reports + procedural FIFO).
//!   * Add-Macros.vbs      — double-click on Windows/Excel to produce the
//!                            macro-enabled GravityLedger.xlsm (Excel writes the
//!                            vbaProject.bin, so it is guaranteed valid).
//!   * SETUP.md            — three ways to reach a working macro workbook.

mod docx;
mod vba;
mod xlsx;

use docx::Block;
use std::fs;
use xlsx::{e, n, t, Cell, Sheet};

fn header(cols: &[&str]) -> Vec<Cell> {
    cols.iter().map(|c| t(*c)).collect()
}

fn main() -> std::io::Result<()> {
    fs::create_dir_all("dist/vba")?;
    let sheets = build_sheets();
    xlsx::write_xlsx("dist/GravityLedger.xlsx", &sheets)?;

    for (name, src) in vba::modules() {
        fs::write(format!("dist/vba/{name}.bas"), src)?;
    }
    fs::write("dist/Add-Macros.vbs", build_vbs())?;
    fs::write("dist/SETUP.md", build_setup())?;
    docx::write_docx("dist/GravityLedger-VBA-Explained.docx", &build_doc())?;

    pack_zip("dist/GravityLedger-workbook.zip")?;

    println!("wrote dist/GravityLedger.xlsx ({} sheets)", sheets.len());
    println!("wrote dist/vba/{{modMain,modEngine,modReports}}.bas");
    println!("wrote dist/Add-Macros.vbs");
    println!("wrote dist/SETUP.md");
    println!("wrote dist/GravityLedger-VBA-Explained.docx");
    println!("wrote dist/GravityLedger-workbook.zip (deliverable bundle)");
    Ok(())
}

/// The Word document: a complete walkthrough of how the VBA engine works.
fn build_doc() -> Vec<Block> {
    use Block::*;
    let mut d = Vec::new();
    d.push(Title("Gravity Ledger — How the VBA Engine Works".into()));
    d.push(Subtitle(
        "A single-file crypto double-entry accounting workbook, 100% VBA macros, end to end."
            .into(),
    ));

    d.push(H1("1. The big picture".into()));
    d.push(Para("Gravity Ledger is an Excel workbook that keeps the books for a spot-only crypto market maker: double-entry general ledger, FIFO cost basis, inventory transfers between exchanges, period-end mark-to-market, loans, and IFRS reports — across multiple legal entities. It is a faithful port of the Rust engine in pgaurabh4/crypto-accounting.".into()));
    d.push(Para("You touch only four input sheets. Everything else is rebuilt by one macro, RunAll, wired to the RUN ALL button. There is no database, add-in, or external provider — the entire engine is VBA, so it runs on any machine that has Excel.".into()));
    d.push(Para(
        "The code is organised into three standard modules:".into(),
    ));
    d.push(Bullet(
        "modMain — the RunAll orchestration and the button handler.".into(),
    ));
    d.push(Bullet(
        "modEngine — ingestion + double-entry posting + FIFO lots (the sequential accounting)."
            .into(),
    ));
    d.push(Bullet(
        "modReports — aggregation of every report sheet, plus bank reconciliation.".into(),
    ));

    d.push(H1("2. The RunAll pipeline".into()));
    d.push(Para(
        "Clicking RUN ALL executes these steps in order:".into(),
    ));
    d.push(Code(vec![
        "Engine_Reset       ' clear in-memory ledger state".into(),
        "Engine_PostOpening ' opening-balance snapshot -> opening journal + seed lots".into(),
        "Engine_PostHistory ' merge Trades+Transfers+BankStatement by timestamp, then post".into(),
        "Engine_RevalueAll  ' period-end mark-to-market for every held (entity, asset)".into(),
        "Engine_DumpFacts   ' write ChartOfAccounts / Journal / JournalLines / Lots".into(),
        "Reports_BuildAll   ' aggregate every report sheet (pure VBA)".into(),
        "Recon_Build        ' match BankStatement lines to posted cash journals".into(),
        "ThisWorkbook.Save  ' persist".into(),
    ]));
    d.push(Para("It finishes with a popup showing the run time and the global imbalance, which must be exactly 0 — the cardinal correctness property of a double-entry ledger.".into()));

    d.push(H1("3. The data model".into()));
    d.push(H2("Inputs you edit".into()));
    d.push(Bullet("OpeningBalances — an as-at cutover snapshot. Rows of kind CASH, CRYPTO, or LOAN; each posts an opening journal balanced against EQ:OPENING (Opening Balance Equity), and CRYPTO rows also seed a FIFO lot so cost basis is correct going forward.".into()));
    d.push(Bullet(
        "Trades — entity, side (buy/sell), base, quote, qty, price, fee, ts.".into(),
    ));
    d.push(Bullet("Transfers — move a coin between exchanges at cost (no P&L); a network fee is paid in-kind from the same inventory.".into()));
    d.push(Bullet("BankStatement — cash in/out. The counter_account column routes it: EQ:CAPITAL (funding), LIAB:LOAN (loan drawdown/repayment), INC:* (income), EXP:* (expense).".into()));
    d.push(Bullet("Marks (optional) — override the mark-to-market price per asset (and optionally per entity). If absent, the mark defaults to the last traded price for that asset.".into()));
    d.push(H2("Derived sheets (rebuilt each run)".into()));
    d.push(Bullet(
        "ChartOfAccounts, Journal, JournalLines, Lots — the facts of record.".into(),
    ));
    d.push(Bullet("TrialBalance, Balances, ExchangeBalances (balances by coin per exchange), Positions, Loans, PnL, IncomeStatement, BalanceSheet, Reconciliation — the reports.".into()));
    d.push(H2("Accounts and the sign convention".into()));
    d.push(Para("Every account has a type inferred from its code prefix: CASH:* and CRYPTO:* are Assets; LIAB:* a Liability; EQ:* Equity; PNL:*/INC:* Income; EXP:* Expense. Asset and Expense accounts are debit-normal; Liability, Equity and Income are credit-normal.".into()));
    d.push(Para("Journal lines store a single signed amount: a debit is positive, a credit is negative. A journal is valid only if its lines sum to exactly zero. An account's display balance is its signed net for debit-normal accounts, or the negation of it for credit-normal accounts.".into()));

    d.push(H1("4. modEngine — posting and FIFO".into()));
    d.push(Para("State is held in memory while a run executes (Dictionaries and Collections), then dumped to sheets. Money uses VBA's Decimal subtype via CDec, so large IDR notionals and fractional-crypto cost bases stay exact.".into()));
    d.push(H2("Posting a balanced journal".into()));
    d.push(Para("Every event builds a small set of lines and calls Post, which enforces the zero-sum invariant before writing the journal and its lines:".into()));
    d.push(Code(vec![
        "Private Function Post(ts, kind, memo, entity, ids(), amts(), nlines)".into(),
        "    s = sum(amts)                       ' Decimal".into(),
        "    If s <> 0 Then Err.Raise ...        ' refuse to post an unbalanced journal".into(),
        "    jId = jId + 1 : record journal header".into(),
        "    For each line: record (jId, account_id, amount)".into(),
        "End Function".into(),
    ]));
    d.push(H2("FIFO cost basis".into()));
    d.push(Para("ConsumeFifo removes a quantity from the oldest open lots of an account (ordered by acquisition timestamp, then lot id) and returns the cost basis consumed. It refuses to consume more than is on hand — spot-only means no shorting. Fully-consumed lots are emptied; a partially-consumed lot keeps its remaining quantity at the same unit cost.".into()));
    d.push(Code(vec![
        "avail = SUM(qty_remaining) for (entity, account)".into(),
        "If avail < qty Then Err.Raise \"insufficient inventory\"".into(),
        "For each lot oldest-first while remaining > 0:".into(),
        "    take = min(lot.qty, remaining)".into(),
        "    cost = cost + take * lot.unit_cost".into(),
        "    lot.qty = lot.qty - take ; remaining = remaining - take".into(),
    ]));
    d.push(H2("Event handlers (each mirrors the Rust engine)".into()));
    d.push(Bullet("Buy: debit CRYPTO inventory at notional (qty*price), debit fee expense, credit cash for (notional+fee). Add a lot at unit_cost = price (fees are expensed, not capitalised).".into()));
    d.push(Bullet("Sell: take FIFO cost basis out of inventory; realized = proceeds - cost_basis. Lines: debit cash (proceeds-fee), debit fee, credit inventory (cost_basis), credit realized P&L (Income).".into()));
    d.push(Bullet("Transfer: FIFO-consume the moved quantity (and any in-kind fee) out of the source exchange; debit the destination, credit the source, expense the fee at cost; re-establish a lot at the destination preserving the moved cost basis.".into()));
    d.push(Bullet("Bank: debit CASH for the signed amount, credit the counter account (or vice-versa for a withdrawal). A LIAB:LOAN counter makes a deposit a loan drawdown and a negative amount a repayment.".into()));
    d.push(Bullet("Opening cash/crypto/loan: post the snapshot against EQ:OPENING; crypto rows also seed a lot.".into()));
    d.push(Bullet("Revalue (mark-to-market): target adjustment = qty*mark - carrying cost; post the delta to CRYPTO:FVADJ and unrealized P&L.".into()));
    d.push(H2("Chronological ordering".into()));
    d.push(Para("Engine_PostHistory merges Trades, Transfers and BankStatement into one list and sorts by ISO-8601 timestamp (which sorts correctly as text), so FIFO consumption happens in true chronological order. Opening balances are always posted first; mark-to-market is posted last, at period end.".into()));

    d.push(H1("5. modReports — the reports".into()));
    d.push(Para("After the facts are written, Reports_BuildAll aggregates them entirely in VBA (using Dictionaries to group). The reports are:".into()));
    d.push(Bullet("Balances / TrialBalance — net signed amount per account; debit and credit columns; a balanced flag (total debits = total credits within a scale-relative tolerance).".into()));
    d.push(Bullet("ExchangeBalances — grouped by entity, exchange and asset: quantity, cost basis, average cost, mark, market value and unrealized gain. This is your 'balances by coin per exchange'.".into()));
    d.push(Bullet(
        "Positions — the same, grouped by asset across all venues.".into(),
    ));
    d.push(Bullet(
        "Loans — outstanding LIAB:LOAN principal per entity.".into(),
    ));
    d.push(Bullet("PnL / IncomeStatement — realized = -(net of PNL:REALIZED), unrealized = -(net of PNL:UNREALIZED), fees = trading + network; net income = realized + unrealized - fees.".into()));
    d.push(Bullet("BalanceSheet — assets, liabilities, equity, retained earnings (income - expenses); balanced when assets = liabilities + equity + retained.".into()));
    d.push(Para("An optional SQL implementation of this layer lives in modReports behind the constant USE_SQL = False; it uses the Microsoft.ACE.OLEDB provider. It is off by default so the workbook is fully self-contained.".into()));
    d.push(H2("Bank reconciliation".into()));
    d.push(Para("Recon_Build matches every BankStatement line to the cash journal the engine posted for it (same entity, timestamp and amount, kind = 'bank'), marking each MATCHED or UNMATCHED, then ties out the statement total against posted cash per entity and currency.".into()));

    d.push(H1("6. Worked example (shipped demo)".into()));
    d.push(Para("Entity BVI: 1,000,000 USD capital; buy 2 BTC @ 60,000 (fee 30); buy 1 BTC @ 65,000 (fee 20); buy 50 ETH @ 3,000 (fee 25); transfer 1.5 BTC MAIN->COLD (fee 0.0005); sell 1 BTC @ 70,000 (fee 35); marks BTC 72,000 / ETH 3,200.".into()));
    d.push(Para("The sell takes FIFO basis: after the transfer consumed 1.5 BTC + 0.0005 fee from the 60,000 lot, the 1 BTC sold is 0.4995 @ 60,000 + 0.5005 @ 65,000 = 62,502.50 cost. Realized = 70,000 - 62,502.50 = 7,497.50.".into()));
    d.push(Para("After RUN ALL, BVI shows: realized 7,497.50; unrealized 31,496.50; fees 140; net income 38,854; BalanceSheet balanced = TRUE; global imbalance 0. These match the Rust ledger to the cent.".into()));

    d.push(H1("7. Running and extending".into()));
    d.push(Bullet("Build the macro file: double-click Add-Macros.vbs (after enabling 'Trust access to the VBA project object model'), or Save-As .xlsm and Alt+F11 > Import the three .bas files.".into()));
    d.push(Bullet("Use it: edit the four input sheets, click RUN ALL. The whole ledger rebuilds deterministically every time.".into()));
    d.push(Bullet("Extend it: add a new event type by writing one Do... handler in modEngine plus an input sheet; add a report by adding one aggregation in modReports. The zero-sum Post invariant guarantees the books always balance.".into()));
    d
}

/// Bundle the deliverable into one zip, preserving the `vba/` folder so
/// Add-Macros.vbs finds the modules.
fn pack_zip(path: &str) -> std::io::Result<()> {
    use std::io::Write;
    let f = fs::File::create(path)?;
    let mut z = zip::ZipWriter::new(f);
    let opts: zip::write::FileOptions =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut add =
        |z: &mut zip::ZipWriter<fs::File>, name: &str, disk: &str| -> std::io::Result<()> {
            z.start_file(name, opts)?;
            z.write_all(&fs::read(disk)?)?;
            Ok(())
        };
    add(&mut z, "GravityLedger.xlsx", "dist/GravityLedger.xlsx")?;
    add(&mut z, "Add-Macros.vbs", "dist/Add-Macros.vbs")?;
    add(&mut z, "SETUP.md", "dist/SETUP.md")?;
    for m in ["modMain", "modEngine", "modReports"] {
        add(
            &mut z,
            &format!("vba/{m}.bas"),
            &format!("dist/vba/{m}.bas"),
        )?;
    }
    z.finish()?;
    Ok(())
}

/// Build every sheet: README, the 4 inputs (seed demo), Marks, and report stubs.
fn build_sheets() -> Vec<Sheet> {
    let mut out = Vec::new();

    // ---- README / Dashboard ------------------------------------------------
    let mut rd = Sheet::new("README");
    rd.widths(&[26.0, 90.0]);
    rd.title("GRAVITY LEDGER — crypto double-entry accounting (SQL + macros)");
    rd.row(vec![e()]);
    rd.row(vec![t("What this is"), t("A single-file ledger for a spot-only crypto market maker. Mirrors the Rust engine in pgaurabh4/crypto-accounting: double-entry, FIFO cost basis, transfers, mark-to-market, multi-entity.")]);
    rd.row(vec![
        t("You edit ONLY"),
        t("OpeningBalances, Trades, Transfers, BankStatement. Everything else is derived."),
    ]);
    rd.row(vec![t("To rebuild"), t("Click the RUN ALL button (added by Add-Macros.vbs). 100% VBA macros, end to end — posts journals + FIFO lots, builds every report, reconciles bank lines. Needs only Excel; no database or add-in.")]);
    rd.row(vec![e()]);
    rd.header(header(&["Input sheet", "Columns / meaning"]));
    rd.row(vec![t("OpeningBalances"), t("entity | kind(CASH/CRYPTO/LOAN) | ccy_or_asset | exchange | qty | unit_cost | amount | counterparty | as_of | memo. As-at cutover snapshot; balances against EQ:OPENING and seeds FIFO lots. (Empty in this demo — it starts from inception via the capital cash-in.)")]);
    rd.row(vec![t("Trades"), t("entity | side(buy/sell) | base | quote | qty | price | fee | ts. Buy debits inventory at cost; sell takes FIFO cost basis and books realized P&L.")]);
    rd.row(vec![t("Transfers"), t("entity | asset | from_exchange | to_exchange | qty | fee | ts. Moves coin between venues at cost (no P&L); network fee expensed in-kind.")]);
    rd.row(vec![t("BankStatement"), t("entity | ts | ccy | amount(+in/-out) | counter_account | memo. counter_account routes it: EQ:CAPITAL (funding), LIAB:LOAN (loan drawdown/repay), INC:* , EXP:* …")]);
    rd.row(vec![t("Marks"), t("asset | price | entity(optional). Overrides the mark-to-market price. If absent, mark = last traded price for that asset in that entity.")]);
    rd.row(vec![e()]);
    rd.header(header(&["Derived sheet", "Built by"]));
    for (s, by) in [
        ("ChartOfAccounts / Journal / JournalLines / Lots", "VBA: procedural posting + FIFO"),
        ("TrialBalance / Balances", "VBA: SUM(amount) per account"),
        ("ExchangeBalances", "VBA: group by entity, exchange, asset over Lots (your 'balances by coin per exchange')"),
        ("Positions", "VBA: group by entity, asset over Lots"),
        ("Loans", "VBA: outstanding LIAB:LOAN per entity"),
        ("PnL / IncomeStatement / BalanceSheet", "VBA aggregates + IFRS derivation"),
        ("Reconciliation", "VBA: match each BankStatement line to its posted cash journal, then tie out statement vs ledger per entity+ccy"),
    ] {
        rd.row(vec![t(s), t(by)]);
    }
    rd.row(vec![e()]);
    rd.row(vec![t("Demo check"), t("After RUN ALL, BVI should show realized 7497.5, unrealized 31496.5, fees 140, net 38854, and BalanceSheet balanced = TRUE. Global imbalance = 0.")]);
    rd.row(vec![
        t("[ RUN ALL ]"),
        t("<- the macro button is placed here by Add-Macros.vbs"),
    ]);
    out.push(rd);

    // ---- Inputs: the Rust seed scenario ------------------------------------
    let mut ob = Sheet::new("OpeningBalances");
    ob.widths(&[10.0, 10.0, 12.0, 12.0, 14.0, 16.0, 18.0, 16.0, 24.0, 30.0]);
    ob.header(header(&[
        "entity",
        "kind",
        "ccy_or_asset",
        "exchange",
        "qty",
        "unit_cost",
        "amount",
        "counterparty",
        "as_of",
        "memo",
    ]));
    // Intentionally empty for the demo. Example (documented in README), e.g.:
    //   IDX, CRYPTO, BTC, BINANCE, 1.0, 950000000, , , 2026-01-01T00:00:00Z, opening BTC
    out.push(ob);

    let mut tr = Sheet::new("Trades");
    tr.widths(&[10.0, 8.0, 8.0, 8.0, 12.0, 16.0, 10.0, 26.0]);
    tr.header(header(&[
        "entity", "side", "base", "quote", "qty", "price", "fee", "ts",
    ]));
    let trades: &[(&str, &str, &str, &str, f64, f64, f64, &str)] = &[
        (
            "BVI",
            "buy",
            "BTC",
            "USD",
            2.0,
            60000.0,
            30.0,
            "2026-01-02T00:00:00Z",
        ),
        (
            "BVI",
            "buy",
            "BTC",
            "USD",
            1.0,
            65000.0,
            20.0,
            "2026-01-03T00:00:00Z",
        ),
        (
            "BVI",
            "buy",
            "ETH",
            "USD",
            50.0,
            3000.0,
            25.0,
            "2026-01-04T00:00:00Z",
        ),
        (
            "BVI",
            "sell",
            "BTC",
            "USD",
            1.0,
            70000.0,
            35.0,
            "2026-01-06T00:00:00Z",
        ),
        (
            "IDX",
            "buy",
            "BTC",
            "IDR",
            0.5,
            1_000_000_000.0,
            100000.0,
            "2026-01-10T00:00:00Z",
        ),
    ];
    for (en, sd, b, q, qty, pr, fe, ts) in trades {
        tr.row(vec![
            t(*en),
            t(*sd),
            t(*b),
            t(*q),
            n(*qty),
            n(*pr),
            n(*fe),
            t(*ts),
        ]);
    }
    out.push(tr);

    let mut tx = Sheet::new("Transfers");
    tx.widths(&[10.0, 8.0, 14.0, 14.0, 12.0, 12.0, 26.0]);
    tx.header(header(&[
        "entity",
        "asset",
        "from_exchange",
        "to_exchange",
        "qty",
        "fee",
        "ts",
    ]));
    tx.row(vec![
        t("BVI"),
        t("BTC"),
        t("MAIN"),
        t("COLD"),
        n(1.5),
        n(0.0005),
        t("2026-01-05T00:00:00Z"),
    ]);
    out.push(tx);

    let mut bk = Sheet::new("BankStatement");
    bk.widths(&[10.0, 26.0, 8.0, 20.0, 18.0, 28.0]);
    bk.header(header(&[
        "entity",
        "ts",
        "ccy",
        "amount",
        "counter_account",
        "memo",
    ]));
    bk.row(vec![
        t("BVI"),
        t("2026-01-01T00:00:00Z"),
        t("USD"),
        n(1_000_000.0),
        t("EQ:CAPITAL"),
        t("capital contribution"),
    ]);
    bk.row(vec![
        t("IDX"),
        t("2026-01-09T00:00:00Z"),
        t("IDR"),
        n(10_000_000_000.0),
        t("EQ:CAPITAL"),
        t("capital contribution"),
    ]);
    out.push(bk);

    let mut mk = Sheet::new("Marks");
    mk.widths(&[10.0, 16.0, 10.0]);
    mk.header(header(&["asset", "price", "entity"]));
    mk.row(vec![t("BTC"), n(72000.0), t("BVI")]);
    mk.row(vec![t("ETH"), n(3200.0), t("BVI")]);
    out.push(mk);

    // ---- Report stubs (rebuilt by RUN ALL) ---------------------------------
    let stub = |name: &str, cols: &[&str]| {
        let mut s = Sheet::new(name);
        s.header(header(cols));
        s.row(vec![t("(click RUN ALL to populate)")]);
        s
    };
    out.push(stub(
        "ChartOfAccounts",
        &["id", "entity", "code", "name", "acct_type", "currency"],
    ));
    out.push(stub("Journal", &["id", "ts", "kind", "memo", "entity"]));
    out.push(stub(
        "JournalLines",
        &["id", "journal_id", "account_id", "amount"],
    ));
    out.push(stub(
        "Lots",
        &[
            "id",
            "entity",
            "asset",
            "exchange",
            "account_id",
            "acquired_ts",
            "qty_remaining",
            "unit_cost",
            "source_journal",
        ],
    ));
    out.push(stub(
        "TrialBalance",
        &["entity", "code", "acct_type", "debit", "credit"],
    ));
    out.push(stub(
        "Balances",
        &[
            "entity",
            "code",
            "acct_type",
            "currency",
            "net_signed",
            "display",
        ],
    ));
    out.push(stub(
        "ExchangeBalances",
        &[
            "entity",
            "exchange",
            "asset",
            "qty",
            "cost_basis",
            "avg_cost",
            "mark",
            "market_value",
            "unrealized",
        ],
    ));
    out.push(stub(
        "Positions",
        &[
            "entity",
            "asset",
            "qty",
            "cost_basis",
            "avg_cost",
            "mark",
            "market_value",
            "unrealized",
        ],
    ));
    out.push(stub("Loans", &["entity", "outstanding"]));
    out.push(stub("PnL", &["entity", "realized", "unrealized", "fees"]));
    out.push(stub(
        "IncomeStatement",
        &["entity", "realized", "unrealized", "fees", "net"],
    ));
    out.push(stub(
        "BalanceSheet",
        &[
            "entity",
            "assets",
            "liabilities",
            "equity",
            "retained",
            "balanced",
        ],
    ));
    out.push(stub(
        "Reconciliation",
        &[
            "entity",
            "ts",
            "ccy",
            "statement_amount",
            "posted_amount",
            "status",
            "journal_id",
        ],
    ));

    out
}

/// The VBScript builder: opens the .xlsx, imports the .bas modules, wires the
/// RUN ALL button, runs the engine once, and saves GravityLedger.xlsm.
fn build_vbs() -> String {
    r#"' Add-Macros.vbs — build GravityLedger.xlsm from GravityLedger.xlsx
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
"#.to_string()
}

fn build_setup() -> String {
    r#"# Gravity Ledger — setup

`GravityLedger.xlsx` is the master workbook. You edit four input sheets —
**OpeningBalances, Trades, Transfers, BankStatement** — and everything else
(chart of accounts, journals, FIFO lots, trial balance, exchange balances by
coin, positions, loans, P&L, balance sheet) is derived by the SQL + macro
engine when you click **RUN ALL**.

The macros live in `vba/*.bas`. Because a `.xlsm`'s macro container must be
written by Excel itself, pick whichever path below fits you.

## Option A — one double-click (recommended)
1. Keep `GravityLedger.xlsx`, `Add-Macros.vbs`, and the `vba/` folder together.
2. In Excel once: **File > Options > Trust Center > Trust Center Settings >
   Macro Settings >** tick **"Trust access to the VBA project object model."**
3. Double-click **`Add-Macros.vbs`**. It produces **`GravityLedger.xlsm`** with
   the engine embedded, the **RUN ALL** button wired, and reports already built.

## Option B — manual import (no .vbs)
1. Open `GravityLedger.xlsx`, then **Save As -> Excel Macro-Enabled Workbook
   (*.xlsm)**.
2. Press **Alt+F11** (VBA editor) -> **File > Import File...** -> import all three
   files in `vba/` (`modMain.bas`, `modEngine.bas`, `modReports.bas`).
3. Back in Excel: **Developer > Insert > Button**, assign it to macro `RunAll`.
4. Click the button. Done.

## 100% VBA — runs on just Excel
The entire engine is VBA macros. On RUN ALL it posts the double-entry journals
and FIFO lots, aggregates every report, and reconciles the bank lines — all in
VBA, with **no database, no add-in, and no external provider** required. Nothing
to install beyond Excel itself.

(An optional SQL implementation of the reporting layer is included in
`modReports` behind `USE_SQL = False`; flip it to `True` only if you specifically
want the reports computed via the `Microsoft.ACE.OLEDB` provider. It is off by
default so the workbook stays self-contained.)

## Verifying it matches the Rust engine
With the shipped demo, entity **BVI** must show: realized **7497.5**, unrealized
**31496.5**, fees **140**, net income **38854**, BalanceSheet **balanced = TRUE**,
and the global imbalance **0**. These are the exact numbers the Rust `ledger`
binary produces for the same scenario.
"#
    .to_string()
}
