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

mod vba;
mod xlsx;

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

    pack_zip("dist/GravityLedger-workbook.zip")?;

    println!("wrote dist/GravityLedger.xlsx ({} sheets)", sheets.len());
    println!("wrote dist/vba/{{modMain,modEngine,modReports}}.bas");
    println!("wrote dist/Add-Macros.vbs");
    println!("wrote dist/SETUP.md");
    println!("wrote dist/GravityLedger-workbook.zip (deliverable bundle)");
    Ok(())
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
    rd.row(vec![t("To rebuild"), t("Click the RUN ALL button (added by Add-Macros.vbs). Macros post the journals + FIFO lots; SQL (ADODB/ACE) builds every report.")]);
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
        ("ChartOfAccounts / Journal / JournalLines / Lots", "macro (procedural posting + FIFO)"),
        ("TrialBalance / Balances", "SQL: SUM(amount) per account, JOIN chart of accounts"),
        ("ExchangeBalances", "SQL: GROUP BY entity, exchange, asset over Lots (your 'balances by coin per exchange')"),
        ("Positions", "SQL: GROUP BY entity, asset over Lots"),
        ("Loans", "SQL: outstanding LIAB:LOAN per entity"),
        ("PnL / IncomeStatement / BalanceSheet", "SQL aggregates + IFRS derivation"),
        ("Reconciliation", "macro: match each BankStatement line to its posted cash journal, then tie out statement vs ledger per entity+ccy"),
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

## How "uses SQL" works
On RUN ALL the macros post the double-entry journals and FIFO lots, then open an
**ADODB** connection (`Microsoft.ACE.OLEDB.12.0`) against the workbook itself and
run `GROUP BY` / `JOIN` SQL to build the reports. If the ACE provider is not
installed, the engine automatically falls back to equivalent in-VBA aggregation
so RUN ALL still completes — the popup tells you which path ran.

## Verifying it matches the Rust engine
With the shipped demo, entity **BVI** must show: realized **7497.5**, unrealized
**31496.5**, fees **140**, net income **38854**, BalanceSheet **balanced = TRUE**,
and the global imbalance **0**. These are the exact numbers the Rust `ledger`
binary produces for the same scenario.
"#
    .to_string()
}
