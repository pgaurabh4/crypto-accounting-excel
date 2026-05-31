# Gravity Ledger — setup

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
