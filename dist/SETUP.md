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
