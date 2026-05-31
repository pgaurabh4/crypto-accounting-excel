# crypto-accounting-excel

Generates **`GravityLedger.xlsx`** — a single-file crypto double-entry accounting
workbook driven by **SQL + VBA macros** — plus a one-double-click builder that
turns it into the macro-enabled `GravityLedger.xlsm`.

It is a faithful Excel port of the Rust ledger engine in
[`pgaurabh4/crypto-accounting`](https://github.com/pgaurabh4/crypto-accounting)
(the Gravity Ledger platform): double-entry general ledger, FIFO cost basis,
inventory transfers between exchanges, period-end mark-to-market, loans, and
multi-entity reporting under IFRS, for a spot-only crypto market maker.

## What you get (in `dist/`)

| File | Purpose |
|------|---------|
| `GravityLedger.xlsx` | The workbook. Opens anywhere; pre-filled with the demo scenario. |
| `Add-Macros.vbs` | Double-click on Windows/Excel → builds `GravityLedger.xlsm` with the engine embedded, the **RUN ALL** button wired, and reports populated. Excel itself writes the macro container, so it is always valid. |
| `vba/*.bas` | The VBA engine as text, for manual `Alt+F11 → Import`. |
| `SETUP.md` | Three ways to reach a working macro workbook. |

## You edit only four sheets

- **OpeningBalances** — as-at cutover snapshot (cash / crypto / loans); balances against `EQ:OPENING` and seeds FIFO lots.
- **Trades** — `buy`/`sell`; FIFO cost basis + realized P&L.
- **Transfers** — move a coin between exchanges at cost (no P&L).
- **BankStatement** — cash in/out; `counter_account` routes it (`EQ:CAPITAL`, `LIAB:LOAN`, `INC:*`, `EXP:*`, …).

Everything else is **derived on RUN ALL**: the macros post journals + FIFO lots,
then **SQL (ADODB / ACE OLEDB) `GROUP BY`/`JOIN`** builds `TrialBalance`,
`Balances`, `ExchangeBalances` (balances by coin per exchange), `Positions`,
`Loans`, `PnL`, `IncomeStatement`, `BalanceSheet`, and `Reconciliation`. If the
ACE provider is missing the engine falls back to equivalent in-VBA aggregation.

## Build

```sh
cargo run --release --bin gen-xlsm   # writes dist/
cargo run --release --bin simulate   # ports the VBA algorithm to Rust and
                                      # asserts it matches the Rust ledger
```

`simulate` reproduces the reference numbers exactly: BVI realized **7497.5**,
unrealized **31496.5**, fees **140**, net **38854**, balance sheet balanced,
global imbalance **0**.

## Layout

- `src/main.rs` — sheet/data definitions, the `.vbs` builder, packaging.
- `src/xlsx.rs` — minimal OOXML writer (inline strings, styles).
- `src/vba.rs` — the VBA engine source (`modMain`, `modEngine`, `modReports`).
- `src/bin/simulate.rs` — Rust port of the engine logic for validation.

MIT.
