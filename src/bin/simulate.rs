//! Faithful port of the VBA engine's ALGORITHM (modEngine + modReports), run
//! on the shipped demo, to prove the workbook reproduces the Rust `ledger`
//! ground truth. This validates the accounting logic independently of VBA
//! language/runtime correctness (which only Excel can check).
//!
//! Uses f64 (the demo values are exact in f64, matching the workbook's Double
//! reporting path). Run: `cargo run --bin simulate`.

use std::collections::HashMap;

#[derive(Clone)]
struct Acct {
    entity: String,
    code: String,
    ty: String,
}
#[derive(Clone)]
struct Lot {
    id: i64,
    entity: String,
    asset: String,
    exch: String,
    acct: i64,
    acq: String,
    qty: f64,
    unit: f64,
}

struct Eng {
    acc_by_id: HashMap<i64, Acct>,
    acc_by_key: HashMap<String, i64>,
    aid: i64,
    jid: i64,
    lotid: i64,
    jlines: Vec<(i64, i64, f64)>, // (journal_id, account_id, amount)
    lots: Vec<Lot>,
    fvadj: HashMap<String, f64>,
}

fn acct_type(code: &str) -> &'static str {
    if code.starts_with("CASH:") || code.starts_with("CRYPTO:") {
        "Asset"
    } else if code.starts_with("LIAB:") {
        "Liability"
    } else if code.starts_with("EQ:") {
        "Equity"
    } else if code.starts_with("PNL:") || code.starts_with("INC:") {
        "Income"
    } else if code.starts_with("EXP:") {
        "Expense"
    } else {
        "Equity"
    }
}
fn debit_normal(ty: &str) -> bool {
    ty == "Asset" || ty == "Expense"
}

impl Eng {
    fn new() -> Self {
        Eng {
            acc_by_id: HashMap::new(),
            acc_by_key: HashMap::new(),
            aid: 0,
            jid: 0,
            lotid: 0,
            jlines: Vec::new(),
            lots: Vec::new(),
            fvadj: HashMap::new(),
        }
    }
    fn get_acct(&mut self, entity: &str, code: &str) -> i64 {
        let k = format!("{entity}|{code}");
        if let Some(id) = self.acc_by_key.get(&k) {
            return *id;
        }
        self.aid += 1;
        self.acc_by_id.insert(
            self.aid,
            Acct {
                entity: entity.into(),
                code: code.into(),
                ty: acct_type(code).into(),
            },
        );
        self.acc_by_key.insert(k, self.aid);
        self.aid
    }
    fn post(&mut self, lines: &[(i64, f64)]) {
        let s: f64 = lines.iter().map(|l| l.1).sum();
        assert!(s.abs() < 1e-6, "unbalanced journal sum={s}");
        self.jid += 1;
        for (a, amt) in lines {
            self.jlines.push((self.jid, *a, *amt));
        }
    }
    fn add_lot(
        &mut self,
        entity: &str,
        asset: &str,
        exch: &str,
        acct: i64,
        acq: &str,
        qty: f64,
        unit: f64,
    ) {
        self.lotid += 1;
        self.lots.push(Lot {
            id: self.lotid,
            entity: entity.into(),
            asset: asset.into(),
            exch: exch.into(),
            acct,
            acq: acq.into(),
            qty,
            unit,
        });
    }
    fn consume_fifo(&mut self, entity: &str, acct: i64, qty: f64) -> f64 {
        if qty == 0.0 {
            return 0.0;
        }
        let avail: f64 = self
            .lots
            .iter()
            .filter(|l| l.entity == entity && l.acct == acct)
            .map(|l| l.qty)
            .sum();
        assert!(
            avail + 1e-9 >= qty,
            "insufficient inventory acct {acct}: need {qty} have {avail}"
        );
        // oldest first
        let mut order: Vec<usize> = (0..self.lots.len())
            .filter(|&i| {
                self.lots[i].entity == entity && self.lots[i].acct == acct && self.lots[i].qty > 0.0
            })
            .collect();
        order.sort_by(|&a, &b| {
            self.lots[a]
                .acq
                .cmp(&self.lots[b].acq)
                .then(self.lots[a].id.cmp(&self.lots[b].id))
        });
        let mut remaining = qty;
        let mut cost = 0.0;
        for i in order {
            if remaining <= 0.0 {
                break;
            }
            let lq = self.lots[i].qty;
            let uc = self.lots[i].unit;
            if lq <= remaining + 1e-12 {
                cost += lq * uc;
                remaining -= lq;
                self.lots[i].qty = 0.0;
            } else {
                cost += remaining * uc;
                self.lots[i].qty = lq - remaining;
                remaining = 0.0;
            }
        }
        cost
    }
    fn do_trade(
        &mut self,
        entity: &str,
        side: &str,
        base: &str,
        quote: &str,
        qty: f64,
        price: f64,
        fee: f64,
    ) {
        let crypto = self.get_acct(entity, &format!("CRYPTO:{base}:MAIN"));
        let cash = self.get_acct(entity, &format!("CASH:{quote}"));
        let feea = self.get_acct(entity, "EXP:FEES:TRADING");
        let notional = qty * price;
        if side == "buy" {
            self.post(&[(crypto, notional), (feea, fee), (cash, -(notional + fee))]);
            self.add_lot(entity, base, "MAIN", crypto, "", qty, price);
        } else {
            let reala = self.get_acct(entity, "PNL:REALIZED");
            let cb = self.consume_fifo(entity, crypto, qty);
            let proceeds = notional;
            let realized = proceeds - cb;
            self.post(&[
                (cash, proceeds - fee),
                (feea, fee),
                (crypto, -cb),
                (reala, -realized),
            ]);
        }
    }
    fn do_transfer(
        &mut self,
        entity: &str,
        asset: &str,
        fromx: &str,
        tox: &str,
        qty: f64,
        fee: f64,
        ts: &str,
    ) {
        let froma = self.get_acct(entity, &format!("CRYPTO:{asset}:{fromx}"));
        let toa = self.get_acct(entity, &format!("CRYPTO:{asset}:{tox}"));
        let feenet = self.get_acct(entity, "EXP:FEES:NETWORK");
        let cost_moved = self.consume_fifo(entity, froma, qty);
        let cost_fee = if fee > 0.0 {
            self.consume_fifo(entity, froma, fee)
        } else {
            0.0
        };
        self.post(&[
            (toa, cost_moved),
            (froma, -(cost_moved + cost_fee)),
            (feenet, cost_fee),
        ]);
        if qty > 0.0 {
            self.add_lot(entity, asset, tox, toa, ts, qty, cost_moved / qty);
        }
    }
    fn do_bank(&mut self, entity: &str, ccy: &str, amount: f64, counter: &str) {
        let cash = self.get_acct(entity, &format!("CASH:{ccy}"));
        let counter = if counter.is_empty() {
            "EQ:CAPITAL"
        } else {
            counter
        };
        let ctr = self.get_acct(entity, counter);
        self.post(&[(cash, amount), (ctr, -amount)]);
    }
    fn do_revalue(&mut self, entity: &str, asset: &str, mark: f64) {
        let mut qty = 0.0;
        let mut cost = 0.0;
        for l in self
            .lots
            .iter()
            .filter(|l| l.entity == entity && l.asset == asset)
        {
            qty += l.qty;
            cost += l.qty * l.unit;
        }
        if qty == 0.0 {
            return;
        }
        let target = qty * mark - cost;
        let k = format!("{entity}|{asset}");
        let cur = *self.fvadj.get(&k).unwrap_or(&0.0);
        let delta = target - cur;
        let fva = self.get_acct(entity, "CRYPTO:FVADJ");
        let una = self.get_acct(entity, "PNL:UNREALIZED");
        self.post(&[(fva, delta), (una, -delta)]);
        self.fvadj.insert(k, target);
    }
    fn net_by_code(&self, entity: &str, code: &str) -> f64 {
        let id = match self.acc_by_key.get(&format!("{entity}|{code}")) {
            Some(i) => *i,
            None => return 0.0,
        };
        self.jlines.iter().filter(|l| l.1 == id).map(|l| l.2).sum()
    }
    fn display_by_code(&self, entity: &str, code: &str) -> f64 {
        let net = self.net_by_code(entity, code);
        let ty = acct_type(code);
        if debit_normal(ty) {
            net
        } else {
            -net
        }
    }
}

fn marks(
    entity: &str,
    asset: &str,
    trades: &[(&str, &str, &str, &str, f64, f64, f64, &str)],
) -> f64 {
    // override table (asset, price, entity)
    let over = [("BTC", 72000.0, "BVI"), ("ETH", 3200.0, "BVI")];
    for (a, p, e) in over {
        if a == asset && e == entity {
            return p;
        }
    }
    // last trade price for entity+asset
    let mut last = 0.0;
    for t in trades {
        if t.0 == entity && t.2 == asset {
            last = t.5;
        }
    }
    last
}

fn approx(a: f64, b: f64) -> bool {
    (a - b).abs() <= 0.005 + 1e-9 * (a.abs() + b.abs())
}

fn main() {
    let mut e = Eng::new();

    // ---- demo inputs (identical to GravityLedger.xlsx) -------------------
    let bank: &[(&str, &str, f64, &str, &str)] = &[
        (
            "BVI",
            "USD",
            1_000_000.0,
            "EQ:CAPITAL",
            "2026-01-01T00:00:00Z",
        ),
        (
            "IDX",
            "IDR",
            10_000_000_000.0,
            "EQ:CAPITAL",
            "2026-01-09T00:00:00Z",
        ),
    ];
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
    let transfers: &[(&str, &str, &str, &str, f64, f64, &str)] = &[(
        "BVI",
        "BTC",
        "MAIN",
        "COLD",
        1.5,
        0.0005,
        "2026-01-05T00:00:00Z",
    )];

    // ---- merge by timestamp (opening first; none here) -------------------
    enum Ev<'a> {
        T(&'a (&'a str, &'a str, &'a str, &'a str, f64, f64, f64, &'a str)),
        X(&'a (&'a str, &'a str, &'a str, &'a str, f64, f64, &'a str)),
        B(&'a (&'a str, &'a str, f64, &'a str, &'a str)),
    }
    let mut evs: Vec<(&str, Ev)> = Vec::new();
    for t in trades {
        evs.push((t.7, Ev::T(t)));
    }
    for x in transfers {
        evs.push((x.6, Ev::X(x)));
    }
    for b in bank {
        evs.push((b.4, Ev::B(b)));
    }
    evs.sort_by(|a, b| a.0.cmp(b.0));
    for (_, ev) in &evs {
        match ev {
            Ev::T(t) => e.do_trade(t.0, t.1, t.2, t.3, t.4, t.5, t.6),
            Ev::X(x) => e.do_transfer(x.0, x.1, x.2, x.3, x.4, x.5, x.6),
            Ev::B(b) => e.do_bank(b.0, b.1, b.2, b.3),
        }
    }

    // ---- period-end revalue for every held (entity, asset) ----------------
    let held: Vec<(String, String)> = {
        let mut seen = Vec::new();
        for l in &e.lots {
            if l.qty > 0.0 {
                let k = (l.entity.clone(), l.asset.clone());
                if !seen.contains(&k) {
                    seen.push(k);
                }
            }
        }
        seen
    };
    for (ent, asset) in held {
        let m = marks(&ent, &asset, trades);
        e.do_revalue(&ent, &asset, m);
    }

    // ---- reports for BVI vs ground truth ----------------------------------
    let realized = -e.net_by_code("BVI", "PNL:REALIZED");
    let unrealized = -e.net_by_code("BVI", "PNL:UNREALIZED");
    let fees = e.net_by_code("BVI", "EXP:FEES:TRADING") + e.net_by_code("BVI", "EXP:FEES:NETWORK");
    let net_income = realized + unrealized - fees;

    // balance sheet (BVI)
    let mut assets = 0.0;
    let mut liab = 0.0;
    let mut equity = 0.0;
    let mut income = 0.0;
    let mut expense = 0.0;
    for (id, a) in &e.acc_by_id {
        if a.entity != "BVI" {
            continue;
        }
        let net: f64 = e.jlines.iter().filter(|l| l.1 == *id).map(|l| l.2).sum();
        let disp = if debit_normal(&a.ty) { net } else { -net };
        match a.ty.as_str() {
            "Asset" => assets += disp,
            "Liability" => liab += disp,
            "Equity" => equity += disp,
            "Income" => income += disp,
            "Expense" => expense += disp,
            _ => {}
        }
    }
    let retained = income - expense;
    let bs_balanced = approx(assets, liab + equity + retained);
    let global_imbalance: f64 = e.jlines.iter().map(|l| l.2).sum();

    println!("=== BVI report (simulated VBA engine) ===");
    println!("realized   = {realized}");
    println!("unrealized = {unrealized}");
    println!("fees       = {fees}");
    println!("net income = {net_income}");
    println!(
        "assets={assets} liab={liab} equity={equity} retained={retained} balanced={bs_balanced}"
    );
    println!("global imbalance (all entities) = {global_imbalance}");

    // exchange balances by coin (BVI)
    println!("\n=== ExchangeBalances (BVI) ===");
    let mut agg: HashMap<(String, String, String), (f64, f64)> = HashMap::new();
    for l in e.lots.iter().filter(|l| l.entity == "BVI" && l.qty > 0.0) {
        let ent = agg
            .entry((l.entity.clone(), l.exch.clone(), l.asset.clone()))
            .or_insert((0.0, 0.0));
        ent.0 += l.qty;
        ent.1 += l.qty * l.unit;
    }
    let mut keys: Vec<_> = agg.keys().cloned().collect();
    keys.sort();
    for k in keys {
        let (q, c) = agg[&k];
        println!("{} {} : qty={} cost={}", k.1, k.2, q, c);
    }

    println!("\n=== ASSERTIONS vs Rust `ledger` ground truth ===");
    let checks = [
        ("realized == 7497.5", approx(realized, 7497.5)),
        ("unrealized == 31496.5", approx(unrealized, 31496.5)),
        ("fees == 140", approx(fees, 140.0)),
        ("net income == 38854", approx(net_income, 38854.0)),
        ("balance sheet balanced", bs_balanced),
        ("global imbalance == 0", approx(global_imbalance, 0.0)),
    ];
    let mut all = true;
    for (name, ok) in checks {
        println!("[{}] {}", if ok { "PASS" } else { "FAIL" }, name);
        all &= ok;
    }
    if !all {
        std::process::exit(1);
    }
    println!("\nALL CHECKS PASS — workbook engine matches the Rust ledger.");
}
