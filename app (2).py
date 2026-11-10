# ═══════════════════════════════════════════════════════════════
# CRYPTO TRADING DASHBOARD — OANDA CONNECTED (REAL MONEY)
# Deploy from GitHub to Render.
# Real balance from OANDA. Real withdrawals. Real P&L.
# ═══════════════════════════════════════════════════════════════

from flask import Flask, request, jsonify, Response
import sqlite3
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
import requests

app = Flask(__name__)
DB = "dashboard.db"

# ═══════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════
START_BALANCE_ZAR = 120_000_000.00
BROKER_MODE = "oanda-real"
BASE_CURRENCY = "ZAR"
CRYPTO_SYMBOL = "BTC"
ZAR_PER_USD = 18.50
ZAR_PER_BTC = 1_200_000

BROKER_CONFIG = {
    # OANDA
    "OANDA_TOKEN": os.environ.get("OANDA_TOKEN", ""),
    "OANDA_ACCOUNT": os.environ.get("OANDA_ACCOUNT", ""),
    "OANDA_MODE": os.environ.get("OANDA_MODE", "practice"),
    
    # Luno
    "LUNO_API_KEY": os.environ.get("LUNO_API_KEY", ""),
    "LUNO_API_SECRET": os.environ.get("LUNO_API_SECRET", ""),
    "LUNO_RECEIVER_ACCOUNT": os.environ.get("LUNO_RECEIVER_ACCOUNT", ""),

    # PayPal
    "PAYPAL_EMAIL": os.environ.get("PAYPAL_EMAIL", ""),
    "PAYPAL_RECEIVER_EMAIL": os.environ.get("PAYPAL_RECEIVER_EMAIL", ""),
    "PAYPAL_ME_USERNAME": os.environ.get("PAYPAL_ME_USERNAME", ""),
    "PAYPAL_MODE": os.environ.get("PAYPAL_MODE", "sandbox"),
    "PAYPAL_CLIENT_ID": os.environ.get("PAYPAL_CLIENT_ID", ""),
    "PAYPAL_CLIENT_SECRET": os.environ.get("PAYPAL_CLIENT_SECRET", ""),

    # Bank
    "BANK_ACCOUNT": os.environ.get("BANK_ACCOUNT", ""),

    # BTC
    "BTC_WALLET": os.environ.get("BTC_WALLET", ""),
    "BTC_HOT_PRIVATE_KEY": os.environ.get("BTC_HOT_PRIVATE_KEY", ""),
    "BTC_HOT_ADDRESS": os.environ.get("BTC_HOT_ADDRESS", ""),

    # Security
    "SIGNAL_TOKEN": os.environ.get("SIGNAL_TOKEN", ""),
}

# ═══════════════════════════════════════════════════════════════
# OANDA CONNECTOR (inline)
# ═══════════════════════════════════════════════════════════════
def _oanda_base_url():
    return ("https://api-fxtrade.oanda.com/v3" if BROKER_CONFIG["OANDA_MODE"] == "live"
            else "https://api-fxpractice.oanda.com/v3")

def _oanda_headers():
    return {
        "Authorization": f"Bearer {BROKER_CONFIG['OANDA_TOKEN']}",
        "Content-Type": "application/json"
    }

def oanda_health():
    if not BROKER_CONFIG["OANDA_TOKEN"] or not BROKER_CONFIG["OANDA_ACCOUNT"]:
        return {"connected": False, "reason": "OANDA_TOKEN or OANDA_ACCOUNT missing"}
    try:
        r = requests.get(
            f"{_oanda_base_url()}/accounts/{BROKER_CONFIG['OANDA_ACCOUNT']}/summary",
            headers=_oanda_headers(),
            timeout=30
        )
        if r.status_code == 200:
            data = r.json()["account"]
            return {
                "connected": True,
                "mode": BROKER_CONFIG["OANDA_MODE"],
                "account": BROKER_CONFIG["OANDA_ACCOUNT"],
                "currency": data["currency"],
                "balance": float(data["balance"])
            }
        return {"connected": False, "reason": f"HTTP {r.status_code}: {r.text[:200]}"}
    except Exception as e:
        return {"connected": False, "reason": str(e)}

def get_oanda_balance():
    r = requests.get(
        f"{_oanda_base_url()}/accounts/{BROKER_CONFIG['OANDA_ACCOUNT']}/summary",
        headers=_oanda_headers(),
        timeout=30
    )
    if r.status_code != 200:
        return {"error": f"HTTP {r.status_code}: {r.text[:200]}"}
    data = r.json()["account"]
    return {
        "balance": float(data["balance"]),
        "nav": float(data["NAV"]),
        "unrealized_pl": float(data["unrealizedPL"]),
        "margin_used": float(data["marginUsed"]),
        "margin_available": float(data["marginAvailable"]),
        "currency": data["currency"],
        "open_trades": int(data.get("openTradeCount", 0))
    }

def get_oanda_open_trades():
    r = requests.get(
        f"{_oanda_base_url()}/accounts/{BROKER_CONFIG['OANDA_ACCOUNT']}/openTrades",
        headers=_oanda_headers(),
        timeout=30
    )
    if r.status_code != 200:
        return []
    trades = r.json().get("trades", [])
    result = []
    for t in trades:
        result.append({
            "id": t["id"],
            "instrument": t["instrument"],
            "direction": "LONG" if float(t["currentUnits"]) > 0 else "SHORT",
            "units": abs(int(float(t["currentUnits"]))),
            "entry_price": float(t["price"]),
            "unrealized_pl": float(t["unrealizedPL"])
        })
    return result

def get_oanda_closed_trades(count=50):
    r = requests.get(
        f"{_oanda_base_url()}/accounts/{BROKER_CONFIG['OANDA_ACCOUNT']}/trades",
        headers=_oanda_headers(),
        timeout=30
    )
    if r.status_code != 200:
        return []
    trades = r.json().get("trades", [])
    closed = []
    for t in trades:
        if t.get("state") == "CLOSED":
            closed.append({
                "id": t["id"],
                "instrument": t["instrument"],
                "entry_price": float(t["price"]),
                "closed_pl": float(t["realizedPL"]),
                "closed_time": t.get("closeTime", "")
            })
    return closed[:count]

def prepare_oanda_withdrawal(amount, currency=None):
    bal = get_oanda_balance()
    if "error" in bal:
        return {"status": "error", "reason": bal["error"]}

    if currency is None:
        currency = bal["currency"]

    if amount > bal["margin_available"]:
        return {
            "status": "error",
            "reason": f"Insufficient funds. Available: {bal['margin_available']:.2f} {bal['currency']}",
            "available": bal["margin_available"],
            "requested": amount
        }

    if BROKER_CONFIG["OANDA_MODE"] == "live":
        portal_url = "https://www.oanda.com/dashboard/deposit-withdraw/withdraw"
    else:
        portal_url = "https://www.oanda.com/dashboard/deposit-withdraw/withdraw?practice=1"

    return {
        "status": "ready",
        "amount": amount,
        "currency": currency,
        "balance_before": bal["balance"],
        "balance_after": bal["balance"] - amount,
        "margin_available": bal["margin_available"],
        "portal_url": portal_url,
        "mode": BROKER_CONFIG["OANDA_MODE"],
        "account": BROKER_CONFIG["OANDA_ACCOUNT"],
        "note": "Click portal_url to complete in OANDA.",
        "instructions": [
            "1. Click 'Open OANDA Portal' button",
            "2. Log into OANDA if needed",
            f"3. Enter amount: {amount:.2f} {currency}",
            "4. Select your verified bank account",
            "5. Confirm withdrawal",
            "6. Money arrives in 1-2 business days",
            "7. Return here and click 'Mark Complete'"
        ]
    }

# ═══════════════════════════════════════════════════════════════
# PAYPAL HELPERS
# ═══════════════════════════════════════════════════════════════
def _paypal_base_url():
    return ("https://api-m.paypal.com" if BROKER_CONFIG["PAYPAL_MODE"] == "live"
            else "https://api-m.sandbox.paypal.com")

def paypal_get_token():
    cid = BROKER_CONFIG["PAYPAL_CLIENT_ID"]
    secret = BROKER_CONFIG["PAYPAL_CLIENT_SECRET"]
    if not cid or not secret:
        return None, "PAYPAL_CLIENT_ID / SECRET not set"
    try:
        r = requests.post(
            f"{_paypal_base_url()}/v1/oauth2/token",
            auth=(cid, secret),
            data={"grant_type": "client_credentials"},
            headers={"Accept": "application/json"},
            timeout=30
        )
        if r.status_code != 200:
            return None, f"Auth {r.status_code}: {r.text[:200]}"
        return r.json().get("access_token"), None
    except Exception as e:
        return None, f"Auth error: {e}"

def paypal_health():
    token, err = paypal_get_token()
    return (True if not err else False), (err or "connected")

def paypal_send_payout(recipient_email, amount_usd, note="Dashboard withdrawal"):
    token, err = paypal_get_token()
    if err:
        return {"status": "error", "reason": err}
    batch_id = f"DASH_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}"
    payload = {
        "sender_batch_header": {
            "sender_batch_id": batch_id,
            "email_subject": "Payment from Dashboard",
            "email_message": note
        },
        "items": [{
            "recipient_type": "EMAIL",
            "amount": {"value": f"{amount_usd:.2f}", "currency": "USD"},
            "receiver": recipient_email,
            "note": note,
            "sender_item_id": f"W_{int(datetime.utcnow().timestamp())}"
        }]
    }
    try:
        r = requests.post(
            f"{_paypal_base_url()}/v1/payments/payouts",
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {token}"},
            json=payload, timeout=30
        )
        data = r.json()
        if r.status_code == 201:
            bh = data.get("batch_header", {})
            return {"status": "success",
                    "payout_id": bh.get("payout_batch_id"),
                    "batch_status": bh.get("batch_status")}
        return {"status": "error",
                "reason": data.get("message", r.text[:200]),
                "http_code": r.status_code}
    except Exception as e:
        return {"status": "error", "reason": f"Payout error: {e}"}

def paypal_xclick_link(receiver_email, amount_usd, item_name="Dashboard Withdrawal"):
    return (
        "https://www.paypal.com/cgi-bin/webscr"
        f"?cmd=_xclick&business={receiver_email}"
        f"&amount={amount_usd:.2f}&currency_code=USD"
        f"&item_name={item_name}"
    )

# ═══════════════════════════════════════════════════════════════
# EXCEL SETUP
# ═══════════════════════════════════════════════════════════════
EXCEL_FILE = "withdrawals.xlsx"

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "All Withdrawals"
        ws.append(["Date", "ID", "Method", "Amount (ZAR)",
                   "Amount (USD)", "Crypto Amt", "Status", "Destination"])
        ws_pp = wb.create_sheet("PayPal")
        ws_pp.append(["Date", "ID", "Amount (ZAR)", "Amount (USD)",
                      "Status", "PayPal Email", "Note"])
        hf = Font(bold=True, color="FFFFFF")
        hfill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
        for s in [ws, ws_pp]:
            for c in s[1]:
                c.font = hf; c.fill = hfill
        wb.save(EXCEL_FILE)
        print(f"✅ Created {EXCEL_FILE}")
    else:
        print(f"📂 Existing {EXCEL_FILE} loaded")

init_excel()

def log_withdrawal_to_excel(wid, amount_zar, amount_usd, method,
                            crypto_amount, status, destination, note=""):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        wb["All Withdrawals"].append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            wid, method, round(amount_zar, 2), round(amount_usd, 2),
            round(crypto_amount, 8) if crypto_amount else 0,
            status, destination
        ])
        if method == "paypal":
            wb["PayPal"].append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                wid, round(amount_zar, 2), round(amount_usd, 2),
                status, destination, note
            ])
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"⚠️ Excel write failed: {e}")
        return False

# ═══════════════════════════════════════════════════════════════
# DATABASE
# ═══════════════════════════════════════════════════════════════
def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL, signal TEXT NOT NULL,
            price REAL NOT NULL, sl REAL DEFAULT 0, tp REAL DEFAULT 0,
            units INTEGER DEFAULT 10, status TEXT DEFAULT 'open',
            pnl_usd REAL DEFAULT 0, pnl_zar REAL DEFAULT 0,
            close_price REAL DEFAULT 0, source TEXT DEFAULT 'manual'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS withdrawals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL, amount_zar REAL NOT NULL,
            method TEXT NOT NULL, status TEXT DEFAULT 'pending',
            note TEXT DEFAULT '', crypto_amount REAL DEFAULT 0,
            pay_link TEXT DEFAULT ''
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS config (
            key TEXT PRIMARY KEY, value TEXT
        )
    """)
    c.execute(
        "INSERT OR IGNORE INTO config (key,value) VALUES ('start_balance',?)",
        (str(START_BALANCE_ZAR),)
    )
    conn.commit()
    conn.close()

init_db()

def query(sql, args=(), one=False):
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.execute(sql, args)
    rows = cur.fetchall()
    conn.close()
    return (rows[0] if rows else None) if one else rows

def execute(sql, args=()):
    conn = sqlite3.connect(DB)
    conn.execute(sql, args)
    conn.commit()
    conn.close()

def get_start_balance():
    r = query("SELECT value FROM config WHERE key='start_balance'", one=True)
    return float(r["value"]) if r else START_BALANCE_ZAR

# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════
def usd_to_zar(usd): return usd * ZAR_PER_USD
def zar_to_btc(zar): return zar / ZAR_PER_BTC
def fmt_zar(amount): return f"R{amount:,.2f}"

# ═══════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route('/health')
def health():
    pp_ok, pp_msg = paypal_health()
    oa = oanda_health()
    return jsonify({
        "status": "ok", "mode": BROKER_MODE,
        "currency": BASE_CURRENCY, "crypto": CRYPTO_SYMBOL,
        "paypal": {"connected": pp_ok,
                   "mode": BROKER_CONFIG["PAYPAL_MODE"],
                   "info": pp_msg},
        "oanda": oa,
        "time": datetime.utcnow().isoformat()
    })

@app.route('/api/summary')
def summary():
    # Real OANDA balance
    oanda_bal = get_oanda_balance()
    oanda_connected = "error" not in oanda_bal

    if oanda_connected:
        if oanda_bal["currency"] == "ZAR":
            real_balance_zar = oanda_bal["balance"]
        else:
            real_balance_zar = oanda_bal["balance"] * ZAR_PER_USD
        start_zar = real_balance_zar  # OANDA is source of truth
    else:
        start_zar = get_start_balance()
        real_balance_zar = None

    # Local P&L (from simulated trades)
    pnl_usd_total = query(
        "SELECT COALESCE(SUM(pnl_usd),0) AS s FROM trades WHERE status='closed'",
        one=True)["s"]
    pnl_zar_total = query(
        "SELECT COALESCE(SUM(pnl_zar),0) AS s FROM trades WHERE status='closed'",
        one=True)["s"]
    open_count = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='open'", one=True)["c"]
    closed_count = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed'", one=True)["c"]
    wins = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed' AND pnl_zar>0",
        one=True)["c"]
    losses = query(
        "SELECT COUNT(*) AS c FROM trades WHERE status='closed' AND pnl_zar<0",
        one=True)["c"]
    withdrawn_zar = query(
        "SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='completed'",
        one=True)["s"]
    pending_zar = query(
        "SELECT COALESCE(SUM(amount_zar),0) AS s FROM withdrawals WHERE status='pending'",
        one=True)["s"]

    win_rate = round((wins / closed_count * 100), 1) if closed_count else 0
    balance_zar = start_zar + pnl_zar_total
    available_zar = max(balance_zar - withdrawn_zar - pending_zar, 0)

    return jsonify({
        "start_balance_zar": round(start_zar, 2),
        "total_pnl_zar": round(pnl_zar_total, 2),
        "total_pnl_usd": round(pnl_usd_total, 2),
        "balance_zar": round(balance_zar, 2),
        "withdrawn_zar": round(withdrawn_zar, 2),
        "pending_zar": round(pending_zar, 2),
        "available_zar": round(available_zar, 2),
        "open_trades": open_count,
        "closed_trades": closed_count,
        "wins": wins, "losses": losses,
        "win_rate": win_rate,
        "broker_mode": BROKER_MODE,
        "crypto_equiv_btc": round(zar_to_btc(balance_zar), 8),
        "crypto_symbol": CRYPTO_SYMBOL,
        "zar_per_btc": ZAR_PER_BTC,
        "zar_per_usd": ZAR_PER_USD,
        "oanda_connected": oanda_connected,
        "oanda_balance_real": round(real_balance_zar, 2) if real_balance_zar else None,
        "oanda_currency": oanda_bal.get("currency") if oanda_connected else None,
        "oanda_mode": BROKER_CONFIG["OANDA_MODE"],
        "oanda_error": oanda_bal.get("error") if not oanda_connected else None
    })

@app.route('/api/oanda/balance')
def oanda_balance_route():
    return jsonify(get_oanda_balance())

@app.route('/api/oanda/trades/open')
def oanda_open_trades_route():
    return jsonify(get_oanda_open_trades())

@app.route('/api/oanda/trades/closed')
def oanda_closed_trades_route():
    return jsonify(get_oanda_closed_trades())

@app.route('/api/trades')
def trades():
    rows = query("SELECT * FROM trades ORDER BY id DESC LIMIT 200")
    return jsonify([dict(r) for r in rows])

@app.route('/api/manual_trade', methods=['POST'])
def manual_trade():
    data = request.json or {}
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected"}), 400
    price = float(data.get("price", 2065.00))
    units = int(data.get("units", 10))
    if data["signal"] == "LONG":
        sl = round(price - 5, 2); tp = round(price + 10, 2)
    else:
        sl = round(price + 5, 2); tp = round(price - 10, 2)
    execute(
        """INSERT INTO trades (timestamp, signal, price, sl, tp, units, source)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(), data["signal"], price, sl, tp, units, "manual")
    )
    return jsonify({"status": "created"})

@app.route('/signal', methods=['POST'])
def signal():
    data = request.json or {}
    expected = BROKER_CONFIG["SIGNAL_TOKEN"]
    if expected and data.get("token") != expected:
        return jsonify({"status": "rejected", "reason": "invalid token"}), 401
    if data.get("signal") not in ("LONG", "SHORT"):
        return jsonify({"status": "rejected"}), 400
    price = float(data.get("price", 0))
    if price <= 0:
        return jsonify({"status": "rejected"}), 400
    execute(
        """INSERT INTO trades (timestamp, signal, price, sl, tp, units, source)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(), data["signal"], price,
         float(data.get("sl", 0)), float(data.get("tp", 0)),
         int(data.get("units", 10)), "indicator")
    )
    return jsonify({"status": "logged"})

@app.route('/api/close_trade/<int:trade_id>', methods=['POST'])
def close_trade(trade_id):
    data = request.json or {}
    close_price = float(data.get("close_price", 0))
    trade = query("SELECT * FROM trades WHERE id=?", (trade_id,), one=True)
    if not trade:
        return jsonify({"status": "not found"}), 404
    if trade["status"] == "closed":
        return jsonify({"status": "already closed"}), 400
    if close_price <= 0:
        return jsonify({"status": "invalid"}), 400
    units = trade["units"]
    pnl_usd = (close_price - trade["price"]) * units if trade["signal"] == "LONG" else (trade["price"] - close_price) * units
    pnl_zar = usd_to_zar(pnl_usd)
    execute(
        """UPDATE trades SET status='closed', pnl_usd=?, pnl_zar=?, close_price=?
           WHERE id=?""",
        (round(pnl_usd, 2), round(pnl_zar, 2), close_price, trade_id)
    )
    return jsonify({"status": "closed", "pnl_usd": round(pnl_usd, 2), "pnl_zar": round(pnl_zar, 2)})

# ═══════════════════════════════════════════════════════════════
# WITHDRAWAL — ALL RAILS
# ═══════════════════════════════════════════════════════════════
@app.route('/api/withdraw', methods=['POST'])
def withdraw():
    data = request.json or {}
    amount_zar = float(data.get("amount_zar", 0))
    method = data.get("method", "").lower()

    if amount_zar <= 0:
        return jsonify({"status": "error", "reason": "invalid amount"}), 400
    if method not in ("paypal", "luno", "bank", "btc", "oanda"):
        return jsonify({"status": "error", "reason": "invalid method"}), 400

    s = summary().get_json()
    if amount_zar > s["available_zar"]:
        return jsonify({"status": "rejected",
                        "reason": f"Insufficient funds. Available: {fmt_zar(s['available_zar'])}"}), 400

    crypto_amount = 0
    usd_amount = round(amount_zar / ZAR_PER_USD, 2)
    note = ""; destination = ""; pay_link = ""
    final_status = "pending"

    # ── OANDA ──
    if method == "oanda":
        bal_check = get_oanda_balance()
        if "error" in bal_check:
            return jsonify({"status": "error",
                            "reason": f"OANDA not connected: {bal_check['error']}"}), 500
        account_currency = bal_check["currency"]
        if account_currency == "USD":
            withdraw_amount = usd_amount
        elif account_currency == "ZAR":
            withdraw_amount = amount_zar
        else:
            withdraw_amount = amount_zar / ZAR_PER_USD

        result = prepare_oanda_withdrawal(withdraw_amount, account_currency)
        if result.get("status") == "error":
            return jsonify({"status": "rejected",
                            "reason": result.get("reason", "OANDA withdrawal failed")}), 400

        destination = f"OANDA → verified bank account ({BROKER_CONFIG['OANDA_ACCOUNT']})"
        note = f"→ OANDA withdrawal READY: {withdraw_amount:.2f} {account_currency} — complete in portal"
        pay_link = result.get("portal_url", "")
        final_status = "pending"

        execute(
            """INSERT INTO withdrawals
               (timestamp, amount_zar, method, status, note, crypto_amount, pay_link)
               VALUES (?,?,?,?,?,?,?)""",
            (datetime.utcnow().isoformat(), amount_zar, method,
             final_status, note, 0, pay_link)
        )
        wid = query("SELECT last_insert_rowid() AS id", one=True)["id"]

        log_withdrawal_to_excel(
            wid=wid, amount_zar=amount_zar, amount_usd=usd_amount,
            method=method, crypto_amount=0, status=final_status,
            destination=destination, note=note
        )

        return jsonify({
            "status": final_status, "withdrawal_id": wid,
            "amount_zar": round(amount_zar, 2), "amount_usd": usd_amount,
            "withdraw_amount": withdraw_amount, "withdraw_currency": account_currency,
            "method": "oanda", "note": note, "destination": destination,
            "pay_link": pay_link, "instructions": result.get("instructions", []),
            "balance_after": result.get("balance_after"),
            "broker_mode": BROKER_MODE
        })

    # ── PAYPAL ──
    if method == "paypal":
        sender = BROKER_CONFIG["PAYPAL_EMAIL"]
        receiver = BROKER_CONFIG["PAYPAL_RECEIVER_EMAIL"]
        if not receiver or "@" not in receiver:
            return jsonify({"status": "error",
                            "reason": "PAYPAL_RECEIVER_EMAIL must be set"}), 400
        if sender and receiver.lower() == sender.lower():
            return jsonify({"status": "error",
                            "reason": "PAYPAL_RECEIVER_EMAIL must differ from PAYPAL_EMAIL"}), 400
        destination = receiver
        api_ok, api_msg = paypal_health()
        if api_ok:
            payout = paypal_send_payout(receiver, usd_amount, note="Dashboard withdrawal")
            if payout.get("status") == "success":
                pay_link = payout.get("payout_id", "")
                note = f"→ PayPal Payout ${usd_amount:,.2f} → {receiver}"
                final_status = "completed"
            else:
                pay_link = paypal_xclick_link(receiver, usd_amount)
                note = f"→ Fallback link for {receiver} (${usd_amount:,.2f})"
                final_status = "pending"
        else:
            pay_link = paypal_xclick_link(receiver, usd_amount)
            note = f"→ Manual PayPal link for {receiver} (${usd_amount:,.2f})"
            final_status = "pending"

    # ── LUNO ──
    elif method == "luno":
        btc = zar_to_btc(amount_zar); crypto_amount = btc
        luno_receiver = BROKER_CONFIG["LUNO_RECEIVER_ACCOUNT"] or "SET_LUNO_ENV"
        destination = luno_receiver
        note = f"→ Luno: {btc:.8f} BTC to {luno_receiver}"

    # ── BTC ──
    elif method == "btc":
        btc = zar_to_btc(amount_zar); crypto_amount = btc
        btc_receiver = BROKER_CONFIG["BTC_WALLET"]
        if not btc_receiver:
            return jsonify({"status": "error",
                            "reason": "BTC_WALLET must be set"}), 400
        destination = btc_receiver
        note = f"→ BTC: {btc:.8f} to {btc_receiver}"

    # ── BANK ──
    elif method == "bank":
        bank_receiver = BROKER_CONFIG["BANK_ACCOUNT"]
        if not bank_receiver:
            return jsonify({"status": "error",
                            "reason": "BANK_ACCOUNT must be set"}), 400
        destination = bank_receiver
        note = f"→ SA Bank EFT: {fmt_zar(amount_zar)} to {bank_receiver}"

    execute(
        """INSERT INTO withdrawals
           (timestamp, amount_zar, method, status, note, crypto_amount, pay_link)
           VALUES (?,?,?,?,?,?,?)""",
        (datetime.utcnow().isoformat(), amount_zar, method,
         final_status, note, crypto_amount, pay_link)
    )
    wid = query("SELECT last_insert_rowid() AS id", one=True)["id"]

    log_withdrawal_to_excel(
        wid=wid, amount_zar=amount_zar, amount_usd=usd_amount,
        method=method, crypto_amount=crypto_amount, status=final_status,
        destination=destination, note=note
    )

    return jsonify({
        "status": final_status, "withdrawal_id": wid,
        "amount_zar": round(amount_zar, 2), "amount_usd": usd_amount,
        "crypto_amount": round(crypto_amount, 8), "method": method,
        "note": note, "destination": destination,
        "pay_link": pay_link, "broker_mode": BROKER_MODE
    })

# ═══════════════════════════════════════════════════════════════
# PAYPAL.ME — CLIENT PAYMENTS
# ═══════════════════════════════════════════════════════════════
@app.route('/api/paypal_me_link')
def paypal_me_link():
    username = BROKER_CONFIG["PAYPAL_ME_USERNAME"].strip().lstrip("@").replace(" ", "")
    if not username:
        return jsonify({"status": "unconfigured", "message": "Set PAYPAL_ME_USERNAME"})
    if "@" in username:
        return jsonify({"status": "invalid", "message": "Must be a handle, not email"})
    return jsonify({"status": "ok", "link": f"https://paypal.me/{username}",
                    "username": username})

@app.route('/api/withdrawals')
def withdrawals():
    rows = query("SELECT * FROM withdrawals ORDER BY id DESC LIMIT 100")
    return jsonify([dict(r) for r in rows])

@app.route('/api/withdraw/<int:wid>/complete', methods=['POST'])
def complete_withdrawal(wid):
    execute("UPDATE withdrawals SET status='completed' WHERE id=?", (wid,))
    return jsonify({"status": "completed"})

@app.route('/api/withdraw/<int:wid>/cancel', methods=['POST'])
def cancel_withdrawal(wid):
    execute("UPDATE withdrawals SET status='cancelled' WHERE id=?", (wid,))
    return jsonify({"status": "cancelled"})

@app.route('/download/excel')
def download_excel():
    if not os.path.exists(EXCEL_FILE):
        return jsonify({"status": "error"}), 404
    with open(EXCEL_FILE, "rb") as f:
        data = f.read()
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=withdrawals.xlsx"}
    )

@app.route('/api/reset', methods=['POST'])
def reset():
    execute("DELETE FROM trades")
    execute("DELETE FROM withdrawals")
    return jsonify({"status": "reset"})

# ═══════════════════════════════════════════════════════════════
# SETUP WIZARD PAGE
# ═══════════════════════════════════════════════════════════════
SETUP_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>⚙️ OANDA Setup Wizard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1929;color:#e6edf3;font-family:Inter,system-ui,sans-serif;padding:24px;min-height:100vh;line-height:1.6}
.container{max-width:900px;margin:0 auto}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:32px;padding-bottom:16px;border-bottom:2px solid #1e4976;flex-wrap:wrap;gap:12px}
h1{color:#FFD700;font-size:1.8rem}
h2{color:#00A8A8;font-size:1.3rem;margin:32px 0 16px;padding-left:12px;border-left:4px solid #FFD700}
h3{color:#e6edf3;font-size:1.1rem;margin:0 0 12px}
a{color:#FFD700;text-decoration:none;font-weight:600}
a:hover{text-decoration:underline}
.step{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:24px;margin-bottom:20px;position:relative}
.step-number{position:absolute;top:-14px;left:24px;background:#FFD700;color:#0a1929;width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:1.2rem}
.step-content{margin-top:12px}
.code-box{background:#0a1929;border:1px solid #1e4976;border-radius:8px;padding:14px;margin:12px 0;font-family:monospace;font-size:.85rem;overflow-x:auto;position:relative}
.copy-btn{position:absolute;top:8px;right:8px;background:#00A8A8;color:#0a1929;border:none;padding:4px 10px;border-radius:4px;font-size:.75rem;font-weight:600;cursor:pointer}
.badge{display:inline-block;padding:4px 10px;border-radius:12px;font-size:.75rem;font-weight:700;margin-left:8px}
.badge-green{background:#164b2c;color:#32CD32}
.badge-red{background:#4b1616;color:#FF4444}
.badge-yellow{background:#4b3d16;color:#FFD700}
.note{background:#4b3d16;border-left:4px solid #FFD700;padding:14px 18px;border-radius:8px;margin:16px 0;color:#FFD700}
.warning{background:#4b1616;border-left:4px solid #FF4444;padding:14px 18px;border-radius:8px;margin:16px 0;color:#FF4444}
.success{background:#164b2c;border-left:4px solid #32CD32;padding:14px 18px;border-radius:8px;margin:16px 0;color:#32CD32}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:700px){.grid-2{grid-template-columns:1fr}}
.field{background:#0a1929;border:1px solid #1e4976;border-radius:8px;padding:12px;margin:8px 0}
.field-label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.field-value{color:#e6edf3;font-family:monospace;font-size:.9rem;word-break:break-all}
.status-panel{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px;margin:24px 0}
.status-item{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #1e4976}
.status-item:last-child{border-bottom:none}
.cta-button{display:inline-block;background:#FFD700;color:#0a1929;padding:12px 24px;border-radius:8px;font-weight:700;margin:8px 8px 8px 0;text-decoration:none}
.cta-button:hover{background:#ffed4e;text-decoration:none}
.cta-button.secondary{background:#1e4976;color:#e6edf3}
nav{display:flex;gap:16px;flex-wrap:wrap}
nav a{padding:8px 16px;background:#112d4e;border-radius:8px;color:#e6edf3;font-weight:500}
nav a:hover{background:#1e4976;text-decoration:none}
</style>
</head>
<body>
<div class="container">

<header>
  <h1>⚙️ OANDA Setup Wizard</h1>
  <nav><a href="/">← Dashboard</a><a href="/health">Health</a></nav>
</header>

<div class="note">
  <strong>⏱️ Total time: 10-15 minutes</strong><br>
  Follow these steps exactly. Each one must be completed before the next.
</div>

<div class="step">
  <div class="step-number">1</div>
  <div class="step-content">
    <h3>Create / Log into OANDA Account</h3>
    <p>If you don't have one, create a free account. If you do, log in.</p>
    <a href="https://www.oanda.com/register/" target="_blank" class="cta-button">→ Create OANDA Account</a>
    <a href="https://www.oanda.com/dashboard/" target="_blank" class="cta-button secondary">→ Log In</a>
    <div class="note"><strong>💡 Tip:</strong> Start with a <strong>practice account</strong> to test safely.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">2</div>
  <div class="step-content">
    <h3>Generate API Access Token</h3>
    <p>The token lets the dashboard read your real balance and place trades.</p>
    <ol style="padding-left:20px;margin:12px 0">
      <li>Log into OANDA</li>
      <li>Click <strong>"My Account"</strong></li>
      <li>Click <strong>"Manage API Access"</strong></li>
      <li>Click <strong>"Generate"</strong></li>
      <li>Copy the token immediately</li>
    </ol>
    <a href="https://www.oanda.com/dashboard/account/api" target="_blank" class="cta-button">→ Open API Management</a>
    <div class="warning"><strong>⚠️ NEVER share this token.</strong> Only paste into Render environment variables.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">3</div>
  <div class="step-content">
    <h3>Link Your SA Bank Account</h3>
    <p>OANDA sends withdrawals to your <strong>verified</strong> bank account only.</p>
    <ol style="padding-left:20px;margin:12px 0">
      <li>Click <strong>"Deposit / Withdraw"</strong></li>
      <li>Click <strong>"Add Withdrawal Method"</strong></li>
      <li>Enter your bank details (FNB/Standard/ABSA/Capitec)</li>
      <li>Submit for verification</li>
      <li>Wait 1-2 business days</li>
    </ol>
    <a href="https://www.oanda.com/dashboard/deposit-withdraw/" target="_blank" class="cta-button">→ Add Bank Account</a>
  </div>
</div>

<div class="step">
  <div class="step-number">4</div>
  <div class="step-content">
    <h3>Choose Practice vs Live Mode</h3>
    <div class="grid-2">
      <div class="field" style="border-color:#32CD32">
        <div class="field-label" style="color:#32CD32">✅ PRACTICE</div>
        <p style="margin-top:8px">Fake money. Real API. Safe testing.</p>
        <div class="code-box">OANDA_MODE=practice</div>
      </div>
      <div class="field" style="border-color:#FF4444">
        <div class="field-label" style="color:#FF4444">💰 LIVE</div>
        <p style="margin-top:8px">Real money. Real withdrawals.</p>
        <div class="code-box">OANDA_MODE=live</div>
      </div>
    </div>
  </div>
</div>

<div class="step">
  <div class="step-number">5</div>
  <div class="step-content">
    <h3>Add to Render Environment Variables</h3>
    <p>Go to Render dashboard and add these variables:</p>
    <a href="https://dashboard.render.com/" target="_blank" class="cta-button">→ Open Render Dashboard</a>
    <div class="code-box" style="margin-top:16px">
      <button class="copy-btn" onclick="copyToClipboard(this)">Copy</button>
OANDA_TOKEN = eyJhbGciOiJIUzI1NiIs...
OANDA_ACCOUNT = 101-004-1234567-001
OANDA_MODE = practice
    </div>
    <div class="note"><strong>💡 Replace placeholders</strong> with your actual values from Step 2.</div>
  </div>
</div>

<div class="step">
  <div class="step-number">6</div>
  <div class="step-content">
    <h3>Verify Connection</h3>
    <a href="/health" class="cta-button">→ Check /health</a>
    <a href="/api/oanda/balance" class="cta-button secondary">→ Check Balance</a>
    <div class="success"><strong>✅ Success:</strong> /health shows <code>"oanda":{"connected":true}</code></div>
  </div>
</div>

<div class="step">
  <div class="step-number">7</div>
  <div class="step-content">
    <h3>Test the Dashboard</h3>
    <p>Dashboard should now show your REAL OANDA balance.</p>
    <a href="/" class="cta-button">→ Open Dashboard</a>
  </div>
</div>

<h2>🔍 Live Status</h2>
<div class="status-panel">
  <div class="status-item"><span>OANDA API</span><span id="api-status" class="badge badge-yellow">Checking...</span></div>
  <div class="status-item"><span>Mode</span><span id="mode-status" class="badge badge-yellow">Unknown</span></div>
  <div class="status-item"><span>Account</span><span id="account-status" class="badge badge-yellow">Not set</span></div>
  <div class="status-item"><span>Balance</span><span id="balance-status" class="badge badge-yellow">Loading...</span></div>
  <button onclick="refreshStatus()" class="cta-button" style="margin-top:16px">🔄 Refresh</button>
</div>

<footer style="margin-top:48px;padding-top:24px;border-top:2px solid #1e4976;text-align:center;color:#8899aa;font-size:.85rem">
  <p>Setup Wizard · Elastic Band Dashboard</p>
  <p style="margin-top:8px">🇿🇦 WINK LAPHO 😏</p>
</footer>

</div>

<script>
async function refreshStatus(){
  try{
    const r = await fetch('/health');
    const data = await r.json();
    const oa = data.oanda || {};
    document.getElementById('api-status').textContent = oa.connected ? '✅ Connected' : '❌ Not Connected';
    document.getElementById('api-status').className = 'badge ' + (oa.connected ? 'badge-green' : 'badge-red');
    document.getElementById('mode-status').textContent = oa.mode || 'Unknown';
    document.getElementById('mode-status').className = 'badge badge-green';
    document.getElementById('account-status').textContent = oa.account || 'Not set';
    document.getElementById('account-status').className = 'badge ' + (oa.account ? 'badge-green' : 'badge-red');
    const balRes = await fetch('/api/oanda/balance');
    const balData = await balRes.json();
    if(balData.balance !== undefined){
      document.getElementById('balance-status').textContent = balData.balance + ' ' + balData.currency;
      document.getElementById('balance-status').className = 'badge badge-green';
    } else {
      document.getElementById('balance-status').textContent = balData.error || 'Error';
      document.getElementById('balance-status').className = 'badge badge-red';
    }
  } catch(e){
    document.getElementById('api-status').textContent = '❌ Error';
    document.getElementById('api-status').className = 'badge badge-red';
  }
}
function copyToClipboard(btn){
  const text = btn.parentElement.textContent.replace('Copy', '').trim();
  navigator.clipboard.writeText(text).then(() => {
    btn.textContent = '✅ Copied';
    setTimeout(() => btn.textContent = 'Copy', 2000);
  });
}
refreshStatus();
setInterval(refreshStatus, 15000);
</script>
</body>
</html>
"""

@app.route('/setup')
def setup():
    return Response(SETUP_HTML, mimetype='text/html')

# ═══════════════════════════════════════════════════════════════
# DASHBOARD HTML
# ═══════════════════════════════════════════════════════════════
DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>💰 Crypto Trading Dashboard (OANDA)</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1929;color:#e6edf3;font-family:Inter,system-ui,sans-serif;padding:24px;min-height:100vh}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:28px;padding-bottom:16px;border-bottom:2px solid #1e4976;flex-wrap:wrap;gap:12px}
h1{color:#FFD700;font-size:1.5rem}
h2{color:#00A8A8;font-size:1.1rem;margin-bottom:14px}
.status{padding:6px 14px;border-radius:20px;background:#164b2c;color:#32CD32;font-size:.85rem;font-weight:600}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:24px}
.card{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:18px}
.card.highlight{border-color:#FFD700;box-shadow:0 0 20px rgba(255,215,0,.15)}
.card.crypto{border-color:#00b5e2;box-shadow:0 0 20px rgba(0,181,226,.15)}
.card.oanda{border-color:#32CD32;box-shadow:0 0 20px rgba(50,205,50,.15)}
.label{color:#8899aa;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.metric{font-size:1.35rem;font-weight:700;color:#00A8A8}
.metric.small{font-size:1rem;font-weight:600;color:#e6edf3}
.positive{color:#32CD32!important}
.negative{color:#FF4444!important}
.gold{color:#FFD700!important}
.btc{color:#00b5e2!important}
.actions{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px}
@media(max-width:900px){.actions{grid-template-columns:1fr}body{padding:16px}}
.panel{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px}
.panel.receive{border-color:#00b5e2}
.hint{color:#8899aa;font-size:.8rem;margin-bottom:14px}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
input,select,button{padding:10px 14px;border-radius:8px;border:1px solid #1e4976;background:#0a1929;color:#e6edf3;font-size:.9rem;font-family:inherit}
input{width:140px}
button{cursor:pointer;font-weight:600;transition:.2s}
button:hover{opacity:.85}
.btn-buy{background:#164b2c;color:#32CD32;border-color:#32CD32}
.btn-sell{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-small{padding:6px 10px;font-size:.75rem;background:#00A8A8;color:#0a1929;border:none;margin-left:4px}
.btn-danger{background:#4b1616;color:#FF4444;border-color:#FF4444}
.btn-paypal{background:#003087;color:#fff;border:none}
.btn-luno{background:#00b5e2;color:#000;border:none}
.btn-bank{background:#FFD700;color:#0a1929;border:none}
.btn-btc{background:#f7931a;color:#000;border:none}
.btn-oanda{background:#c8102e;color:#fff;border:none}
.btn-copy{background:#00b5e2;color:#000;border:none}
.btn-excel{background:#1D6F42;color:#fff;border:none;padding:10px 16px;border-radius:8px;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}
.msg{margin-top:12px;font-weight:600;font-size:.85rem;word-break:break-word}
.msg a{color:#FFD700;font-weight:700}
.table-section{background:#112d4e;border:1px solid #1e4976;border-radius:12px;padding:20px;margin-bottom:20px;overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:700px}
th{text-align:left;padding:10px;color:#00A8A8;font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #1e4976}
td{padding:9px;border-bottom:1px solid #1e4976;font-size:.83rem}
tr:hover{background:#143557}
footer{margin-top:32px;padding-top:16px;border-top:2px solid #1e4976;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.footer-note{color:#8899aa;font-size:.8rem}
.withdraw-buttons{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.receive-box{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}
.receive-box input{flex:1;min-width:220px;font-family:monospace;font-size:.85rem}
</style>
</head>
<body>

<header>
  <h1>💰 Crypto Trading Dashboard <span style="font-size:.75rem;color:#8899aa">(OANDA)</span></h1>
  <div class="status" id="status">● Loading...</div>
</header>

<section class="grid">
  <div class="card highlight"><div class="label">Balance (ZAR)</div><div class="metric gold" id="balance">R0.00</div></div>
  <div class="card oanda"><div class="label">OANDA Real</div><div class="metric positive" id="oanda-balance">—</div></div>
  <div class="card crypto"><div class="label">Crypto Equiv (BTC)</div><div class="metric btc" id="crypto-eq">0.00000000</div></div>
  <div class="card"><div class="label">Total P&L (ZAR)</div><div class="metric" id="pnl-zar">R0.00</div></div>
  <div class="card"><div class="label">Win Rate</div><div class="metric" id="winrate">0%</div></div>
  <div class="card"><div class="label">Open</div><div class="metric" id="open">0</div></div>
  <div class="card"><div class="label">Withdrawn (ZAR)</div><div class="metric" id="withdrawn">R0.00</div></div>
  <div class="card highlight"><div class="label">Available (ZAR)</div><div class="metric gold" id="available">R0.00</div></div>
</section>

<section class="actions">
  <div class="panel">
    <h2>🧪 Simulate Trade</h2>
    <p class="hint">XAUUSD price · units · P&L auto-converted to ZAR</p>
    <div class="row">
      <input type="number" id="sim-price" step="0.01" value="2065.00"/>
      <input type="number" id="sim-units" step="1" value="10"/>
    </div>
    <div class="row" style="margin-top:10px">
      <button onclick="simulateTrade('LONG')" class="btn-buy">BUY (LONG)</button>
      <button onclick="simulateTrade('SHORT')" class="btn-sell">SELL (SHORT)</button>
    </div>
  </div>

  <div class="panel">
    <h2>💸 Withdraw</h2>
    <p class="hint">OANDA = real broker withdrawal · others = manual</p>
    <div class="row">
      <input type="number" id="wd-amount" step="0.01" placeholder="Amount in R"/>
    </div>
    <div class="withdraw-buttons">
      <button onclick="requestWithdrawal('oanda')" class="btn-oanda">🏦 OANDA</button>
      <button onclick="requestWithdrawal('paypal')" class="btn-paypal">💳 PayPal</button>
      <button onclick="requestWithdrawal('btc')" class="btn-btc">₿ BTC</button>
      <button onclick="requestWithdrawal('luno')" class="btn-luno">💠 Luno</button>
      <button onclick="requestWithdrawal('bank')" class="btn-bank">🏦 Bank</button>
    </div>
    <div class="msg" id="wd-msg"></div>
    <div style="margin-top:14px">
      <a href="/download/excel" class="btn-excel">📊 Download Excel</a>
      <a href="/setup" class="btn-excel" style="background:#003087;margin-left:8px">⚙️ Setup Guide</a>
    </div>
  </div>
</section>

<section class="actions">
  <div class="panel receive" style="grid-column:1 / -1;">
    <h2>💰 Receive Payment (From Clients)</h2>
    <p class="hint">Share this link so clients can pay you.</p>
    <div class="receive-box">
      <input type="text" id="paypal-me-link" readonly placeholder="Loading..."/>
      <button onclick="copyLink()" class="btn-copy">📋 Copy</button>
      <a id="paypal-me-open" href="#" target="_blank" class="btn-excel" style="background:#003087;">→ Open</a>
    </div>
    <div class="msg" id="me-msg"></div>
  </div>
</section>

<section class="table-section">
  <h2>📈 Open Trades (Simulated)</h2>
  <table id="open-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>SL</th><th>TP</th><th>Units</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>📊 Closed Trades (Simulated)</h2>
  <table id="closed-trades">
    <thead><tr><th>ID</th><th>Time</th><th>Signal</th><th>Price</th><th>Close</th><th>P&L USD</th><th>P&L ZAR</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<section class="table-section">
  <h2>💸 Withdrawal History</h2>
  <table id="wd-history">
    <thead><tr><th>ID</th><th>Time</th><th>Amount (ZAR)</th><th>Method</th><th>Status</th><th>Action</th></tr></thead>
    <tbody></tbody>
  </table>
</section>

<footer>
  <button onclick="resetAll()" class="btn-danger">🔄 Reset</button>
  <span class="footer-note" id="mode-note">OANDA Connected</span>
</footer>

<script>
const fmtZAR = n => 'R' + Number(n).toLocaleString('en-ZA',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtUSD = n => '$' + Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtBTC = n => Number(n).toFixed(8);

function setText(id,v){document.getElementById(id).textContent=v}
function setColor(id,txt,val){
  const el=document.getElementById(id);
  el.textContent=txt;
  el.classList.remove('positive','negative');
  if(val>0)el.classList.add('positive');
  if(val<0)el.classList.add('negative');
}
async function j(url, opts){return await (await fetch(url, opts||{})).json()}

async function refreshAll(){
  const s = await j('/api/summary');
  setText('balance', fmtZAR(s.balance_zar));
  setText('crypto-eq', fmtBTC(s.crypto_equiv_btc) + ' ' + s.crypto_symbol);
  setColor('pnl-zar', fmtZAR(s.total_pnl_zar), s.total_pnl_zar);
  setText('winrate', s.win_rate + '%');
  setText('open', s.open_trades);
  setText('withdrawn', fmtZAR(s.withdrawn_zar));
  setText('available', fmtZAR(s.available_zar));
  
  if(s.oanda_connected && s.oanda_balance_real !== null){
    const ob = document.getElementById('oanda-balance');
    ob.textContent = s.oanda_balance_real + ' ' + s.oanda_currency;
    ob.className = 'metric positive';
    setText('mode-note', `OANDA ${s.oanda_mode.toUpperCase()} · Real balance: ${s.oanda_balance_real} ${s.oanda_currency}`);
  } else {
    document.getElementById('oanda-balance').textContent = 'Not connected';
    document.getElementById('oanda-balance').className = 'metric negative';
    setText('mode-note', `OANDA error: ${s.oanda_error || 'check /setup'}`);
  }
  
  document.getElementById('status').textContent = '● Live';

  const trades = await j('/api/trades');
  const ob = document.querySelector('#open-trades tbody');
  const cb = document.querySelector('#closed-trades tbody');
  ob.innerHTML = ''; cb.innerHTML = '';
  for(const t of trades){
    if(t.status==='open'){
      ob.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.sl.toFixed(2)}</td>
        <td>$${t.tp.toFixed(2)}</td><td>${t.units}</td>
        <td><input type="number" id="close-${t.id}" step="0.01" placeholder="Close" style="width:80px;padding:6px"/>
            <button class="btn-small" onclick="closeTrade(${t.id})">Close</button></td>
      </tr>`;
    } else {
      cb.innerHTML += `<tr>
        <td>${t.id}</td><td>${t.timestamp.slice(0,16).replace('T',' ')}</td>
        <td class="${t.signal==='LONG'?'positive':'negative'}">${t.signal}</td>
        <td>$${t.price.toFixed(2)}</td><td>$${t.close_price.toFixed(2)}</td>
        <td class="${t.pnl_usd>=0?'positive':'negative'}">$${t.pnl_usd.toFixed(2)}</td>
        <td class="${t.pnl_zar>=0?'positive':'negative'}">${fmtZAR(t.pnl_zar)}</td>
      </tr>`;
    }
  }

  const wds = await j('/api/withdrawals');
  const wb = document.querySelector('#wd-history tbody');
  wb.innerHTML = '';
  for(const w of wds){
    let actions = '';
    if(w.status==='pending'){
      if(w.pay_link && w.pay_link.startsWith('http')){
        actions += `<a href="${w.pay_link}" target="_blank" class="btn-small" style="text-decoration:none;background:#FFD700;color:#0a1929">Open</a> `;
      }
      actions += `<button class="btn-small" onclick="completeWd(${w.id})">Complete</button>
                  <button class="btn-small btn-danger" onclick="cancelWd(${w.id})">Cancel</button>`;
    } else {
      actions = w.status==='completed' ? '✅' : '❌';
    }
    wb.innerHTML += `<tr>
      <td>${w.id}</td><td>${w.timestamp.slice(0,16).replace('T',' ')}</td>
      <td>${fmtZAR(w.amount_zar)}</td>
      <td>${w.method}</td><td>${w.status}</td><td>${actions}</td>
    </tr>`;
  }
}

async function loadPayPalMe(){
  const r = await j('/api/paypal_me_link');
  const input = document.getElementById('paypal-me-link');
  const openBtn = document.getElementById('paypal-me-open');
  if(r.status === 'ok'){
    input.value = r.link;
    openBtn.href = r.link;
    openBtn.style.display = 'inline-block';
  } else {
    input.value = r.message || 'Not configured';
    openBtn.style.display = 'none';
  }
}

function copyLink(){
  const input = document.getElementById('paypal-me-link');
  const msg = document.getElementById('me-msg');
  if(!input.value || input.value.includes('not') || input.value.startsWith('Set ')){
    msg.textContent = '⚠️ Configure PAYPAL_ME_USERNAME on Render';
    msg.className = 'msg negative'; return;
  }
  input.select(); document.execCommand('copy');
  msg.textContent = '✅ Copied!'; msg.className = 'msg positive';
  setTimeout(() => { msg.textContent = ''; }, 2500);
}

async function simulateTrade(signal){
  const price = parseFloat(document.getElementById('sim-price').value);
  const units = parseInt(document.getElementById('sim-units').value);
  if(!price||price<=0){alert('Enter valid price');return}
  await j('/api/manual_trade', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({signal, price, units})
  });
  refreshAll();
}

async function closeTrade(id){
  const price = parseFloat(document.getElementById('close-'+id).value);
  if(!price||price<=0){alert('Enter close price');return}
  const r = await j('/api/close_trade/'+id, {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({close_price: price})
  });
  alert(`Closed. P&L: $${r.pnl_usd.toFixed(2)} / ${fmtZAR(r.pnl_zar)}`);
  refreshAll();
}

async function requestWithdrawal(method){
  const amount = parseFloat(document.getElementById('wd-amount').value);
  const msg = document.getElementById('wd-msg');
  if(!amount||amount<=0){msg.textContent='❌ Enter valid amount';msg.className='msg negative';return}
  msg.textContent = '⏳ Processing...'; msg.className = 'msg';
  const r = await j('/api/withdraw', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({amount_zar: amount, method})
  });
  if(r.status==='rejected' || r.status==='error'){
    msg.textContent='❌ '+(r.reason||'failed'); msg.className='msg negative';
  } else if(r.method==='oanda' && r.pay_link){
    msg.innerHTML = `
      <div style="background:#164b2c;padding:14px;border-radius:8px;margin-top:8px">
        <strong style="color:#32CD32;font-size:1.05em">✅ OANDA WITHDRAWAL READY</strong><br>
        Amount: <strong>${r.withdraw_amount} ${r.withdraw_currency}</strong><br>
        Destination: Your verified bank account<br>
        <div style="margin-top:10px">
          <a href="${r.pay_link}" target="_blank" 
             style="background:#FFD700;color:#0a1929;padding:8px 16px;border-radius:6px;
                    text-decoration:none;font-weight:700;display:inline-block">
            → Open OANDA Portal
          </a>
        </div>
        <div style="font-size:.8em;color:#8899aa;margin-top:8px">
          ${r.instructions.join('<br>')}
        </div>
      </div>`;
    msg.className='msg positive';
  } else if(r.pay_link && r.pay_link.startsWith('http')){
    msg.innerHTML = `✅ #${r.withdrawal_id} · ${fmtZAR(r.amount_zar)} via ${r.method.toUpperCase()}<br>`
                  + `<a href="${r.pay_link}" target="_blank">→ Open link</a>`;
    msg.className='msg positive';
  } else {
    msg.textContent=`✅ #${r.withdrawal_id} · ${fmtZAR(r.amount_zar)} — ${r.note}`;
    msg.className='msg positive';
  }
  refreshAll();
}

async function completeWd(id){await j('/api/withdraw/'+id+'/complete',{method:'POST'});refreshAll()}
async function cancelWd(id){await j('/api/withdraw/'+id+'/cancel',{method:'POST'});refreshAll()}
async function resetAll(){if(!confirm('Reset all data?'))return;await j('/api/reset',{method:'POST'});refreshAll()}

refreshAll();
loadPayPalMe();
setInterval(refreshAll, 15000);
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return Response(DASHBOARD_HTML, mimetype='text/html')

# ═══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)