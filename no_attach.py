#!/usr/bin/env python3
"""
Generate a mock SETTLEMENT-AFFIRMATION email dataset — new scenarios S13-S21 only.
Based on 8 real attached emails (ANZ, TD, RBC, Sinopac, NatWest, BNS, Santander, CSOP, FHLB).
All counterparty names masked; domains contain no CP abbreviations.
RBC (S15) and Sinopac (S16) emit only the first/original email instance (no reply-chain body).
No noise or irrelevant emails generated.
"""

import io
import os
import random
import shutil
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageDraw, ImageFont

random.seed(20260711)

ROOT = os.path.dirname(os.path.abspath(__file__))
F1 = os.path.join(ROOT, "folder1_inbox")
F2 = os.path.join(ROOT, "folder2_fresh")
for d in (F1, F2):
    if os.path.isdir(d):
        shutil.rmtree(d)
    os.makedirs(d)

# ------------------------------------------------------------------ products & systems
PRODUCTS = ["FX Spot", "FX Forward", "FX Swap", "FX Option", "NDF", "Interest Rate Swap",
            "Cross-Currency Swap", "Equity Swap", "Equity-Linked Note", "Equity Option",
            "Accumulator", "Credit Default Swap", "Commodity Swap"]
PRODUCT_SYSTEM = {
    "Equity Swap": "System 1", "Equity-Linked Note": "System 1", "Equity Option": "System 1",
    "Accumulator": "System 1",
    "Interest Rate Swap": "System 2", "Cross-Currency Swap": "System 2",
    "FX Spot": "System 3", "FX Forward": "System 3", "FX Swap": "System 3",
    "FX Option": "System 3", "NDF": "System 3",
    "Credit Default Swap": "System 4", "Commodity Swap": "System 5",
}
CASHFLOW_TYPE = {
    "FX Spot": "Principal", "FX Forward": "Net Settlement", "FX Swap": "Net Settlement",
    "FX Option": "Premium", "NDF": "NDF Close-out", "Interest Rate Swap": "Coupon",
    "Cross-Currency Swap": "Coupon", "Equity Swap": "Performance",
    "Equity-Linked Note": "Coupon", "Equity Option": "Premium",
    "Accumulator": "Settlement", "Credit Default Swap": "Premium",
    "Commodity Swap": "Settlement",
}
CCYS = ["USD", "EUR", "GBP", "JPY", "CHF", "AUD", "SGD", "HKD", "CNH"]

# ------------------------------------------------------------------ bank (recipient)
BANK_TO = "Settlements <bank@settlements.com>"
BANK_MAILBOX = "bank@settlements.com"
BANK_CC_POOL = ["oversight@settlements.com", "settlement.ops@settlements.com"]
BANK_ENTITIES = ["Bank PLC (London)", "Bank N.A. (New York)", "Bank AG (Frankfurt)",
                 "Bank Securities (Tokyo)", "Bank International (Singapore)"]
BANK_COUNTERPARTS = ["BANKBK/LDN", "BANK FIN INC*TYO", "BANK INTL*LDN", "BANKGB2LXXX", "BANKJPJTXXX"]

# anonymized settlement agents
SSI_BANKS = [
    ("AGENT BANK 1", "AGABUS33"), ("AGENT BANK 2", "AGBBGB2L"),
    ("AGENT BANK 3", "AGCBDEFF"), ("AGENT BANK 4", "AGDBFRPP"),
    ("AGENT BANK 5", "AGEBGB22"), ("AGENT BANK 6", "AGFBSG22"),
    ("AGENT BANK 7", "AGGBUS3N"), ("AGENT BANK 8", "AGHBGB21"),
]
INTERMEDIARIES = [
    ("INTERMEDIARY BANK 1", "INTAUS3N"), ("INTERMEDIARY BANK 2", "INTBCHZ8"),
    ("INTERMEDIARY BANK 3", "INTCFRPP"), ("INTERMEDIARY BANK 4", "INTDJPJT"),
]

# ------------------------------------------------------------------ counterparties (S13-S21 only)
# Domains deliberately contain no CP name or abbreviation
SCEN_NEW = [
    # (masked_name, domain, local_part, entity, scenario_code, products)
    ("Counterparty20", "ratesops.cp20.example",    "settlements",
     "Counterparty20 Bank Ltd",
     "S13", ["Interest Rate Swap", "Cross-Currency Swap", "FX Swap"]),

    ("Counterparty21", "eqswap.cp21.example",      "otc.settlements",
     "Counterparty21 Securities Inc",
     "S14", ["Equity Swap"]),

    ("Counterparty22", "otcderiv.cp22.example",    "otcsettlement",
     "Counterparty22 Capital Markets",
     "S15", ["Cross-Currency Swap", "Interest Rate Swap", "FX Forward"]),

    ("Counterparty23", "eqsettle.cp23.example",    "ops.settlements",
     "Counterparty23 Bank (Hong Kong)",
     "S16", ["Equity Swap", "Equity Option"]),

    ("Counterparty24", "ratessettle.cp24.example", "ird.settlements",
     "Counterparty24 Markets Plc",
     "S17", ["Interest Rate Swap", "Cross-Currency Swap"]),

    ("Counterparty25", "otcderiv.cp25.example",    "otcderivatives",
     "Counterparty25 Financial Products Inc",
     "S18", ["Cross-Currency Swap", "Interest Rate Swap", "FX Swap"]),

    ("Counterparty26", "ccssettle.cp26.example",   "asia.operations",
     "Counterparty26 Hong Kong Branch",
     "S19", ["Cross-Currency Swap"]),

    ("Counterparty27", "fundops.cp27.example",     "fund.operations",
     "Counterparty27 Asset Management Ltd",
     "S20", ["Equity Swap"]),

    ("Counterparty28", "swapops.cp28.example",     "swap.settlements",
     "Counterparty28 Bank",
     "S21", ["Interest Rate Swap"]),
]

SCENARIO_CPS = []
for name, dom, local, entity, scen, prods in SCEN_NEW:
    code = "CP" + name.replace("Counterparty", "")
    SCENARIO_CPS.append({
        "name": name, "code": code,
        "bic": f"CPTY{dom[:2].upper()}2L",
        "domain": dom,
        "email": f"{local}@{dom}",
        "backup": f"settlements.backup@{dom}",
        "entity": entity,
        "kind": "scenario",
        "fmt": scen,
        "scenario": scen,
        "products": prods,
    })

ALL_CPS = SCENARIO_CPS

# ------------------------------------------------------------------ counterparty directory
DIRECTORY_ROWS = []
for cp in SCENARIO_CPS:
    DIRECTORY_ROWS.append((
        cp["email"], cp["name"], cp["entity"],
        "/".join(cp["products"])[:40], "primary", f"scenario {cp['scenario']}"
    ))
    DIRECTORY_ROWS.append((
        cp["backup"], cp["name"], cp["entity"],
        "All OTC (backup)", "backup", f"scenario {cp['scenario']}"
    ))
_seen = set()
DIRECTORY_ROWS = [
    r for r in DIRECTORY_ROWS
    if not (r[0].lower() in _seen or _seen.add(r[0].lower()))
]
DIR = {r[0].lower(): r[1] for r in DIRECTORY_ROWS}

def derive_counterparty(addresses):
    for a in addresses:
        org = DIR.get((a or "").strip().lower())
        if org:
            return org
    return ""

# ------------------------------------------------------------------ formatting helpers
START = datetime(2026, 6, 1, 8, 0, 0)

def rand_dt(span=24):
    return START + timedelta(
        days=random.randint(0, span),
        hours=random.randint(0, 9),
        minutes=random.randint(0, 59),
        seconds=random.randint(0, 59),
    )

def money(n, sign=True):
    return f"{n:,.2f}"

def dnum(n):
    return f"{n:,.2f}"

def fmt_date(d, style="iso"):
    return {
        "iso":  d.isoformat(),
        "dmy":  d.strftime("%d/%m/%Y"),
        "dmon": d.strftime("%d-%b-%Y"),
        "dmony":d.strftime("%d %b %Y"),
        "mdY":  d.strftime("%m/%d/%Y"),
        "bY":   d.strftime("%b %d, %Y"),
    }[style]

def maybe(v, p):
    return v if random.random() < p else ""

# ------------------------------------------------------------------ image helper (kept for completeness)
def _load_font(size, bold=False):
    cands = (["Arial Bold.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf", "Helvetica.ttc"] if bold
             else ["Arial.ttf", "arial.ttf", "DejaVuSans.ttf", "Helvetica.ttc"])
    dirs = ["/System/Library/Fonts/Supplemental/", "/Library/Fonts/",
            "C:/Windows/Fonts/", "/usr/share/fonts/truetype/dejavu/",
            "/usr/share/fonts/truetype/liberation/", ""]
    for dd in dirs:
        for n in cands:
            try:
                return ImageFont.truetype(dd + n, size)
            except Exception:
                continue
    return ImageFont.load_default()

_FONT = _load_font(14)
_FONT_B = _load_font(15, bold=True)

def render_png(title, sections):
    pad, lh, w = 16, 24, 600
    nrows = sum(len(p) + (1 if h else 0) for h, p in sections)
    img = Image.new("RGB", (w, pad * 2 + lh * (nrows + 1)), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, lh + pad // 2], fill=(31, 58, 95))
    d.text((pad, 8), title, fill="white", font=_FONT_B)
    y = pad + lh + 6
    for header, pairs in sections:
        if header:
            d.text((pad, y), header, fill=(31, 58, 95), font=_FONT_B)
            y += lh
        for label, val in pairs:
            d.text((pad, y), str(label), fill=(90, 90, 90), font=_FONT)
            d.text((pad + 230, y), str(val), fill=(10, 10, 10), font=_FONT)
            d.line([pad, y + lh - 6, w - pad, y + lh - 6], fill=(228, 228, 228))
            y += lh
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

# ------------------------------------------------------------------ internal cash-flow store
cashflows = []
_cf_seq = 0
_bk_seq = 0

def _ssi_for(cp, ccy):
    bank = random.choice(SSI_BANKS)
    inter = random.choice(INTERMEDIARIES)
    return {
        "bank_name": bank[0], "bank_bic": bank[1],
        "account_number": str(random.randint(10**9, 10**10 - 1)),
        "beneficiary_name": cp["name"], "beneficiary_bic": cp["bic"],
        "intermediary_bank": f"{inter[0]} ({inter[1]})",
    }

def gen_cashflow(cp, dt, product=None, shows_ssi=True, vdate=None, ccy=None):
    global _cf_seq, _bk_seq
    product = product or random.choice(cp["products"])
    ccy = ccy or random.choice(CCYS)
    amount = round(random.uniform(500, 6_000_000), 2)
    if random.random() < 0.12:
        amount = -amount
    direction = "Receive" if amount < 0 else random.choice(["Pay", "Receive"])
    value_date = vdate or (dt + timedelta(days=random.choice([1, 2, 2, 3]))).date()
    cf_type = CASHFLOW_TYPE[product]
    ssi = _ssi_for(cp, ccy)
    cp_trade_ref = str(random.randint(340_000_000, 355_000_000))
    prod_ref = str(random.randint(90_000_000, 106_000_000))
    bank_trade_ref = f"NB-{random.randint(100000, 999999)}"
    _bk_seq += 1
    booking_ref = f"BK-{_bk_seq:05d}"

    roll = random.random()
    if roll < 0.015:
        outcome = "AlreadySettled"
    elif roll < 0.55:
        outcome = "Agreed"
    elif roll < 0.70:
        outcome = "AmountMismatch"
    elif roll < 0.82:
        outcome = "SSIMissing" if shows_ssi else "Agreed"
    else:
        outcome = "Unmatched"

    cf_id = ""
    email_ssi = dict(ssi) if shows_ssi else None
    if outcome != "Unmatched":
        _cf_seq += 1
        cf_id = f"CF-{_cf_seq:06d}"
        status = "Settled" if outcome == "AlreadySettled" else "Not Settled"
        int_amount = amount
        if outcome == "AmountMismatch":
            int_amount = round(amount + random.choice([-1, 1]) * random.uniform(50, 5000), 2)
        if outcome == "SSIMissing":
            email_ssi = None
        _rr = random.random()
        internal_cp_ref = (
            cp_trade_ref if _rr < 0.5
            else ("" if _rr < 0.7 else str(random.randint(340_000_000, 355_000_000)))
        )
        signed_int_amount = -abs(int_amount) if direction == "Pay" else abs(int_amount)
        cashflows.append({
            "cashflow_id": cf_id, "booking_ref": booking_ref,
            "counterparty_org_name": cp["name"], "counterparty_entity": cp["entity"],
            "product_type": product, "trade_system": PRODUCT_SYSTEM[product],
            "cashflow_type": cf_type, "currency": ccy, "amount": signed_int_amount,
            "direction": direction, "value_date": str(value_date), "status": status,
            "bank_trade_ref": bank_trade_ref, "counterparty_trade_ref": internal_cp_ref,
            "product_ref": prod_ref,
            "ssi_bank_name": ssi["bank_name"], "ssi_bank_bic": ssi["bank_bic"],
            "ssi_account_number": ssi["account_number"],
            "ssi_beneficiary_name": ssi["beneficiary_name"],
            "ssi_beneficiary_bic": ssi["beneficiary_bic"],
            "ssi_intermediary_bank": ssi["intermediary_bank"],
        })
    action = {
        "Agreed": "Affirm to counterparty",
        "AmountMismatch": "Query counterparty",
        "SSIMissing": "Obtain SSI from counterparty",
        "Unmatched": "Escalate to Middle Office",
        "AlreadySettled": "Highlight to analyst",
    }[outcome]
    return {
        "product": product, "cashflow_type": cf_type, "currency": ccy,
        "amount": amount, "direction": direction, "value_date": value_date,
        "cp_trade_ref": cp_trade_ref, "prod_ref": prod_ref, "ssi": email_ssi,
        "outcome": outcome, "matched_cashflow_id": cf_id, "action": action,
        "is_allege": "N" if outcome == "Agreed" else "Y",
    }

def add_orphan_cashflows(n):
    global _cf_seq, _bk_seq
    for _ in range(n):
        cp = random.choice(ALL_CPS)
        product = random.choice(cp["products"])
        ccy = random.choice(CCYS)
        ssi = _ssi_for(cp, ccy)
        _cf_seq += 1
        _bk_seq += 1
        orphan_direction = random.choice(["Pay", "Receive"])
        orphan_magnitude = round(random.uniform(1000, 2_000_000), 2)
        cashflows.append({
            "cashflow_id": f"CF-{_cf_seq:06d}",
            "booking_ref": f"BK-{_bk_seq:05d}",
            "counterparty_org_name": cp["name"],
            "counterparty_entity": cp["entity"],
            "product_type": product,
            "trade_system": PRODUCT_SYSTEM[product],
            "cashflow_type": CASHFLOW_TYPE[product],
            "currency": ccy,
            "amount": -orphan_magnitude if orphan_direction == "Pay" else orphan_magnitude,
            "direction": orphan_direction,
            "value_date": str(rand_dt().date()),
            "status": random.choice(["Not Settled", "Not Settled", "Not Settled", "Settled"]),
            "bank_trade_ref": f"NB-{random.randint(100000, 999999)}",
            "counterparty_trade_ref": str(random.randint(340_000_000, 355_000_000)),
            "product_ref": str(random.randint(90_000_000, 106_000_000)),
            "ssi_bank_name": ssi["bank_name"], "ssi_bank_bic": ssi["bank_bic"],
            "ssi_account_number": ssi["account_number"],
            "ssi_beneficiary_name": ssi["beneficiary_name"],
            "ssi_beneficiary_bic": ssi["beneficiary_bic"],
            "ssi_intermediary_bank": ssi["intermediary_bank"],
        })

# ------------------------------------------------------------------ display helpers
def kv_block(pairs, sep=": "):
    w = max(len(l) for l, _ in pairs)
    return "\n".join(f"{l.ljust(w)}{sep}{v}" for l, v in pairs)

def text_table(headers, rows):
    widths = [
        max(len(str(headers[i])), *[len(str(r[i])) for r in rows])
        for i in range(len(headers))
    ]
    line = lambda vals: " | ".join(str(v).ljust(widths[i]) for i, v in enumerate(vals))
    return "\n".join(
        [line(headers), "-+-".join("-" * w for w in widths)] + [line(r) for r in rows]
    )

def html_table(headers, rows, title=None):
    head = "".join(
        f"<th style='background:#1F3A5F;color:#fff;padding:4px 8px'>{h}</th>"
        for h in headers
    )
    body = "".join(
        "<tr>" + "".join(
            f"<td style='padding:3px 8px;border:1px solid #ccc'>{c}</td>" for c in r
        ) + "</tr>"
        for r in rows
    )
    cap = (
        f"<caption style='text-align:left;font-weight:bold;padding:4px'>{title}</caption>"
        if title else ""
    )
    return (
        f"<table style='border-collapse:collapse;font-family:Arial;font-size:12px'>"
        f"{cap}<thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"
    )

# ================================================================== SCENARIO RENDERERS S13-S21

CAUTION_BANNER = (
    "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n"
    "Do not click any links, open attachments or reply if you do not believe "
    "that the email is legitimate.\n\n"
)

def s13(cfs, cp):
    """
    ANZ-style: C/D direction legend, multi-leg swap table, caution banner,
    escalation block. C = cp receives (bank pays); D = cp pays (bank receives).
    """
    vd = fmt_date(cfs[0]["value_date"], "dmon")
    # C = counterparty receives = bank pays; D = counterparty pays = bank receives
    def cd(direction):
        return "C" if direction == "Receive" else "D"

    headers = ["TRN# (INTERNAL)", "VALUE DATE", "CURRENCY", "CASH FLOW AMOUNT", "PAY/RECEIVE"]
    rows = [
        [c["cp_trade_ref"], vd, c["currency"], dnum(abs(c["amount"])), cd(c["direction"])]
        for c in cfs
    ]
    legend = f"C = {cp['name']} receives / bank payment\nD = {cp['name']} pays / bank receipt\n"
    esc1 = f"escalation1@{cp['domain']}"
    esc2 = f"escalation2@{cp['domain']}"
    esc3 = f"escalation3@{cp['domain']}"

    text = (
        CAUTION_BANNER
        + f"Dear Sir/Madam,\n\n"
        f"This email is being sent on behalf of the Settlement Team in regards to the "
        f"below cashflows for value date {vd}.\n"
        f"Please confirm if you agree with the below settlement amounts and settlement "
        f"instructions to process.\n\n"
        + legend + "\n"
        + text_table(headers, rows)
        + f"\n\nEscalations:\n"
        f"Level 1: {esc1}\nLevel 2: {esc2}\nLevel 3: {esc3}\n\n"
        f"Thanks and Regards,\n{cp['name']} Settlement Operations\n"
    )
    head_html = "".join(
        f"<th style='background:#1F3A5F;color:#fff;padding:4px 8px'>{h}</th>"
        for h in headers
    )
    body_html = "".join(
        "<tr>" + "".join(
            f"<td style='padding:3px 8px;border:1px solid #ccc'>{c}</td>" for c in r
        ) + "</tr>"
        for r in rows
    )
    html = (
        f"<html><body>"
        f"<pre>{CAUTION_BANNER}</pre>"
        f"<p>Dear Sir/Madam,</p>"
        f"<p>This email is being sent on behalf of the Settlement Team in regards to the "
        f"below cashflows for value date {vd}.<br>"
        f"Please confirm if you agree with the below settlement amounts and settlement "
        f"instructions to process.</p>"
        f"<p><b>C = {cp['name']} receives / bank payment</b><br>"
        f"<b>D = {cp['name']} pays / bank receipt</b></p>"
        + html_table(headers, rows)
        + f"<p>Escalations:<br>Level 1: {esc1}<br>Level 2: {esc2}<br>Level 3: {esc3}</p>"
        f"<p>Thanks and Regards,<br>{cp['name']} Settlement Operations</p>"
        f"</body></html>"
    )
    return text, html, []


def s14(cfs, cp):
    """
    TD-style: component breakdown per leg (Interest Payment rows + Equity Performance row),
    src_id column, NOIL_LDN-style ref column, total row. Plain text + HTML.
    """
    vd = fmt_date(cfs[0]["value_date"], "mdY")
    all_rows = []
    totals_by_cf = []
    for c in cfs:
        t = c["amount"]
        int1 = round(abs(t) * random.uniform(0.03, 0.08), 2)
        int2 = round(abs(t) * random.uniform(0.01, 0.04), 2)
        eq   = round(t - int1 - int2, 2)
        src1 = str(random.randint(500_000_000, 600_000_000))
        src2 = str(random.randint(500_000_000, 600_000_000))
        src3 = str(random.randint(500_000_000, 600_000_000))
        oms  = f"EQS_{random.randint(100,999):03d}_{cp['code']}_{random.randint(20260101,20280101)}"
        ref  = c["prod_ref"]
        n_code = "n" + cp["code"][2:]
        entity = f"{n_code}_SPD"
        all_rows += [
            [src1, vd, n_code, c["currency"], f"{int1:>20,.2f}", entity,
             "Equity Swap", "Interest Payment", oms, ref, c["cp_trade_ref"]],
            [src2, vd, n_code, c["currency"], f"{int2:>20,.2f}", entity,
             "Equity Swap", "Interest Payment", oms, ref, c["cp_trade_ref"]],
            [src3, vd, n_code, c["currency"], f"{eq:>20,.2f}", entity,
             "Equity Swap", "Equity Performance", oms, ref, c["cp_trade_ref"]],
        ]
        totals_by_cf.append(t)

    total = sum(totals_by_cf)
    headers = ["src_id", "settle_dt", "cp_short_code", "settle_ccy", "settle_amt",
               "entity_short_code", "allotment", "business_event",
               "instrument_name", "ident", f"{cp['code']} Ref"]
    total_row = ["", "", "", "", f"{total:>20,.2f}", "", "", "", "", "", ""]

    text = (
        CAUTION_BANNER
        + f"Hi team,\n\nWe see the {cp['name']} is due to receive "
        f"{cfs[0]['currency']} {dnum(abs(total))} on Value Date {fmt_date(cfs[0]['value_date'], 'dmon')}. "
        f"In the event of a mismatch, please provide your calculations for review.\n\n"
        + text_table(headers, all_rows + [total_row])
        + f"\n\n{cp['name']} Operations | Settlement\n"
    )
    html = (
        f"<html><body><pre>{CAUTION_BANNER}</pre>"
        f"<p>Hi team,</p>"
        f"<p>We see the {cp['name']} is due to receive "
        f"<b>{cfs[0]['currency']} {dnum(abs(total))}</b> on Value Date "
        f"{fmt_date(cfs[0]['value_date'], 'dmon')}. "
        f"In the event of a mismatch, please provide your calculations for review.</p>"
        + html_table(headers, all_rows + [total_row])
        + f"<p>{cp['name']} Operations | Settlement</p></body></html>"
    )
    return text, html, []


def s15(cfs, cp):
    """
    RBC-style: UPDATED prefix, reply-chain format but FIRST INSTANCE ONLY emitted.
    PAY/RECEIVE table with Transaction ID column + SSI confirmation block at bottom.
    No CP name in body. No reply-chain quoted content.
    """
    vd = fmt_date(cfs[0]["value_date"], "dmy")
    headers = [f"{cp['name']} Direction", "Currency", "Amount", "Transaction ID", "Payment Date"]
    rows = [
        [
            "PAY" if c["direction"] == "Pay" else "RECEIVE",
            c["currency"],
            dnum(abs(c["amount"])),
            c["cp_trade_ref"],
            vd,
        ]
        for c in cfs
    ]
    # SSI block — use first CF's SSI if present
    ssi_block_text = ""
    ssi_block_html = ""
    if cfs[0]["ssi"]:
        s = cfs[0]["ssi"]
        ssi_block_text = (
            f"\nPlease also confirm that the SSI account details below are correct? "
            f"If any of the below information requires correction, please notify us in your reply. "
            f"We will process payments to these accounts unless you explicitly advise otherwise.\n\n"
            f"Your SSI Information -- {cfs[0]['currency']}\n"
            + text_table(
                ["BENEFICIARY_BIC_CODE", "BENEFICIARY_ACCOUNT", "ACCOUNT_WITH_INST_BIC"],
                [[s["beneficiary_bic"], s["account_number"], s["bank_bic"]]]
            ) + "\n"
        )
        ssi_block_html = (
            f"<p>Please also confirm that the SSI account details below are correct? "
            f"If any of the below information requires correction, please notify us in your reply. "
            f"We will process payments to these accounts unless you explicitly advise otherwise.</p>"
            + html_table(
                ["BENEFICIARY_BIC_CODE", "BENEFICIARY_ACCOUNT", "ACCOUNT_WITH_INST_BIC"],
                [[s["beneficiary_bic"], s["account_number"], s["bank_bic"]]],
                title=f"Your SSI Information -- {cfs[0]['currency']}"
            )
        )

    automation_tag = f"OTC AUTOMATION TAG - {random.randint(10**7,10**8-1):08x}-" \
                     f"{random.randint(10**3,10**4-1):04x}-" \
                     f"{random.randint(10**3,10**4-1):04x}-" \
                     f"{random.randint(10**3,10**4-1):04x}-" \
                     f"{random.randint(10**11,10**12-1):012x}"

    text = (
        CAUTION_BANNER
        + "Hi team,\n\n"
        "Could you please kindly review and confirm the below settlements? "
        "Please note the revised table below.\n\n"
        + text_table(headers, rows)
        + ssi_block_text
        + f"\nMany Thanks,\nSettlement Operations\n\nEventID:#\n\n{automation_tag}\n"
    )
    html = (
        f"<html><body><pre>{CAUTION_BANNER}</pre>"
        f"<p>Hi team,</p>"
        f"<p>Could you please kindly review and confirm the below settlements? "
        f"Please note the revised table below.</p>"
        + html_table(headers, rows, title="Payment Information")
        + ssi_block_html
        + f"<p>Many Thanks,<br>Settlement Operations</p>"
        f"<p>EventID:#<br><br>{automation_tag}</p>"
        f"</body></html>"
    )
    return text, html, []


def s16(cfs, cp):
    """
    Sinopac-style: URGENT REMINDER, first instance only (no reply-chain body).
    Chinese sender names fully masked. Table: ref, CCY, amount, direction, VD, cpty-ref.
    Inline SSI pay-to instruction below table.
    No CP name in body.
    """
    vd = fmt_date(cfs[0]["value_date"], "dmon")
    headers = ["Ref No", "CCY", "Amount", f"{cp['code']} Direction", "Value Date", "Counterparty Ref No"]
    rows = [
        [
            c["cp_trade_ref"],
            c["currency"],
            dnum(abs(c["amount"])),
            "Receive",
            vd,
            c["prod_ref"],
        ]
        for c in cfs
    ]
    ssi_line = ""
    if cfs[0]["ssi"]:
        s = cfs[0]["ssi"]
        ssi_line = (
            f"\nPlease pay to {s['bank_bic']} a/c no {s['account_number']} "
            f"in favor of {s['beneficiary_bic']}.\n"
        )

    text = (
        CAUTION_BANNER
        + "Hi Team,\n\n"
        "Please confirm the following settlement and SSI.\n\n"
        + text_table(headers, rows)
        + ssi_line
        + "\nThanks & Best Regards,\nSettlement Operations\nTreasury Operations Center\n"
    )
    # No HTML variant — plain text only (matches original)
    return text, None, []


def s17(cfs, cp):
    """
    NatWest-style: terse 7-column flat row per CF.
    Columns: deal_id | internal_ref | ccy | amount(signed, negative=pay) | direction | swap_ref | vd
    Single-line per CF, minimal prose.
    """
    vd = fmt_date(cfs[0]["value_date"], "dmon")
    rows = []
    for c in cfs:
        rows.append([
            c["cp_trade_ref"],
            c["prod_ref"],
            c["currency"],
            f"({dnum(abs(c['amount']))})" if c["direction"] == "Pay" else dnum(c["amount"]),
            "Pay",
            f"SWOPT{random.randint(1000000, 9999999)}",
            vd,
        ])
    esc1 = f"escalation1@{cp['domain']}"
    esc2 = f"escalation2@{cp['domain']}"
    row_line = " | ".join(str(v) for v in rows[0])
    row_html = "<tr>" + "".join(
        f"<td style='padding:3px 8px;border:1px solid #ccc'>{v}</td>" for v in rows[0]
    ) + "</tr>"
    text = (
        CAUTION_BANNER
        + "Hello Team,\n\nPlease confirm the below cash flows.\n\n"
        + row_line
        + f"\n\nEscalation 1: {esc1}\nEscalation 2: {esc2}\n\n"
        f"Settlement Operations\n{cp['name']} | Rates Settlements\n"
    )
    html = (
        f"<html><body><pre>{CAUTION_BANNER}</pre>"
        f"<p>Hello Team,</p><p>Please confirm the below cash flows.</p>"
        f"<table style='border-collapse:collapse;font-family:Arial;font-size:12px'>"
        f"<tbody>{row_html}</tbody></table>"
        + f"<p>Escalation 1: {esc1}<br>Escalation 2: {esc2}</p>"
        f"<p>Settlement Operations<br>{cp['name']} | Rates Settlements</p>"
        f"</body></html>"
    )
    return text, html, []


def s18(cfs, cp):
    """
    BNS-style: Settlement ID table with Deal ID, Settlement Date, PAY/REC (verbose),
    CCY, Amount (negative for pays), Product, Counter Party Name masked.
    Bullet-point instructions preamble. SSI change notice footer.
    """
    vd = fmt_date(cfs[0]["value_date"], "bY")
    headers = [
        "Settlement ID", "Deal ID", "Settlement Date",
        f"{cp['code']} PAY/REC", "Currency", "Amount", "Product", "Counter Party Name",
    ]
    rows = []
    for c in cfs:
        s_id = f"S-{random.randint(2026,2026)}-{random.randint(900000,1100000)}"
        deal_id = f"C{random.randint(80000,99999)}"
        pays = c["direction"] == "Pay"
        direction_label = f"{cp['code']} Pays" if pays else f"{cp['code']} Receive"
        amt_str = f"-{dnum(abs(c['amount']))}" if pays else dnum(c["amount"])
        rows.append([
            s_id, deal_id, vd, direction_label,
            c["currency"], amt_str, c["product"], "BANK GLOBAL FINANCIAL PRODUCTS INC",
        ])
    text = (
        CAUTION_BANNER
        + "Hello,\n\n"
        "We are requesting confirmation that you are in agreement with the cash flows "
        "in the schedule below, which relate to derivative product transactions.\n"
        "* In the case of a dispute, please include your calculations when responding\n"
        "* Please note that amounts will be gross settled unless net settlements have been specified\n"
        f"* Please provide your agreement by responding to this email.\n\n"
        "Summary of Settlements by Deal:\n\n"
        + text_table(headers, rows)
        + f"\n\nRegards,\nDerivative Products Settlements\n"
        f"*** Settlement banking instructions for USD may be subject to change. "
        f"A copy of the updated SSI will be provided on request ***\n"
    )
    html = (
        f"<html><body><pre>{CAUTION_BANNER}</pre>"
        f"<p>Hello,</p>"
        f"<p>We are requesting confirmation that you are in agreement with the cash flows "
        f"in the schedule below, which relate to derivative product transactions.</p>"
        f"<ul><li>In the case of a dispute, please include your calculations when responding</li>"
        f"<li>Please note that amounts will be gross settled unless net settlements have been specified</li>"
        f"<li>Please provide your agreement by responding to this email.</li></ul>"
        + html_table(headers, rows, title="Summary of Settlements by Deal")
        + f"<p>Regards,<br>Derivative Products Settlements</p>"
        f"<p><b>*** Settlement banking instructions for USD may be subject to change. "
        f"A copy of the updated SSI will be provided on request ***</b></p>"
        f"</body></html>"
    )
    return text, html, []


def s19(cfs, cp):
    """
    Santander-style: CCS table with Registry ID, Sett Date, Entity, CCY, Net Amount, Instrument.
    Negative = counterparty pays; positive = counterparty receives.
    Sign convention note. CAD SSI change notice footer.
    """
    vd = fmt_date(cfs[0]["value_date"], "dmy")
    registry_id = str(random.randint(10_000_000, 99_999_999)) + ".00"
    headers = ["Registry Id", "Sett. Date", "Settle. Entity", "Currency", "Net Amount", "Instrument"]
    rows = []
    for c in cfs:
        signed = -abs(c["amount"]) if c["direction"] == "Pay" else abs(c["amount"])
        c["_gt_reference"] = registry_id
        rows.append([
            registry_id,
            vd,
            f"{cp['code']}U",
            c["currency"],
            dnum(signed),
            c["product"],
        ])
    sign_note = "(Negative amounts: counterparty pays, positive amounts: counterparty receives)"
    new_ssi_bic = random.choice(["TDOMCATTTOR", "CHASUS33", "BOFAUS3N"])
    text = (
        "Confidential\n\n"
        + CAUTION_BANNER
        + "Hi Team,\n\n"
        f"Please help to confirm the below CCS settlement value on {fmt_date(cfs[0]['value_date'], 'dmon')}.\n"
        "Please note that we will not be responsible for any costs arising from a "
        "failed settlement if confirmation is not received.\n\n"
        + text_table(headers, rows)
        + f"\n{sign_note}\n\n"
        f"Important note: Please ensure SSI details are up to date. "
        f"New correspondent BIC: {new_ssi_bic}\n\n"
        f"Thank you.\n\nRegards,\nGlobal Markets Operations\n{cp['name']}\n"
    )
    html = (
        f"<html><body>"
        f"<p><b>Confidential</b></p>"
        f"<pre>{CAUTION_BANNER}</pre>"
        f"<p>Hi Team,</p>"
        f"<p>Please help to confirm the below CCS settlement value on "
        f"{fmt_date(cfs[0]['value_date'], 'dmon')}.<br>"
        f"Please note that we will not be responsible for any costs arising from a "
        f"failed settlement if confirmation is not received.</p>"
        + html_table(headers, rows)
        + f"<p><i>{sign_note}</i></p>"
        f"<p>Important note: Please ensure SSI details are up to date. "
        f"New correspondent BIC: <b>{new_ssi_bic}</b></p>"
        f"<p>Thank you.</p><p>Regards,<br>Global Markets Operations<br>{cp['name']}</p>"
        f"</body></html>"
    )
    return text, html, []


def s20(cfs, cp):
    """
    CSOP-style: fund equity swap table with component columns (EQ, Fin, RealizedPayment,
    OpenComm, CloseComm, RlzPayment(SettCCY), ActualFXRate) + fund name mapping table below.
    Large KRW-scale amounts; FX rate column.
    """
    vd_today = fmt_date(cfs[0]["value_date"], "dmy")
    # Fund codes pool (anonymized)
    FUND_CODES = [
        ("FUND-A-2L", str(random.randint(10000000, 19999999)),
         "FUND A DAILY (2X) LEVERAGED PRODUCT", "FUND SERIES - FUND A DAILY (2X) LEVERAGED PRODUCT"),
        ("FUND-B-2I", str(random.randint(10000000, 19999999)),
         "FUND B DAILY (-2X) INVERSE", "FUND SERIES - FUND B DAILY (-2X) INVERSE PRODUCT"),
        ("FUND-C-2L", str(random.randint(10000000, 19999999)),
         "FUND C DAILY 2X LEV PR", "FUND SERIES - FUND C DAILY (2X) LEVERAGED PRODUCT"),
    ]
    headers = [
        "SettCCY", "RealizedPayDate", "Cpty", "Fund", "OMS_SWAPID",
        "EQ", "Fin", "RealizedPayment", "OpenComm", "CloseComm",
        "RlzPayment(SettCCY)", "ActualFXRate",
    ]
    rows = []
    fund_map_used = []
    for c in cfs:
        fc = random.choice(FUND_CODES)
        fx_rate = round(random.uniform(1300, 1450), 1)
        eq_val   = round(c["amount"] * random.uniform(0.85, 1.10), 2)
        fin_val  = round(c["amount"] * random.uniform(-0.05, -0.01), 2)
        realized = round(eq_val + fin_val, 2)
        open_c   = round(abs(c["amount"]) * random.uniform(0.001, 0.005), 3)
        close_c  = 0.0
        rlz_sett = round(realized / fx_rate, 2)
        oms_id   = f"{fc[0]}-{random.randint(10000,99999)}=U6 SWAP-{cp['code']}"
        c["_gt_reference"] = oms_id
        c["_gt_amount_raw"] = rlz_sett
        c["_gt_direction"] = "Receive" if rlz_sett < 0 else "Pay"
        rows.append([
            c["currency"],
            fmt_date(c["value_date"], "dmy"),
            cp["code"],
            fc[0],
            oms_id,
            dnum(eq_val),
            dnum(fin_val),
            dnum(realized),
            f"{open_c:,.3f}",
            f"{close_c:.3f}",
            dnum(rlz_sett),
            str(fx_rate),
        ])
        if fc not in fund_map_used:
            fund_map_used.append(fc)

    # Fund name mapping table
    map_headers = ["Fund Code", "Long name", "Full name"]
    map_rows = [[fc[0], fc[2], fc[3]] for fc in fund_map_used]

    text = (
        CAUTION_BANNER
        + "Hi team\n\n"
        f"Please confirm below settlement for VD {vd_today}.\n\n"
        + text_table(headers, rows)
        + "\n\nFund full name mapping:\n"
        + text_table(map_headers, map_rows)
        + f"\n\nBest Regards,\nFund Operations\n{cp['name']}\n"
    )
    html = (
        f"<html><body><pre>{CAUTION_BANNER}</pre>"
        f"<p>Hi team</p>"
        f"<p>Please confirm below settlement for VD {vd_today}.</p>"
        + html_table(headers, rows)
        + f"<p><b>Fund full name mapping:</b></p>"
        + html_table(map_headers, map_rows)
        + f"<p>Best Regards,<br>Fund Operations<br>{cp['name']}</p>"
        f"</body></html>"
    )
    return text, html, []


def s21(cfs, cp):
    """
    FHLB-style: minimal 3-column table — Reference | CP Reference | Amount (USD).
    Parentheses = negative (pay). All amounts in single currency (USD).
    Terse preamble: 'Please confirm the payments below for [date]'.
    """
    vd = fmt_date(cfs[0]["value_date"], "mdY")
    headers = ["Reference", f"{cp['code']} Reference", "Amount (USD)"]
    rows = []
    for c in cfs:
        amt_str = (
            f"({dnum(abs(c['amount']))})"
            if c["direction"] == "Pay"
            else dnum(abs(c["amount"]))
        )
        rows.append([c["cp_trade_ref"], c["prod_ref"], amt_str])

    text = (
        CAUTION_BANNER
        + f"Hello,\n\n"
        f"Please confirm the payments below for {vd}, "
        f"all amounts are from the counterparty's perspective and in USD.\n\n"
        + text_table(headers, rows)
        + f"\n\nThank you,\n\nSettlement Operations\n{cp['name']}\n"
    )
    # Plain text only — matches original FHLB style
    return text, None, []


# ================================================================== DISPATCH TABLES

SCENARIO = {
    "S13": s13, "S14": s14, "S15": s15, "S16": s16, "S17": s17,
    "S18": s18, "S19": s19, "S20": s20, "S21": s21,
}

SCEN_CFCOUNT = {
    "S13": (2, 5),   # ANZ multi-leg
    "S14": (1, 3),   # TD component breakdown
    "S15": (5, 10),  # RBC multi-payment
    "S16": (1, 1),   # Sinopac: forced to a single cash flow per email
    "S17": (1, 1),   # NatWest: forced to a single cash flow per email
    "S18": (3, 6),   # BNS settlement ID table
    "S19": (1, 3),   # Santander CCS
    "S20": (1, 4),   # CSOP fund swap
    "S21": (1, 1),   # FHLB: forced to a single cash flow per email
}

# Whether each format prints SSI in the body
SHOWS_SSI = {
    "S13": False,   # ANZ: no SSI in body
    "S14": False,   # TD: no SSI
    "S15": True,    # RBC: SSI block at bottom
    "S16": True,    # Sinopac: inline pay-to SSI
    "S17": False,   # NatWest: no SSI
    "S18": False,   # BNS: SSI change notice only, no per-CF SSI
    "S19": False,   # Santander: SSI note only
    "S20": False,   # CSOP: no SSI
    "S21": False,   # FHLB: no SSI
}

# SSI ground-truth population per format
SSI_GT = {
    "S15": "full",   # RBC prints beneficiary BIC + account + bank BIC
    "S16": "bic_acct_benbic",  # Sinopac prints bank BIC + account + beneficiary BIC ("in favor of" line)
}

SSI_KEYS = [
    "ssi_bank_name", "ssi_bank_bic", "ssi_account_number",
    "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank",
]

def gt_ssi(fmt, ssi):
    out = {k: "" for k in SSI_KEYS}
    if not ssi:
        return out
    keys = {
        "full":     SSI_KEYS,
        "bic":      ["ssi_bank_bic"],
        "acct":     ["ssi_account_number"],
        "bic_acct": ["ssi_bank_bic", "ssi_account_number"],
        "bic_acct_benbic": ["ssi_bank_bic", "ssi_account_number", "ssi_beneficiary_bic"],
        "bank":     ["ssi_bank_name"],
        "none":     [],
    }[SSI_GT.get(fmt, "none")]
    m = {
        "ssi_bank_name":       ssi["bank_name"],
        "ssi_bank_bic":        ssi["bank_bic"],
        "ssi_account_number":  ssi["account_number"],
        "ssi_beneficiary_name":ssi["beneficiary_name"],
        "ssi_beneficiary_bic": ssi["beneficiary_bic"],
        "ssi_intermediary_bank":ssi["intermediary_bank"],
    }
    for k in keys:
        out[k] = m[k]
    return out

# Formats that do NOT print a counterparty reference
REF_HIDDEN = set()   # all new scenarios print a reference

# Formats that print the product name
PRODUCT_SHOWN = {"S13", "S18", "S19", "S20"}


# ================================================================== SUBJECT BUILDER

def make_subject(fmt, cp, cfs):
    vd = fmt_date(cfs[0]["value_date"], "dmon")
    vd_dmy = fmt_date(cfs[0]["value_date"], "dmy")
    if fmt == "S13":
        return f"GCM | Please confirm: Swap Settlement vd {vd}"
    if fmt == "S14":
        ref = cfs[0]["prod_ref"]
        return f"Counterparty vs BANK_LDN - {cfs[0]['currency']}| VD {vd} Equity Swap - ID {ref}"
    if fmt == "S15":
        return f"UPDATED: Settlement pre-confirmation value date: {vd_dmy}"
    if fmt == "S16":
        return f"URGENT REMINDER: BANK / Counterparty - Equities settlement value {vd}"
    if fmt == "S17":
        ref = cfs[0]["cp_trade_ref"]
        return f"Settlement Confirmation//{ref}//VD {vd}"
    if fmt == "S18":
        tag = f"G-{random.randint(2026,2026)}-{random.randint(100000,999999)}"
        return f"Settlement Confirmation +{tag}+ - {fmt_date(cfs[0]['value_date'], 'bY')}"
    if fmt == "S19":
        return f"CCS Settlements - {cp['name']} vs BANK value {vd}"
    if fmt == "S20":
        vd2 = fmt_date(
            cfs[0]["value_date"] + timedelta(days=1), "dmy"
        ) if len(cfs) > 1 else vd_dmy
        return f"Settlement Notice VD {vd_dmy} & {vd2}"
    if fmt == "S21":
        return f"Swap Payments {fmt_date(cfs[0]['value_date'], 'mdY')}"
    return f"Settlement confirmation for value {vd}"


# ================================================================== EMAIL ASSEMBLY

emails, ground_truth = [], []
_idx = 0

def next_id():
    global _idx
    _idx += 1
    return f"EML-{_idx:04d}"

def make_cc(cp):
    cc = []
    if random.random() < 0.5:
        cc.append(cp["backup"])
        if random.random() < 0.2:
            cc.append(random.choice(BANK_CC_POOL))
    return cc

BLANK_CF = (
    "counterparty_reference", "product", "currency", "amount", "direction",
    "value_date", "ssi_bank_name", "ssi_bank_bic", "ssi_account_number",
    "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank",
)

# Scenario-specific ground-truth direction rule, derived from what each email format
# actually displays (see changes_for_no_attach.md). Applied ONLY to email_ground_truth.xlsx --
# it does not change what the rendered .eml itself shows.
DIRECTION_RULE = {
    "S13": "invert",          # C/D legend: shown Receive -> true Pay, shown Pay -> true Receive
    "S14": "always_pay",      # always a Pay from our side
    "S15": "invert",          # shown PAY -> true Receive, shown RECEIVE -> true Pay
    "S16": "always_pay",      # table always displays "Receive" but the truth is Pay
    "S17": "always_receive",  # table always displays "Pay" but the truth is Receive
    "S18": "invert",          # shown "CP25 Pays" -> true Receive, "CP25 Receive" -> true Pay
    "S19": "invert",          # negative Net Amount -> true Receive, positive -> true Pay
    "S21": "invert",          # parenthesized (Pay-looking) amount -> true Receive
}

def true_direction(fmt, raw_direction):
    rule = DIRECTION_RULE.get(fmt)
    if rule == "invert":
        return "Pay" if raw_direction == "Receive" else "Receive"
    if rule == "always_pay":
        return "Pay"
    if rule == "always_receive":
        return "Receive"
    return raw_direction

# Currencies that must be identical across every cash flow within one email
SHARED_CCY_SCENARIOS = {"S14", "S15", "S21"}
FIXED_CCY = {"S21": "USD"}   # S21 is always USD, not just "one random currency"

def emit_email(cp, dt, fmt, is_primary, n_cf, scenario_render=None):
    eid = next_id()
    shows_ssi = SHOWS_SSI.get(fmt, False)
    vd = (dt + timedelta(days=random.choice([1, 2, 2, 3]))).date()
    shared_ccy = FIXED_CCY.get(fmt) or (random.choice(CCYS) if fmt in SHARED_CCY_SCENARIOS else None)
    cfs = [gen_cashflow(cp, dt, shows_ssi=shows_ssi, vdate=vd, ccy=shared_ccy) for _ in range(n_cf)]
    if scenario_render:
        text, html, images = scenario_render(cfs, cp)
    else:
        text, html, images = ("", None, [])
    cc_list = make_cc(cp)
    cc = "; ".join(cc_list)
    derived = derive_counterparty([cp["email"]] + cc_list)
    subj = make_subject(fmt, cp, cfs)
    body_mime = "Plain text + HTML" if html else "Plain text"
    sender_disp = f"{cp['name']} Settlements"
    emails.append({
        "email_id": eid,
        "sender_name": sender_disp,
        "sender_email": cp["email"],
        "sender_domain": cp["domain"],
        "cc": cc,
        "email_type": "SettlementAffirmation",
        "email_format": fmt,
        "body_mime": body_mime,
        "subject": subj,
        "received_date": format_datetime(dt),
        "_dt": dt,
        "_render": (text, html, images),
        "_attach": [],
        "folder2": "N",
    })
    for k, c in enumerate(cfs, 1):
        gs = gt_ssi(fmt, c["ssi"])
        gt_direction = c.get("_gt_direction", true_direction(fmt, c["direction"]))
        gt_reference = c.get("_gt_reference", c["cp_trade_ref"])
        gt_amount_mag = abs(c.get("_gt_amount_raw", c["amount"]))
        gt_amount = -gt_amount_mag if gt_direction == "Pay" else gt_amount_mag
        ground_truth.append({
            "email_id": eid,
            "cashflow_index": f"{eid}#{k}",
            "sender_name": sender_disp,
            "sender_email": cp["email"],
            "sender_domain": cp["domain"],
            "from_email": cp["email"],
            "to_email": BANK_MAILBOX,
            "cc_email": cc,
            "derived_counterparty_org_name": derived,
            "sender_category": "Counterparty",
            "email_type": "SettlementAffirmation",
            "email_format": fmt,
            "format_is_primary": is_primary,
            "body_mime": body_mime,
            "subject": subj,
            "received_date": format_datetime(dt),
            "counterparty_reference": ("" if fmt in REF_HIDDEN else gt_reference),
            "product": (c["product"] if fmt in PRODUCT_SHOWN else ""),
            "currency": c["currency"],
            "amount": gt_amount,
            "direction": gt_direction,
            "value_date": str(c["value_date"]),
            "ssi_bank_name":        gs["ssi_bank_name"],
            "ssi_bank_bic":         gs["ssi_bank_bic"],
            "ssi_account_number":   gs["ssi_account_number"],
            "ssi_beneficiary_name": gs["ssi_beneficiary_name"],
            "ssi_beneficiary_bic":  gs["ssi_beneficiary_bic"],
            "ssi_intermediary_bank":gs["ssi_intermediary_bank"],
            "internal_match_status": c["outcome"],
            "matched_cashflow_id":   c["matched_cashflow_id"],
            "match_basis": "Economic fields" if c["outcome"] != "Unmatched" else "",
            "is_allege": c["is_allege"],
            "action": c["action"],
        })


# ================================================================== GENERATION LOOP
# 10 emails per scenario (S13-S21), no noise, no SSI-only emails

for cp in SCENARIO_CPS:
    scen = cp["scenario"]
    lo, hi = SCEN_CFCOUNT[scen]
    for _ in range(5):
        emit_email(
            cp, rand_dt(), scen, "Y",
            random.randint(lo, hi),
            scenario_render=SCENARIO[scen],
        )

add_orphan_cashflows(20)

# folder2: freshest 40 affirmations
aff  = [e for e in emails if e["email_type"] == "SettlementAffirmation"]
f2ids = {
    e["email_id"]
    for e in sorted(aff, key=lambda e: e["_dt"], reverse=True)[:40]
}
for e in emails:
    e["folder2"] = "Y" if e["email_id"] in f2ids else "N"
f2map = {e["email_id"]: e["folder2"] for e in emails}
for r in ground_truth:
    r["folder2"] = f2map[r["email_id"]]
    r["folder1"] = "Y"


# ================================================================== WRITE .eml

def write_eml(folder, e):
    text, html, images = e["_render"]
    msg = EmailMessage()
    msg["From"]           = f'{e["sender_name"]} <{e["sender_email"]}>'
    msg["To"]             = BANK_TO
    if e.get("cc"):
        msg["Cc"]         = e["cc"]
    msg["Date"]           = e["received_date"]
    msg["Subject"]        = e["subject"]
    msg["Message-ID"]     = make_msgid(domain=e["sender_domain"])
    msg["X-Email-ID"]     = e["email_id"]
    msg["X-Email-Type"]   = e["email_type"]
    msg["X-Email-Format"] = e["email_format"]
    msg.set_content(text, cte="8bit")
    if html is not None:
        msg.add_alternative(html, subtype="html", cte="8bit")
        if images:
            hp = msg.get_payload()[-1]
            for cid, png in images:
                hp.add_related(png, maintype="image", subtype="png", cid=f"<{cid}>")
    for fn, data in e.get("_attach", []):
        msg.add_attachment(data, maintype="application", subtype="pdf", filename=fn)
    with open(os.path.join(folder, f'{e["email_id"]}.eml'), "wb") as fh:
        fh.write(bytes(msg))

for e in emails:
    write_eml(F1, e)
    if e["folder2"] == "Y":
        shutil.copyfile(
            os.path.join(F1, f'{e["email_id"]}.eml'),
            os.path.join(F2, f'{e["email_id"]}.eml'),
        )


# ================================================================== XLSX HELPERS

HFILL = PatternFill("solid", fgColor="1F3A5F")
HFONT = Font(color="FFFFFF", bold=True)
AFILL = PatternFill("solid", fgColor="FCE4E4")
SFILL = PatternFill("solid", fgColor="FFF2CC")


def col_letter(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def style_header(ws, n):
    for c in range(1, n + 1):
        ws.cell(row=1, column=c).fill = HFILL
        ws.cell(row=1, column=c).font = HFONT
    ws.freeze_panes = "A2"


# ================================================================== email_ground_truth.xlsx

wb = Workbook()
ws = wb.active
ws.title = "ground_truth"

GCOLS = [
    "email_id", "cashflow_index", "folder1", "folder2",
    "sender_name", "sender_domain", "from_email", "to_email", "cc_email",
    "derived_counterparty_org_name", "sender_category", "email_type",
    "email_format", "format_is_primary", "body_mime", "subject", "received_date",
    "counterparty_reference", "product", "currency", "amount", "direction",
    "value_date", "ssi_bank_name", "ssi_bank_bic", "ssi_account_number",
    "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank",
    "internal_match_status", "matched_cashflow_id", "match_basis",
    "is_allege", "action",
]

ws.append(GCOLS)
for r in ground_truth:
    ws.append([r.get(c, "") for c in GCOLS])
    if r.get("is_allege") == "Y":
        for c in range(1, len(GCOLS) + 1):
            ws.cell(row=ws.max_row, column=c).fill = AFILL
style_header(ws, len(GCOLS))
for i in range(1, len(GCOLS) + 1):
    ws.column_dimensions[col_letter(i)].width = 16


# ------------------------------------------------------------------ folder2 sheet

wf = wb.create_sheet("folder2")
FCOLS = [
    "email_id", "cashflow_index", "sender_name", "email_type", "email_format",
    "subject", "received_date", "scenario_type", "scenario_outcome",
    "is_allege", "action",
]
wf.append(FCOLS)
for r in sorted(
    [r for r in ground_truth if r["folder2"] == "Y"],
    key=lambda r: (r["email_id"], r.get("cashflow_index", "")),
):
    st = "Cash-flow affirmation"
    so = r["internal_match_status"]
    wf.append([
        r["email_id"], r.get("cashflow_index", ""), r["sender_name"],
        r["email_type"], r["email_format"], r["subject"], r["received_date"],
        st, so, r["is_allege"], r.get("action", ""),
    ])
    if r.get("is_allege") == "Y":
        for c in range(1, len(FCOLS) + 1):
            wf.cell(row=wf.max_row, column=c).fill = AFILL
style_header(wf, len(FCOLS))
for i, w in enumerate([10, 12, 26, 20, 14, 54, 30, 20, 26, 8, 24], 1):
    wf.column_dimensions[col_letter(i)].width = w


# ------------------------------------------------------------------ formats catalog sheet

wc = wb.create_sheet("formats")
wc.append(["format_code", "family", "used_by", "description"])
FORMAT_CATALOG = [
    ("S13", "scenario", "CP20",
     "ANZ-style: C/D direction legend, multi-leg swap table, caution banner, escalation block"),
    ("S14", "scenario", "CP21",
     "TD-style: component breakdown (Interest+Equity rows), src_id col, total row"),
    ("S15", "scenario", "CP22",
     "RBC-style: UPDATED prefix, PAY/RECEIVE table + SSI confirmation block; first instance only"),
    ("S16", "scenario", "CP23",
     "Sinopac-style: URGENT REMINDER, ref/CCY/amount/direction/VD/cpty-ref, inline SSI pay-to; first instance only"),
    ("S17", "scenario", "CP24",
     "NatWest-style: terse 7-col flat row, signed amount (parentheses=pay), swap ref"),
    ("S18", "scenario", "CP25",
     "BNS-style: Settlement ID table, verbose PAY/REC label, negative amounts for pays, bullet preamble"),
    ("S19", "scenario", "CP26",
     "Santander-style: CCS table, Registry ID, signed net amount, sign convention note, SSI change footer"),
    ("S20", "scenario", "CP27",
     "CSOP-style: fund equity swap table with EQ/Fin/RealizedPayment/FXRate cols + fund name mapping"),
    ("S21", "scenario", "CP28",
     "FHLB-style: minimal 3-col (Reference | CP Reference | Amount), parentheses=pay, plain text only"),
]
for r in FORMAT_CATALOG:
    wc.append(list(r))
style_header(wc, 4)
for i, w in enumerate([10, 12, 12, 72], 1):
    wc.column_dimensions[col_letter(i)].width = w


# ------------------------------------------------------------------ summary sheet

sm = wb.create_sheet("summary")
aff_rows = [r for r in ground_truth if r["email_type"] == "SettlementAffirmation"]
oc = Counter(r["internal_match_status"] for r in aff_rows)
CFMAP = {c["cashflow_id"]: c for c in cashflows}
_refm = [r for r in aff_rows if r["internal_match_status"] != "Unmatched"]
_refsame = sum(
    1 for r in _refm
    if r["matched_cashflow_id"] and
       str(r["counterparty_reference"]) ==
       str(CFMAP.get(r["matched_cashflow_id"], {}).get("counterparty_trade_ref", ""))
)

srows = [
    ("Sr No", "Metric", "Value"),
    ("1",   "Total emails",                         len(emails)),
    ("1.1", "Counterparty affirmations",             len(aff)),
    ("1.1.1","scenario S13 (ANZ-style)",             sum(1 for e in aff if e["email_format"] == "S13")),
    ("1.1.2","scenario S14 (TD-style)",              sum(1 for e in aff if e["email_format"] == "S14")),
    ("1.1.3","scenario S15 (RBC-style)",             sum(1 for e in aff if e["email_format"] == "S15")),
    ("1.1.4","scenario S16 (Sinopac-style)",         sum(1 for e in aff if e["email_format"] == "S16")),
    ("1.1.5","scenario S17 (NatWest-style)",         sum(1 for e in aff if e["email_format"] == "S17")),
    ("1.1.6","scenario S18 (BNS-style)",             sum(1 for e in aff if e["email_format"] == "S18")),
    ("1.1.7","scenario S19 (Santander-style)",       sum(1 for e in aff if e["email_format"] == "S19")),
    ("1.1.8","scenario S20 (CSOP-style)",            sum(1 for e in aff if e["email_format"] == "S20")),
    ("1.1.9","scenario S21 (FHLB-style)",            sum(1 for e in aff if e["email_format"] == "S21")),
    ("", "", ""),
    ("2",   "Total cash flows (ground-truth rows)",  len(aff_rows)),
    ("2.1", "Agreed",                                oc["Agreed"]),
    ("2.2", "AmountMismatch",                        oc["AmountMismatch"]),
    ("2.3", "SSIMissing",                            oc["SSIMissing"]),
    ("2.4", "Unmatched",                             oc["Unmatched"]),
    ("2.5", "AlreadySettled",                        oc["AlreadySettled"]),
    ("2.6", "of matched: cpty ref == internal ref",  f"{_refsame} / {len(_refm)}"),
    ("3",   "Allege cash flows (is_allege=Y)",
     sum(1 for r in ground_truth if r["is_allege"] == "Y")),
    ("4",   "Internal cash-flow records",            len(cashflows)),
    ("4.1", "status Settled",
     sum(1 for c in cashflows if c["status"] == "Settled")),
    ("5",   "Total ground-truth rows",               len(ground_truth)),
]

for row in srows:
    depth = row[0].count(".") if row[0] and row[0] != "Sr No" else 0
    sm.append([row[0], ("    " * depth) + str(row[1]), row[2]])
    if row[0] and "." not in row[0] and row[0] != "Sr No":
        for c in (1, 2, 3):
            sm.cell(row=sm.max_row, column=c).font = Font(bold=True)
style_header(sm, 3)
sm.column_dimensions["A"].width = 10
sm.column_dimensions["B"].width = 44
sm.column_dimensions["C"].width = 12

wb.save(os.path.join(ROOT, "email_ground_truth.xlsx"))


# ================================================================== internal_bookings.xlsx

wb2 = Workbook()
ws2 = wb2.active
ws2.title = "cashflows"

CFCOLS = [
    "cashflow_id", "booking_ref", "counterparty_org_name", "counterparty_entity",
    "product_type", "trade_system", "cashflow_type", "currency", "amount",
    "direction", "value_date", "status", "bank_trade_ref", "counterparty_trade_ref",
    "product_ref", "ssi_bank_name", "ssi_bank_bic", "ssi_account_number",
    "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank",
]
ws2.append(CFCOLS)
for c in cashflows:
    ws2.append([c.get(k, "") for k in CFCOLS])
    if c["status"] == "Settled":
        for i in range(1, len(CFCOLS) + 1):
            ws2.cell(row=ws2.max_row, column=i).fill = SFILL
style_header(ws2, len(CFCOLS))
for i in range(1, len(CFCOLS) + 1):
    ws2.column_dimensions[col_letter(i)].width = 16

wp = wb2.create_sheet("product_systems")
wp.append(["product_type", "trade_system"])
for p in PRODUCTS:
    wp.append([p, PRODUCT_SYSTEM[p]])
style_header(wp, 2)
wp.column_dimensions["A"].width = 24
wp.column_dimensions["B"].width = 14

wb2.save(os.path.join(ROOT, "internal_bookings.xlsx"))


# ================================================================== counterparty_directory.xlsx

wb3 = Workbook()
ws3 = wb3.active
ws3.title = "counterparty_directory"

DCOLS = [
    "email_address", "counterparty_org_name", "counterparty_entity",
    "product_scope", "dl_role", "notes",
]
ws3.append(DCOLS)
for r in DIRECTORY_ROWS:
    ws3.append(list(r))
style_header(ws3, len(DCOLS))
for i, w in enumerate([40, 20, 34, 28, 10, 24], 1):
    ws3.column_dimensions[col_letter(i)].width = w

wb3.save(os.path.join(ROOT, "counterparty_directory.xlsx"))


# ================================================================== PRINT SUMMARY

noise_count = sum(1 for e in emails if e["email_type"] == "Noise")
ssi_count   = sum(1 for e in emails if e["email_type"] == "SSIVerification")

print(f"Emails       : {len(emails)}  "
      f"(affirmations {len(aff)} / SSI {ssi_count} / noise {noise_count})")
print(f"  by scenario: " + " | ".join(
    f"{scen}={sum(1 for e in aff if e['email_format'] == scen)}"
    for scen in ["S13","S14","S15","S16","S17","S18","S19","S20","S21"]
))
print(f"Cash-flow rows : {len(aff_rows)}  | outcomes: {dict(oc)}")
print(f"Internal CFs   : {len(cashflows)} "
      f"(settled {sum(1 for c in cashflows if c['status'] == 'Settled')})")
print(f"Ground-truth rows: {len(ground_truth)}")
print(f"folder1_inbox  : {len(emails)} .eml files")
print(f"folder2_fresh  : {sum(1 for e in emails if e['folder2'] == 'Y')} .eml files")
