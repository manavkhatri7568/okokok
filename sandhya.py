#!/usr/bin/env python3
"""
Generate a mock SETTLEMENT-AFFIRMATION email dataset for the Allege prototype.
 
A counterparty emails the bank's settlement-operations desk asking it to AGREE/CONFIRM the cash flows
(payments) that are due to settle. Pipeline modelled:
    email -> extract cash flow(s) -> look up in internal system (internal_bookings.xlsx `cashflows`)
       Agreed / AmountMismatch / SSIMissing / Unmatched / AlreadySettled -> action.
 
Outputs (next to this script):
  folder1_inbox/           .eml files (counterparty affirmations + SSI verifications + noise)
  folder2_fresh/           fresh subset
  internal_bookings.xlsx   sheets: cashflows (the bank's cash-flow records) + product_systems (mapping)
  email_ground_truth.xlsx  sheets: ground_truth (1 row/cash flow + answer key) + folder2 + formats + summary
  counterparty_directory.xlsx  exact email -> counterparty org (+ entity + product + dl_role)
 
Design notes live in CLAUDE.md (gitignored). No real firm/bank names; the receiving bank is "the bank".
Matching is on ECONOMIC fields (org derived from sender + amount + currency + value date + direction +
product) — NOT on the counterparty's reference (which may differ from the bank's internal ref).
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
    "Equity Swap": "System 1", "Equity-Linked Note": "System 1", "Equity Option": "System 1", "Accumulator": "System 1",
    "Interest Rate Swap": "System 2", "Cross-Currency Swap": "System 2",
    "FX Spot": "System 3", "FX Forward": "System 3", "FX Swap": "System 3", "FX Option": "System 3", "NDF": "System 3",
    "Credit Default Swap": "System 4", "Commodity Swap": "System 5",
}
CASHFLOW_TYPE = {
    "FX Spot": "Principal", "FX Forward": "Net Settlement", "FX Swap": "Net Settlement", "FX Option": "Premium",
    "NDF": "NDF Close-out", "Interest Rate Swap": "Coupon", "Cross-Currency Swap": "Coupon",
    "Equity Swap": "Performance", "Equity-Linked Note": "Coupon", "Equity Option": "Premium",
    "Accumulator": "Settlement", "Credit Default Swap": "Premium", "Commodity Swap": "Settlement",
}
CCYS = ["USD", "EUR", "GBP", "JPY", "CHF", "AUD", "SGD", "HKD"]
 
# ------------------------------------------------------------------ bank (recipient) — never a real bank name
BANK_TO = "Settlements <bank@settlements.com>"
BANK_MAILBOX = "bank@settlements.com"
BANK_CC_POOL = ["oversight@settlements.com", "settlement.ops@settlements.com"]
BANK_ENTITIES = ["Bank PLC (London)", "Bank N.A. (New York)", "Bank AG (Frankfurt)",
                 "Bank Securities (Tokyo)", "Bank International (Singapore)"]
# how the counterparty names the bank as "counterpart" in their tables (anonymized)
BANK_COUNTERPARTS = ["BANKBK/LDN", "BANK FIN INC*TYO", "BANK INTL*LDN", "BANKGB2LXXX", "BANKJPJTXXX"]
# minimal PDF stub used as the "attached SSI" for S9
PDF_STUB = (b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n"
            b"3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 120]>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF\n")
 
# anonymized settlement agents / correspondent banks (no real names)
SSI_BANKS = [("AGENT BANK 1", "AGABUS33"), ("AGENT BANK 2", "AGBBGB2L"), ("AGENT BANK 3", "AGCBDEFF"),
             ("AGENT BANK 4", "AGDBFRPP"), ("AGENT BANK 5", "AGEBGB22"), ("AGENT BANK 6", "AGFBSG22"),
             ("AGENT BANK 7", "AGGBUS3N"), ("AGENT BANK 8", "AGHBGB21")]
INTERMEDIARIES = [("INTERMEDIARY BANK 1", "INTAUS3N"), ("INTERMEDIARY BANK 2", "INTBCHZ8"),
                  ("INTERMEDIARY BANK 3", "INTCFRPP"), ("INTERMEDIARY BANK 4", "INTDJPJT")]
 
# ------------------------------------------------------------------ counterparties
CP_COUNTRY = ['GB', 'US', 'DE', 'FR', 'JP', 'CH', 'AU', 'CA', 'SG', 'HK']
# coined nonsense tokens (web-verified as not real firms) on the RFC-reserved .example TLD
BRANDS = ["zylith", "braxen", "drenvo", "kesvin", "oskira", "plessa", "wexolt", "jindar", "sabreth", "tulvex"]
GENERIC_FORMATS = ["SemiStructured", "FreeForm", "TabularPlain", "TabularHTML",
                   "MixedProseTable", "ForwardedThread", "ReplyChain", "ImageEmbedded"]
 
# CP1-10: generic counterparties, each a primary structural format, own domain, multi-DL
COUNTERPARTIES = []
for i in range(1, 11):
    dom = f"{BRANDS[i-1]}.cp{i}.example"
    COUNTERPARTIES.append({
        "name": f"Counterparty{i}", "code": f"CP{i}", "bic": f"CPTY{CP_COUNTRY[i-1]}2L", "domain": dom,
        "email": f"otc.settlements@{dom}", "backup": f"settlements.ny@{dom}",
        "entity": f"Counterparty{i} International plc", "kind": "generic",
        "fmt": GENERIC_FORMATS[(i - 1) % 8], "products": PRODUCTS,
    })
 
# CP11-17: scenario counterparties (S1-S8). S3 & S5 are the same org (CP13) with two DLs.
SCEN = [
    ("Counterparty11", "grenoll.cp11.example", "otc.settlements", "Counterparty11 International", "S1", ["FX Option"]),
    ("Counterparty12", "yavista.cp12.example", "settlements", "Counterparty12 Master Fund LP", "S2", ["Equity Swap", "Interest Rate Swap"]),
    ("Counterparty13", "vornil.cp13.example", "otcderiv.settlement", "Counterparty13 Bank (London)", "S3", ["Interest Rate Swap", "Equity Swap", "FX Forward", "FX Option", "Cross-Currency Swap"]),
    ("Counterparty14", "xendar.cp14.example", "otcsettlement", "Counterparty14 Securities", "S4", ["Equity-Linked Note"]),
    ("Counterparty13", "vornil.cp13.example", "ndf.settlements", "Counterparty13 Bank (Singapore)", "S5", ["NDF"]),
    ("Counterparty15", "quorlim.cp15.example", "ops.settlements", "Counterparty15 Global Markets", "S6", ["Accumulator", "Equity Swap"]),
    ("Counterparty16", "zandril.cp16.example", "eq.presettlements", "Counterparty16 Bank", "S7", ["Equity Swap"]),
    ("Counterparty17", "lortez.cp17.example", "trsyops.settlements", "Counterparty17 Bank SG", "S8", ["FX Forward", "FX Swap", "Equity Swap"]),
    # new scenarios S9-S12
    ("Counterparty18", "quenvel.cp18.example", "settlement.ops", "Counterparty18 Bank (Hong Kong)", "S9", ["Equity Option", "Interest Rate Swap"]),
    ("Counterparty11", "grenoll.cp11.example", "cmd.settlements", "Counterparty11 Commodities", "S10", ["Commodity Swap", "FX Forward"]),
    ("Counterparty11", "grenoll.cp11.example", "eqd.settlements", "Counterparty11 Equities Synthetics", "S11", ["Equity Swap"]),
    ("Counterparty19", "sethra.cp19.example", "eqd.settlements", "Counterparty19 Financial Products Inc", "S12", ["Equity Swap", "Equity Option"]),
]
SCENARIO_CPS = []
for name, dom, local, entity, scen, prods in SCEN:
    code = "CP" + name.replace("Counterparty", "")
    SCENARIO_CPS.append({
        "name": name, "code": code, "bic": f"CPTY{dom[:2].upper()}2L", "domain": dom,
        "email": f"{local}@{dom}", "backup": f"settlements.backup@{dom}", "entity": entity,
        "kind": "scenario", "fmt": scen, "scenario": scen, "products": prods,
    })
 
ALL_CPS = COUNTERPARTIES + SCENARIO_CPS
 
# ------------------------------------------------------------------ counterparty directory (exact email -> org)
DIRECTORY_ROWS = []   # (email_address, org, entity, product_scope, dl_role, notes)
for cp in COUNTERPARTIES:
    d = cp["domain"]
    DIRECTORY_ROWS += [
        (f"otc.settlements@{d}", cp["name"], f"{cp['name']} International plc", "OTC Derivatives", "primary", "generic counterparty"),
        (f"eq.settlements@{d}", cp["name"], f"{cp['name']} Securities Ltd", "Equity Derivatives", "primary", "generic counterparty"),
        (cp["backup"], cp["name"], f"{cp['name']} International plc", "All OTC (backup)", "backup", "generic counterparty"),
    ]
for cp in SCENARIO_CPS:
    DIRECTORY_ROWS.append((cp["email"], cp["name"], cp["entity"], "/".join(cp["products"])[:40], "primary", f"scenario {cp['scenario']}"))
    DIRECTORY_ROWS.append((cp["backup"], cp["name"], cp["entity"], "All OTC (backup)", "backup", f"scenario {cp['scenario']}"))
# dedupe by email address (CP11/GS has several DLs but one shared backup)
_seen = set(); DIRECTORY_ROWS = [r for r in DIRECTORY_ROWS if not (r[0].lower() in _seen or _seen.add(r[0].lower()))]
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
    return START + timedelta(days=random.randint(0, span), hours=random.randint(0, 9),
                             minutes=random.randint(0, 59), seconds=random.randint(0, 59))
 
def money(n, sign=True):
    return f"{n:,.2f}" if sign or n >= 0 else f"{n:,.2f}"
 
def dnum(n):  # plain grouped number
    return f"{n:,.2f}"
 
def fmt_date(d, style="iso"):
    return {"iso": d.isoformat(), "dmy": d.strftime("%d/%m/%Y"), "dmon": d.strftime("%d-%b-%Y"),
            "dmony": d.strftime("%d %b %Y")}[style]
 
def maybe(v, p):
    return v if random.random() < p else ""
 
# ------------------------------------------------------------------ image (cross-platform font)
def _load_font(size, bold=False):
    cands = (["Arial Bold.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf", "Helvetica.ttc"] if bold
             else ["Arial.ttf", "arial.ttf", "DejaVuSans.ttf", "Helvetica.ttc"])
    dirs = ["/System/Library/Fonts/Supplemental/", "/Library/Fonts/", "C:/Windows/Fonts/",
            "/usr/share/fonts/truetype/dejavu/", "/usr/share/fonts/truetype/liberation/", ""]
    for dd in dirs:
        for n in cands:
            try:
                return ImageFont.truetype(dd + n, size)
            except Exception:
                continue
    return ImageFont.load_default()
_FONT = _load_font(14); _FONT_B = _load_font(15, bold=True)
 
def render_png(title, sections):
    pad, lh, w = 16, 24, 600
    nrows = sum(len(p) + (1 if h else 0) for h, p in sections)
    img = Image.new("RGB", (w, pad * 2 + lh * (nrows + 1)), "white"); d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, lh + pad // 2], fill=(31, 58, 95)); d.text((pad, 8), title, fill="white", font=_FONT_B)
    y = pad + lh + 6
    for header, pairs in sections:
        if header:
            d.text((pad, y), header, fill=(31, 58, 95), font=_FONT_B); y += lh
        for label, val in pairs:
            d.text((pad, y), str(label), fill=(90, 90, 90), font=_FONT)
            d.text((pad + 230, y), str(val), fill=(10, 10, 10), font=_FONT)
            d.line([pad, y + lh - 6, w - pad, y + lh - 6], fill=(228, 228, 228)); y += lh
    buf = io.BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()
 
# ------------------------------------------------------------------ internal cash-flow store
cashflows = []        # the bank's internal cash-flow records
_cf_seq = 0
_bk_seq = 0
def _ssi_for(cp, ccy):
    bank = random.choice(SSI_BANKS); inter = random.choice(INTERMEDIARIES)
    return {"bank_name": bank[0], "bank_bic": bank[1], "account_number": str(random.randint(10**9, 10**10 - 1)),
            "beneficiary_name": cp["name"], "beneficiary_bic": cp["bic"],
            "intermediary_bank": f"{inter[0]} ({inter[1]})"}
 
def gen_cashflow(cp, dt, product=None, shows_ssi=True, vdate=None):
    """Create one email cash flow + (unless Unmatched) the matching internal record. Returns a dict.
    shows_ssi: whether this email format actually prints SSI (else the email carries no SSI at all).
    vdate: shared value date for the whole email (an affirmation email settles for one value date)."""
    global _cf_seq, _bk_seq
    product = product or random.choice(cp["products"])
    ccy = random.choice(CCYS)
    amount = round(random.uniform(500, 6_000_000), 2)
    if random.random() < 0.12:
        amount = -amount  # reversal / receive-side
    direction = "Receive" if amount < 0 else random.choice(["Pay", "Receive"])
    value_date = vdate or (dt + timedelta(days=random.choice([1, 2, 2, 3]))).date()
    cf_type = CASHFLOW_TYPE[product]
    ssi = _ssi_for(cp, ccy)
    # counterparty's own refs (may differ from the bank's internal refs)
    cp_trade_ref = str(random.randint(340_000_000, 355_000_000))
    prod_ref = str(random.randint(90_000_000, 106_000_000))
    bank_trade_ref = f"NB-{random.randint(100000, 999999)}"
    _bk_seq += 1
    booking_ref = f"BK-{_bk_seq:05d}"
 
    # outcome distribution — AlreadySettled is a deliberately RARE edge case (~1.5%)
    roll = random.random()
    if roll < 0.015:
        outcome = "AlreadySettled"
    elif roll < 0.55:
        outcome = "Agreed"
    elif roll < 0.70:
        outcome = "AmountMismatch"
    elif roll < 0.82:
        outcome = "SSIMissing" if shows_ssi else "Agreed"  # SSIMissing only if the email prints SSI
    else:
        outcome = "Unmatched"
 
    cf_id = ""
    email_ssi = dict(ssi) if shows_ssi else None   # SSI in ground truth ONLY if the email format prints it
    if outcome != "Unmatched":
        _cf_seq += 1
        cf_id = f"CF-{_cf_seq:06d}"
        status = "Settled" if outcome == "AlreadySettled" else "Not Settled"
        int_amount = amount
        if outcome == "AmountMismatch":
            int_amount = round(amount + random.choice([-1, 1]) * random.uniform(50, 5000), 2)
        if outcome == "SSIMissing":
            email_ssi = None  # email omits SSI; internal has it
        # the counterparty's ref may or may not equal what the bank recorded internally
        # (~50% same, ~20% not recorded, ~30% a different internal counterparty ref)
        _rr = random.random()
        internal_cp_ref = (cp_trade_ref if _rr < 0.5 else ("" if _rr < 0.7 else str(random.randint(340_000_000, 355_000_000))))
        cashflows.append({
            "cashflow_id": cf_id, "booking_ref": booking_ref, "counterparty_org_name": cp["name"],
            "counterparty_entity": cp["entity"], "product_type": product, "trade_system": PRODUCT_SYSTEM[product],
            "cashflow_type": cf_type, "currency": ccy, "amount": int_amount,
            "direction": direction, "value_date": str(value_date), "status": status,
            "bank_trade_ref": bank_trade_ref, "counterparty_trade_ref": internal_cp_ref, "product_ref": prod_ref,
            "ssi_bank_name": ssi["bank_name"], "ssi_bank_bic": ssi["bank_bic"], "ssi_account_number": ssi["account_number"],
            "ssi_beneficiary_name": ssi["beneficiary_name"], "ssi_beneficiary_bic": ssi["beneficiary_bic"],
            "ssi_intermediary_bank": ssi["intermediary_bank"],
        })
    action = {"Agreed": "Affirm to counterparty", "AmountMismatch": "Query counterparty",
              "SSIMissing": "Obtain SSI from counterparty", "Unmatched": "Escalate to Middle Office",
              "AlreadySettled": "Highlight to analyst"}[outcome]
    return {
        "product": product, "cashflow_type": cf_type, "currency": ccy,
        "amount": amount, "direction": direction, "value_date": value_date,
        "cp_trade_ref": cp_trade_ref, "prod_ref": prod_ref, "ssi": email_ssi,
        "outcome": outcome, "matched_cashflow_id": cf_id, "action": action,
        "is_allege": "N" if outcome == "Agreed" else "Y",
    }
 
# extra internal cash flows with no matching email
def add_orphan_cashflows(n):
    global _cf_seq, _bk_seq
    for _ in range(n):
        cp = random.choice(ALL_CPS); product = random.choice(cp["products"]); ccy = random.choice(CCYS)
        ssi = _ssi_for(cp, ccy); _cf_seq += 1; _bk_seq += 1
        cashflows.append({
            "cashflow_id": f"CF-{_cf_seq:06d}", "booking_ref": f"BK-{_bk_seq:05d}", "counterparty_org_name": cp["name"],
            "counterparty_entity": cp["entity"], "product_type": product, "trade_system": PRODUCT_SYSTEM[product],
            "cashflow_type": CASHFLOW_TYPE[product], "currency": ccy,
            "amount": round(random.uniform(1000, 2_000_000), 2), "direction": random.choice(["Pay", "Receive"]),
            "value_date": str(rand_dt().date()), "status": random.choice(["Not Settled", "Not Settled", "Not Settled", "Settled"]),
            "bank_trade_ref": f"NB-{random.randint(100000, 999999)}",
            "counterparty_trade_ref": str(random.randint(340_000_000, 355_000_000)),
            "product_ref": str(random.randint(90_000_000, 106_000_000)),
            "ssi_bank_name": ssi["bank_name"], "ssi_bank_bic": ssi["bank_bic"], "ssi_account_number": ssi["account_number"],
            "ssi_beneficiary_name": ssi["beneficiary_name"], "ssi_beneficiary_bic": ssi["beneficiary_bic"],
            "ssi_intermediary_bank": ssi["intermediary_bank"],
        })
 
# ------------------------------------------------------------------ cash-flow display helpers
def cf_pairs(cf):
    """label:value pairs for a cash flow (generic renderers)."""
    p = [("Reference", cf["cp_trade_ref"]), ("Product", cf["product"]), ("Cashflow Type", cf["cashflow_type"]),
         ("Currency", cf["currency"]), ("Amount", dnum(cf["amount"])),
         ("Direction", cf["direction"]), ("Value Date", fmt_date(cf["value_date"]))]
    if cf["ssi"]:
        s = cf["ssi"]
        p += [("SSI Bank (BIC)", f'{s["bank_name"]} ({s["bank_bic"]})'), ("SSI Account", s["account_number"]),
              ("Beneficiary (BIC)", f'{s["beneficiary_name"]} ({s["beneficiary_bic"]})')]
    return p
 
def kv_block(pairs, sep=": "):
    w = max(len(l) for l, _ in pairs)
    return "\n".join(f"{l.ljust(w)}{sep}{v}" for l, v in pairs)
 
def text_table(headers, rows):
    widths = [max(len(str(headers[i])), *[len(str(r[i])) for r in rows]) for i in range(len(headers))]
    line = lambda vals: " | ".join(str(v).ljust(widths[i]) for i, v in enumerate(vals))
    return "\n".join([line(headers), "-+-".join("-" * w for w in widths)] + [line(r) for r in rows])
 
def html_table(headers, rows, title=None):
    head = "".join(f"<th style='background:#1F3A5F;color:#fff;padding:4px 8px'>{h}</th>" for h in headers)
    body = "".join("<tr>" + "".join(f"<td style='padding:3px 8px;border:1px solid #ccc'>{c}</td>" for c in r) + "</tr>" for r in rows)
    cap = f"<caption style='text-align:left;font-weight:bold;padding:4px'>{title}</caption>" if title else ""
    return f"<table style='border-collapse:collapse;font-family:Arial;font-size:12px'>{cap}<thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"
 
# ------------------------------------------------------------------ GENERIC renderers (CP1-10)
def r_semi(cfs):
    blocks = [(f"Cash flow {i+1}:\n" if len(cfs) > 1 else "") + kv_block(cf_pairs(c)) for i, c in enumerate(cfs)]
    return "Dear Settlements,\n\nPlease confirm the settlement of the below cash flow(s):\n\n" + "\n\n".join(blocks) + "\n\nKindly agree and advise settlement.\n", None, []
 
def r_free(cfs):
    items = "; ".join(f"{c['currency']} {dnum(c['amount'])} ({c['direction']}, value {fmt_date(c['value_date'])}, ref {c['cp_trade_ref']})" for c in cfs)
    return f"Hi team, please agree the following payment(s) due to settle: {items}. Kindly confirm. Thanks.\n", None, []
 
def r_swift(cfs):
    msgs = []
    for c in cfs:
        lines = [f":20:{c['cp_trade_ref']}", f":32A:{c['value_date'].strftime('%y%m%d')}{c['currency']}{abs(c['amount']):.2f}".replace('.', ','),
                 f":21:{c['prod_ref']}", f":72:/{c['direction'].upper()}/ {c['cashflow_type']}"]
        if c["ssi"]:
            lines.append(f":57A:{c['ssi']['bank_bic']}")
        msgs.append("--- begin message ---\n" + "\n".join(lines) + "\n--- end message ---")
    return "\n".join(msgs) + "\n", None, []
 
def r_html(cfs):
    headers = ["Reference", "Product", "Ccy", "Amount", "Direction", "Value Date", "SSI"]
    rows = [[c["cp_trade_ref"], c["product"], c["currency"], dnum(c["amount"]), c["direction"], fmt_date(c["value_date"]),
             (c["ssi"]["bank_bic"] + "/" + c["ssi"]["account_number"]) if c["ssi"] else ""] for c in cfs]
    html = f"<html><body><p>Please confirm the cash flow(s) below.</p>{html_table(headers, rows)}<p>Kindly agree by value date.</p></body></html>"
    text = "Please confirm the cash flow(s) below.\n\n" + text_table(headers, rows) + "\n"
    return text, html, []
 
def r_mixed(cfs):
    headers = ["Reference", "Product", "Ccy", "Amount", "Direction", "Value Date"]
    rows = [[c["cp_trade_ref"], c["product"], c["currency"], dnum(c["amount"]), c["direction"], fmt_date(c["value_date"])] for c in cfs]
    return "Hi team,\n\nWe'd like to confirm the payment(s) in the table below; please advise settlement.\n\n" + text_table(headers, rows) + "\n\nRegards\n", None, []
 
def r_forward(cfs):
    inner = "\n".join(">> " + l for c in cfs for l in kv_block(cf_pairs(c)).split("\n") + [">>"])
    return ("FYI - forwarding our desk's confirmation for your settlement. Please action.\n\n"
            "----- Forwarded message -----\n> From: Trading Desk\n> To: Operations\n> Subject: Cash flows to confirm\n>\n" + inner + "\n"), None, []
 
def r_reply(cfs):
    top = "\n\n".join(kv_block(cf_pairs(c)) for c in cfs)
    return ("Adding the cash-flow details you asked for, please confirm:\n\n" + top +
            "\n\nThanks\n\nOn earlier date, Settlements <bank@settlements.com> wrote:\n"
            "> Please send the economics and SSI for the cash flows pending confirmation. Regards, Settlements.\n"), None, []
 
def r_image(cfs):
    cid = make_msgid(domain="outlook.com")[1:-1]
    sections = [(f"Cash flow {i+1}" if len(cfs) > 1 else None, cf_pairs(c)) for i, c in enumerate(cfs)]
    png = render_png("SETTLEMENT CONFIRMATION", sections)
    text = "Hi team, please confirm the cash flow(s) shown in the screenshot below.\nRegards\n"
    html = f"<html><body><p>Hi team, please confirm the cash flow(s) shown in the screenshot below.</p><img src='cid:{cid}'><p>Regards</p></body></html>"
    return text, html, [(cid, png)]
 
GENERIC = {"SemiStructured": r_semi, "FreeForm": r_free, "TabularPlain": r_swift, "TabularHTML": r_html,
           "MixedProseTable": r_mixed, "ForwardedThread": r_forward, "ReplyChain": r_reply, "ImageEmbedded": r_image}
 
# ------------------------------------------------------------------ SCENARIO renderers (CP11-17, S1-S8)
def pr(cf):  # P/R short
    return "P" if cf["direction"] == "Pay" else "R"
 
def s1(cfs, cp):  # terse "please agree" option premium
    ent = f"{cp['code']}IL:NFPS"
    lines = "\n".join(f"{c['currency']} {pr(c)} {dnum(abs(c['amount']))}" for c in cfs)
    return f"Hi,\n\nPlease agree :\n\n{ent}\n\n{lines}\n\nRegards,\n{cp['name']} Markets Operations\n", None, []
 
def s2(cfs, cp):  # fund settlement, single HTML table + FFC
    headers = ["Fund", "Book", "Portfolio", "TradeId", "ValueDate", "Currency", "Type", "FlowAmount", "Client Ref No", "Cpty SSI Wire Ref"]
    rows = []
    for c in cfs:
        rows.append(["DMF I", "SGIL", "SGIL_DMF", c["prod_ref"], fmt_date(c["value_date"], "dmy"), c["currency"], "SW",
                     dnum(c["amount"]), c["cp_trade_ref"], (c["ssi"]["account_number"] if c["ssi"] else "")])
    ffc = f"FFC: {random.randint(10000000,99999999)} {cp['name'].upper()} MASTER FUND L.P"
    html = f"<html><body><p>Please can you confirm if you agree with the below payments for value date {fmt_date(cfs[0]['value_date'],'dmy')}?</p>{html_table(headers, rows)}<p>SSI attached and FFC below.<br>{ffc}</p><p>Thanks,</p></body></html>"
    text = f"Please can you confirm if you agree with the below payments for value date {fmt_date(cfs[0]['value_date'],'dmy')}?\n\n" + text_table(headers, rows) + f"\n\nSSI attached and FFC below.\n{ffc}\n\nThanks,\n"
    return text, html, []
 
def s3(cfs, cp):  # multiple cash-flow tables + signature/escalation
    parts = []
    cpart = random.choice(BANK_COUNTERPARTS)
    for c in cfs:
        headers = ["Flow Id", "Trn Id", "Group", "Counterpart", "Currency", "Amount", "Value Date", "Paying To", "SWIFT1", "Pay/Receive", "Typology"]
        row = [f"M{random.randint(10**10,10**11-1)}", c["cp_trade_ref"], "IRD|OSWP|SMP", cpart, c["currency"],
               dnum(c["amount"]), fmt_date(c["value_date"], "dmon"), (c["ssi"]["account_number"] if c["ssi"] else ""),
               (c["ssi"]["bank_bic"] if c["ssi"] else ""), c["direction"], c["cashflow_type"]]
        parts.append(html_table(headers, [row]))
    sig = (f"<br><b>{cp['name']} Ops</b><br>OTC Derivatives Ops<br>{cp['entity']}<br>"
           f"E-mail: {cp['email']}<br>Escalation 1: escalation1@{cp['domain']}<br>Escalation 2: escalation2@{cp['domain']}")
    html = f"<html><body><p>Hi Team,</p><p>Please confirm the following cash flows.</p>{''.join('<p>'+p+'</p>' for p in parts)}{sig}</body></html>"
    text = "Hi Team,\n\nPlease confirm the following cash flows.\n\n" + "\n\n".join(
        text_table(["Flow Id", "Ccy", "Amount", "Value Date", "Pay/Receive", "Typology"],
                   [[f"M{random.randint(10**10,10**11-1)}", c["currency"], dnum(c["amount"]), fmt_date(c["value_date"], "dmon"), c["direction"], c["cashflow_type"]]]) for c in cfs) + f"\n\n{cp['name']} Ops | {cp['entity']}\n"
    return text, html, []
 
# Harry Potter names in Hangul for S4
HP_HANGUL = ["해리 포터", "헤르미온느 그레인저", "론 위즐리", "알버스 덤블도어", "세베루스 스네이프", "미네르바 맥고나걸", "루비우스 해그리드"]
def s4(cfs, cp):  # bilingual Korean, forwarded original, ELS
    who = random.choice(HP_HANGUL)
    headers = ["SHS REFERENCE", "STATE", "Counterpart", "Settlement Date", "Direction", "Currency", "Settlement Amount", "C.P REFERENCE", "BANK NAME"]
    cpart = "BANK"
    rows = [[f"OTC-S-ELS-{random.randint(27000,27999)}", "CD", cpart, fmt_date(c["value_date"], "dmon"), f"{cp['name'][:3].upper()} {c['direction'].upper()}",
             c["currency"], dnum(c["amount"]), c["cp_trade_ref"], (c["ssi"]["bank_name"] if c["ssi"] else "")] for c in cfs]
    body = ("--- Original Message ---\n"
            f'From: "파생결제팀" <{cp["email"]}>\nTo: {BANK_MAILBOX}\nSubject: [SHS] BANK Settlement\n\n'
            "안녕하세요 (Hello Team),\n\nPlease kindly confirm the below.\n\n" + text_table(headers, rows)
            + f"\n\nThanks and regards,\n{who}\nSettlement Department 2 | OTC Derivatives Operations\n{cp['entity']}\n")
    return body, None, []
 
def s5(cfs, cp):  # NDF close-out, caution + CONFIDENTIAL, yellow table
    banner = ("< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n"
              "Do not click links or open attachments if you do not believe the email is legitimate.\n\nCONFIDENTIAL\n\n")
    headers = ["Entity", "Amount", "Ccy", "Value Date", "R", "Type", "Ref"]
    rows = [[random.choice(BANK_COUNTERPARTS), dnum(abs(c["amount"])), c["currency"], fmt_date(c["value_date"], "dmon"), pr(c), "NDF", c["cp_trade_ref"]] for c in cfs]
    html = f"<html><body><pre>{banner}</pre><p>Hi Team,</p><p>Could you please confirm the below NDF Close out amount and SSI.</p>{html_table(headers, rows)}<p>Regards,<br>{cp['name']} FXD Operations</p></body></html>"
    text = banner + "Hi Team,\n\nCould you please confirm the below NDF Close out amount and SSI.\n\n" + text_table(headers, rows) + f"\n\nRegards,\n{cp['name']} FXD Operations\n"
    return text, html, []
 
def s6(cfs, cp):  # wide payment-affirmation table + missing-SSI prompts
    headers = ["Entity", "Customer Name", "PaymentDate", "Amount", "Currency", "Direction", "Payment Instructions", "Reference", "System Ref", "Product Class", "TradeDate", "MaturityDate"]
    rows = []
    for c in cfs:
        pinstr = (f"{c['ssi']['bank_name']} BIC:{c['ssi']['bank_bic']} A/C:{c['ssi']['account_number']}"
                  if c["ssi"] else "Please provide your payment instructions.")
        rows.append([cp["code"] + "GML", "BANK INTL PLC", fmt_date(c["value_date"], "dmon"), dnum(c["amount"]), c["currency"], c["direction"],
                     pinstr, c["cp_trade_ref"], c["prod_ref"], c["product"], fmt_date((c["value_date"] - timedelta(days=2))), fmt_date(c["value_date"], "dmon")])
    html = (f"<html><body><p>Below are upcoming payments scheduled to settle on {fmt_date(cfs[0]['value_date'],'dmon')}.</p>"
            f"<p>We are requesting you to affirm/agree each of the payments and payment instructions.</p>"
            f"<p><b>In case of null, please be advised that we require all economic details be agreed.</b></p>{html_table(headers, rows)}<p>Regards,<br>{cp['name']} Operations</p></body></html>")
    text = (f"Below are upcoming payments scheduled to settle on {fmt_date(cfs[0]['value_date'],'dmon')}.\n"
            "We are requesting you to affirm/agree each of the payments and payment instructions.\n\n"
            + text_table(["Reference", "Ccy", "Amount", "Direction", "Payment Instructions"],
                         [[r[7], r[4], r[3], r[5], r[6]] for r in rows]) + f"\n\nRegards,\n{cp['name']} Operations\n")
    return text, html, []
 
def s7(cfs, cp):  # EQ cashflows, on-behalf-of, gross note, UnMatched status
    banner = "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n\n"
    headers = ["DEAL ID", "SDS ID", "ORIGINAL VALUE DATE", "CCY", "SETTLEMENT AMOUNT", "P/R", "SI STATUS", "BOOK", "SETTLEMENT STATUS", "TRADE ID", "LEGAL ENTITY", "TEAM LOCATION"]
    rows = [[f"EDR{random.randint(10**7,10**8-1)}EU", c["prod_ref"], fmt_date(c["value_date"], "dmon"), c["currency"], dnum(c["amount"]),
             cp["name"][:4] + " " + c["direction"], random.choice(["Allocated", "UnMatched"]), str(random.randint(20000, 29999)),
             random.choice(["Matched", "UnMatched"]), c["cp_trade_ref"], str(random.randint(40000000, 49999999)), "CH EDG"] for c in cfs]
    html = (f"<html><body><pre>{banner}</pre><p>Hi Team,</p><p>Please can you confirm the below cash flows for value date {fmt_date(cfs[0]['value_date'],'dmon')}.</p>"
            f"<p>Kindly be advised that unless otherwise specified, {cp['name']} will settle the below cashflows as Gross.</p>{html_table(headers, rows)}"
            f"<p>**Any payments due to you from {cp['name']} shall be remitted as per the usual SSI's in our records**</p><p>Kind Regards,<br>{cp['name']}</p></body></html>")
    text = (banner + f"Hi Team,\n\nPlease can you confirm the below cash flows for value date {fmt_date(cfs[0]['value_date'],'dmon')}.\n"
            f"Kindly be advised that unless otherwise specified, {cp['name']} will settle the below cashflows as Gross.\n\n"
            + text_table(["DEAL ID", "CCY", "SETTLEMENT AMOUNT", "P/R", "SI STATUS", "SETTLEMENT STATUS"],
                         [[r[0], r[3], r[4], r[5], r[6], r[8]] for r in rows]) + f"\n\nKind Regards,\n{cp['name']}\n")
    return text, html, []
 
def s8(cfs, cp):  # high-volume flat table, deem-good, info-classification banner
    banner = "Information Classification: [ ] Secret  [x] Confidential  [ ] Internal  [ ] Unclassified / Public\n\n"
    headers = ["Prod. Ref", "Trade reference", "Cur.", "ValueDate", "Quantity"]
    rows = [[c["prod_ref"], c["cp_trade_ref"], c["currency"], fmt_date(c["value_date"], "dmy"), dnum(c["amount"])] for c in cfs]
    text = (banner + "Hi All,\n\nKindly confirm the below and settle in gross by VALUE DATE. If no reply, we will deem "
            "the below as good.\n\n" + text_table(headers, rows) + f"\n\nBest Regards,\nTreasury Operations Settlement | {cp['name']} SG\n")
    return text, None, []
 
def s9(cfs, cp):  # upfront-fee confirmation; SSI is a PDF attachment; asks for trade ref + UTI
    banner = "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n\n"
    headers = ["Reference", "Trade Date", "Maturity Date", "Notional Ccy", "Notional Amount", "Fee Ccy", "Upfront Fee Amount", "Counterparty"]
    rows = []
    for c in cfs:
        td = c["value_date"] - timedelta(days=random.randint(30, 120))
        mat = c["value_date"] + timedelta(days=random.randint(365, 365 * 3))
        rows.append([c["cp_trade_ref"], td.strftime("%d%m%y"), mat.strftime("%d%m%y"), random.choice(["CNY", "HKD", "USD"]),
                     dnum(random.choice([1, 2, 5, 10]) * 1_000_000), c["currency"], dnum(abs(c["amount"])), "BANK Singapore"])
    text = (banner + "Dear Team,\n\nBelow trade(s) are done with your Bank today. According to the attached term sheet "
            f"(if any), please confirm below payment value on {fmt_date(cfs[0]['value_date'],'dmon')}.\n"
            "Please affirm the trade by providing your trade reference number and UTI by return e-mail.\n"
            "**As per attached SSI**\n\n" + text_table(headers, rows) +
            f"\n\n**Term sheet will be provided later**\nBest Regards,\n{cp['name']} Settlement Operations Department\n")
    return text, None, []
 
def s10(cfs, cp):  # system-generated settlement summary with recipient-fill response columns
    headers = ["Confirmation Status", "Commentary", "Netting ID", "Version", "Entity", "Customer Account",
               "Pay/Rec", "Currency", "Settlement Amount", "Payment Date", "Number of Trades"]
    rows = [["Please Select", "", c["cp_trade_ref"], "2", cp["code"] + "IL", "BANK FINANCIAL PRODUCTS & SERVICES INC",
             c["direction"], c["currency"], dnum(c["amount"]), fmt_date(c["value_date"], "dmon"), str(random.randint(1, 6))] for c in cfs]
    html = ("<html><body><p>The Settlements team would like to notify you of settlement amounts that require your "
            "feedback as soon as possible.</p><p>This is a system-generated notification - please select Reply All, "
            "then follow the below instructions. Please provide responses in the table.</p><ol>"
            "<li>If the Settlement Amount is incorrect, modify the first two columns of each row.</li>"
            "<li>When remitting funds, please enter the Netting ID into Field 72 or the FFC field in your SWIFT payment.</li>"
            "<li>Please do <b>not</b> provide commentary outside of the table below.</li></ol>"
            f"{html_table(headers, rows, 'Settlement Summary')}</body></html>")
    text = ("The Settlements team would like to notify you of settlement amounts that require your feedback.\n"
            "This is a system-generated notification - please select Reply All and provide responses in the table.\n\n"
            + text_table(["Confirmation Status", "Netting ID", "Pay/Rec", "Currency", "Settlement Amount", "Payment Date", "# Trades"],
                         [[r[0], r[2], r[6], r[7], r[8], r[9], r[10]] for r in rows]) + "\n")
    return text, html, []
 
def s11(cfs, cp):  # equity total-return-swap settlement with component breakdown
    headers = ["Payment Date", "Settlement Currency", "Client Account", "Total Equity", "Total Dividend",
               "Total Interest", "Total Withholding", "Total Settlement Amount", "Group Id"]
    rows = []
    for c in cfs:
        t = c["amount"]
        div = round(abs(t) * random.uniform(0, 0.05), 2); intr = round(abs(t) * random.uniform(0, 0.05), 2)
        wht = round(abs(t) * random.uniform(0, 0.02), 2); eq = round(t - div - intr + wht, 2)  # components sum to total
        rows.append([fmt_date(c["value_date"], "dmy"), c["currency"], str(random.randint(10**8, 10**9)),
                     dnum(eq), dnum(div), dnum(intr), dnum(wht), dnum(t), c["cp_trade_ref"]])
    html = ("<html><body><p>Hi Team,</p><p>We have below cashflow for settlement for VD today. Requesting you to please "
            f"check and advise if we are good to settle the same.</p>{html_table(headers, rows)}"
            f"<p>Regards,<br>{cp['name']} Equities Synthetics Settlements</p></body></html>")
    text = ("Hi Team,\n\nWe have below cashflow for settlement for VD today. Please check and advise if we are good to "
            "settle the same.\n\n" + text_table(["Payment Date", "Ccy", "Total Equity", "Total Dividend", "Total Interest",
            "Total Withholding", "Total Settlement Amount", "Group Id"], [[r[0], r[1], r[3], r[4], r[5], r[6], r[7], r[8]] for r in rows])
            + f"\n\nRegards,\n{cp['name']} Equities Synthetics Settlements\n")
    return text, html, []
 
def s12(cfs, cp):  # high-value EQ pre-confirmation; signed PAY/RECV table + detailed SSI blocks
    headers = ["Deal", "Value Date", "Currency", "Payment Direction", "Amount", "Entity", "Counterparty Legal Name"]
    rows = [[c["cp_trade_ref"], fmt_date(c["value_date"], "dmy"), c["currency"],
             ("PAY" if c["direction"] == "Pay" else "RECV"),
             dnum(-abs(c["amount"]) if c["direction"] == "Pay" else abs(c["amount"])),
             random.choice(["BNKA", "MLIL"]), random.choice(["BANK Global Financial Products Inc", "BANK INTERNATIONAL PLC"])] for c in cfs]
    blocks = []
    for c in cfs:
        if c["ssi"]:
            s = c["ssi"]
            blocks.append(f"{c['direction'].upper()}: {s['bank_name']} ({s['bank_bic']})  CPTY: {cp['name']}  "
                          f"FAO: {s['beneficiary_name']} ({s['beneficiary_bic']})  Acct: {s['account_number']}  NIP:")
    text = (f"Hi Team,\n\nPlease confirm the below cash for value date VD {fmt_date(cfs[0]['value_date'],'dmon')}.\n\n"
            + text_table(headers, rows) + "\n\n" + "\n".join(blocks[:4]) +
            f"\n\nThanks & Regards,\n{cp['name']} EQD Settlements\nEscalation level 1 - escalation1@{cp['domain']}\n"
            f"Escalation level 2 - escalation2@{cp['domain']}\n[eTask_Case_ID: {random.randint(100000000, 199999999)}]\n")
    return text, None, []
 
SCENARIO = {"S1": s1, "S2": s2, "S3": s3, "S4": s4, "S5": s5, "S6": s6, "S7": s7, "S8": s8,
            "S9": s9, "S10": s10, "S11": s11, "S12": s12}
SCEN_CFCOUNT = {"S1": (1, 3), "S2": (1, 3), "S3": (2, 5), "S4": (2, 5), "S5": (1, 3),
                "S6": (2, 5), "S7": (1, 4), "S8": (20, 60),
                "S9": (1, 2), "S10": (1, 3), "S11": (1, 2), "S12": (2, 5)}  # cash flows per email
# whether each format actually PRINTS SSI in the body (else the email carries no SSI -> blank in ground truth)
SHOWS_SSI = {"SemiStructured": True, "FreeForm": False, "TabularPlain": True, "TabularHTML": True,
             "MixedProseTable": False, "ForwardedThread": True, "ReplyChain": True, "ImageEmbedded": True,
             "S1": False, "S2": True, "S3": True, "S4": True, "S5": False, "S6": True, "S7": False, "S8": False,
             "S9": False, "S10": False, "S11": False, "S12": True}  # S9 SSI is a PDF attachment; S12 detailed blocks
 
def make_subject(fmt, cp, cfs):
    vd = fmt_date(cfs[0]["value_date"], "dmon")
    if fmt == "S1": return f"OPT Premium vd {vd} (FX Setts:Interbank-{random.randint(1000000, 9999999)})"
    if fmt == "S2": return f"{cp['code']} - BANK - SW - Settlement for {vd}"
    if fmt == "S3": return f"BANKBK/LDN - Cash Flow Confirmation - {vd}"
    if fmt == "S4": return f"[SHS] BANK Settlement VAL {vd}"
    if fmt == "S5": return f"BANKJPJTXXX/ BANKGB2LXXX - NDF SETTLEMENT - {vd}"
    if fmt == "S6": return f"{cp['code']} - BANK - {vd} - Upcoming Payment(s) Affirmation Request ({random.randint(10000, 99999)})"
    if fmt == "S7": return f"EMUCF_{cp['name']} EQ Cashflows for {vd} SDS -{random.randint(10000000, 99999999)}"
    if fmt == "S8": return f"[BANK] PAYMENT CONFIRMATION FOR VALUE {vd}"
    if fmt == "S9": return f"Upfront Fee due on {vd} Ref {cfs[0]['cp_trade_ref']} (BANK Singapore)"
    if fmt == "S10": return f"Settlement for Value Date {vd}, Ref Num {cfs[0]['cp_trade_ref']}"
    if fmt == "S11": return f"BANK vs {cp['code']} | {vd}"
    if fmt == "S12": return f"High Value//Pre-Confirmation / VD {vd}//BANK"
    return random.choice([f"Settlement confirmation for value {vd}", f"Please confirm cash flows - value {vd}",
                          f"Payment confirmation - {vd}", f"Cashflow affirmation for {vd}"])
 
# SSI in the ground truth reflects ONLY what each email format actually prints in the body.
SSI_KEYS = ["ssi_bank_name", "ssi_bank_bic", "ssi_account_number", "ssi_beneficiary_name",
            "ssi_beneficiary_bic", "ssi_intermediary_bank"]
SSI_GT = {"SemiStructured": "full", "TabularHTML": "full", "ForwardedThread": "full", "ReplyChain": "full",
          "ImageEmbedded": "full", "S6": "full", "S12": "full", "TabularPlain": "bic", "S3": "bic_acct", "S2": "acct", "S4": "bank"}
def gt_ssi(fmt, ssi):
    out = {k: "" for k in SSI_KEYS}
    if not ssi:
        return out
    keys = {"full": SSI_KEYS, "bic": ["ssi_bank_bic"], "acct": ["ssi_account_number"],
            "bic_acct": ["ssi_bank_bic", "ssi_account_number"], "bank": ["ssi_bank_name"],
            "none": []}[SSI_GT.get(fmt, "none")]
    m = {"ssi_bank_name": ssi["bank_name"], "ssi_bank_bic": ssi["bank_bic"], "ssi_account_number": ssi["account_number"],
         "ssi_beneficiary_name": ssi["beneficiary_name"], "ssi_beneficiary_bic": ssi["beneficiary_bic"],
         "ssi_intermediary_bank": ssi["intermediary_bank"]}
    for k in keys:
        out[k] = m[k]
    return out
 
# which non-SSI extraction fields each format actually prints (else blank in ground truth)
REF_HIDDEN = {"S1"}   # formats that do NOT print a counterparty reference
PRODUCT_SHOWN = {"SemiStructured", "TabularHTML", "MixedProseTable", "ForwardedThread", "ReplyChain",
                 "ImageEmbedded", "S5", "S6"}   # formats that print the product
 
# ------------------------------------------------------------------ assemble emails
NOISE_SENDERS = [
    ("Bloomberg Markets Wrap", "newsletter@bloomberg.nonoutlook.com", "Daily FX & Rates Wrap"),
    ("Reuters Eikon Alerts", "alerts@reuters.nonoutlook.com", "Market alert"),
    ("Bank IT Service Desk", "servicedesk@bank.nonoutlook.com", "Scheduled maintenance window"),
    ("HR Communications", "hr.comms@bank.nonoutlook.com", "Reminder: mandatory training"),
    ("SWIFT Network Notices", "no-reply@swift.nonoutlook.com", "Network advisory notice"),
    ("Vendor - DataFeed Co", "billing@datafeedco.nonoutlook.com", "Invoice available"),
    ("Internal Risk Reporting", "risk.reports@bank.nonoutlook.com", "EOD risk pack"),
    ("Facilities", "facilities@bank.nonoutlook.com", "Desk move notification"),
    ("LinkedIn", "notifications@linkedin.nonoutlook.com", "You have new notifications"),
    ("CLS Bank Service", "service@cls-group.nonoutlook.com", "CLS session times update"),
]
 
emails, ground_truth = [], []
_idx = 0
def next_id():
    global _idx; _idx += 1; return f"EML-{_idx:04d}"
 
def make_cc(cp):
    cc = []
    if random.random() < 0.5:
        cc.append(cp["backup"])
        if random.random() < 0.2:
            cc.append(random.choice(BANK_CC_POOL))
    return cc
 
def emit_email(cp, dt, fmt, is_primary, n_cf, scenario_render=None):
    eid = next_id()
    shows_ssi = SHOWS_SSI.get(fmt, True)
    vd = (dt + timedelta(days=random.choice([1, 2, 2, 3]))).date()  # one value date per email
    cfs = [gen_cashflow(cp, dt, shows_ssi=shows_ssi, vdate=vd) for _ in range(n_cf)]
    if scenario_render:
        text, html, images = scenario_render(cfs, cp)
    else:
        text, html, images = GENERIC[fmt](cfs)
    cc_list = make_cc(cp); cc = "; ".join(cc_list)
    derived = derive_counterparty([cp["email"]] + cc_list)
    prefix = {"ForwardedThread": "FW: ", "ReplyChain": "RE: "}.get(fmt, "")
    on_behalf = fmt in ("S7", "S8", "S12")
    subj = prefix + make_subject(fmt, cp, cfs)
    attach = [("Standard_Settlement_Instructions.pdf", PDF_STUB)] if fmt == "S9" else []
    body_mime = "Plain text + image" if images else "Plain text + HTML" if html else "Plain text"
    if attach:
        body_mime += " + attachment"
    sender_disp = (f"{cp['name']} Ops On Behalf Of Markets Ops - Settlements" if on_behalf else f"{cp['name']} Settlements")
    emails.append({"email_id": eid, "sender_name": sender_disp, "sender_email": cp["email"], "sender_domain": cp["domain"],
                   "cc": cc, "email_type": "SettlementAffirmation", "email_format": fmt, "body_mime": body_mime,
                   "subject": subj, "received_date": format_datetime(dt), "_dt": dt, "_render": (text, html, images),
                   "_attach": attach, "folder2": "N"})
    for k, c in enumerate(cfs, 1):
        gs = gt_ssi(fmt, c["ssi"])
        ground_truth.append({
            "email_id": eid, "cashflow_index": f"{eid}#{k}", "sender_name": sender_disp, "sender_email": cp["email"],
            "sender_domain": cp["domain"], "from_email": cp["email"], "to_email": BANK_MAILBOX, "cc_email": cc,
            "derived_counterparty_org_name": derived, "sender_category": "Counterparty", "email_type": "SettlementAffirmation",
            "email_format": fmt, "format_is_primary": is_primary, "body_mime": body_mime, "subject": subj,
            "received_date": format_datetime(dt),
            "counterparty_reference": ("" if fmt in REF_HIDDEN else c["cp_trade_ref"]),
            "product": (c["product"] if fmt in PRODUCT_SHOWN else ""),
            "currency": c["currency"], "amount": c["amount"],
            "direction": c["direction"], "value_date": str(c["value_date"]),
            "ssi_bank_name": gs["ssi_bank_name"], "ssi_bank_bic": gs["ssi_bank_bic"],
            "ssi_account_number": gs["ssi_account_number"], "ssi_beneficiary_name": gs["ssi_beneficiary_name"],
            "ssi_beneficiary_bic": gs["ssi_beneficiary_bic"], "ssi_intermediary_bank": gs["ssi_intermediary_bank"],
            "internal_match_status": c["outcome"], "matched_cashflow_id": c["matched_cashflow_id"],
            "match_basis": "Economic fields" if c["outcome"] != "Unmatched" else "", "is_allege": c["is_allege"], "action": c["action"],
        })
 
BLANK_CF = ("counterparty_reference", "product", "currency", "amount", "direction",
            "value_date", "ssi_bank_name", "ssi_bank_bic", "ssi_account_number", "ssi_beneficiary_name",
            "ssi_beneficiary_bic", "ssi_intermediary_bank")
 
def emit_ssi_email(dt):
    cp = random.choice(COUNTERPARTIES); eid = next_id(); ccy = random.choice(CCYS); ssi = _ssi_for(cp, ccy)
    is_update = random.random() < 0.45
    fmt = "SSI"   # SSI notes are always plain-text, not one of the structural formats
    body = (f"Dear Settlements,\n\n{'Please UPDATE' if is_update else 'For your verification, our'} standing settlement "
            f"instructions for {ccy}:\n\n  Bank (BIC)   : {ssi['bank_name']} ({ssi['bank_bic']})\n  Account      : {ssi['account_number']}\n"
            f"  Beneficiary  : {ssi['beneficiary_name']} ({ssi['beneficiary_bic']})\n\nRegards,\n{cp['name']}\n")
    cc_list = make_cc(cp); cc = "; ".join(cc_list)
    subj = f"{'Updated ' if is_update else ''}SSI {ccy} - settlement instructions"
    emails.append({"email_id": eid, "sender_name": f"{cp['name']} Settlements", "sender_email": cp["email"], "sender_domain": cp["domain"],
                   "cc": cc, "email_type": "SSIVerification", "email_format": fmt, "body_mime": "Plain text",
                   "subject": subj, "received_date": format_datetime(dt), "_dt": dt, "_render": (body, None, []), "folder2": "N"})
    row = {"email_id": eid, "cashflow_index": "", "sender_name": f"{cp['name']} Settlements", "sender_email": cp["email"],
           "sender_domain": cp["domain"], "from_email": cp["email"], "to_email": BANK_MAILBOX, "cc_email": cc,
           "derived_counterparty_org_name": derive_counterparty([cp["email"]] + cc_list), "sender_category": "Counterparty",
           "email_type": "SSIVerification", "email_format": fmt, "format_is_primary": "Y", "body_mime": "Plain text",
           "subject": subj, "received_date": format_datetime(dt), "internal_match_status": "NA", "matched_cashflow_id": "",
           "match_basis": "", "is_allege": "N", "action": "", "_ssi_outcome": "Update requested" if is_update else "Confirmed - no change"}
    for c in BLANK_CF: row[c] = ""
    ground_truth.append(row)
 
def emit_noise(dt):
    name, addr, subj = random.choice(NOISE_SENDERS); eid = next_id()
    body = f"Hello,\n\n{subj}.\n\nAutomated message - no action required by the settlements desk.\n\nRegards,\n{name}\n"
    emails.append({"email_id": eid, "sender_name": name, "sender_email": addr, "sender_domain": addr.split("@")[-1],
                   "cc": "", "email_type": "Noise", "email_format": "Noise", "body_mime": "Plain text",
                   "subject": subj, "received_date": format_datetime(dt), "_dt": dt, "_render": (body, None, []), "folder2": "N"})
    row = {"email_id": eid, "cashflow_index": "", "sender_name": name, "sender_email": addr, "sender_domain": addr.split("@")[-1],
           "from_email": addr, "to_email": BANK_MAILBOX, "cc_email": "", "derived_counterparty_org_name": "",
           "sender_category": "Noise", "email_type": "Noise", "email_format": "Noise", "format_is_primary": "",
           "body_mime": "Plain text", "subject": subj, "received_date": format_datetime(dt),
           "internal_match_status": "NA", "matched_cashflow_id": "", "match_basis": "", "is_allege": "N", "action": ""}
    for c in BLANK_CF: row[c] = ""
    ground_truth.append(row)
 
# --- 65 generic counterparty affirmations (CP1-10) ---
for _ in range(65):
    cp = random.choice(COUNTERPARTIES)
    is_primary = "Y" if random.random() < 0.8 else "N"
    fmt = cp["fmt"] if is_primary == "Y" else random.choice([f for f in GENERIC_FORMATS if f != cp["fmt"]])
    n_cf = random.choice([1, 1, 1, 2, 3])
    emit_email(cp, rand_dt(), fmt, is_primary, n_cf)
 
# --- 80 scenario affirmations (CP11-17, S1-S8, 10 each) ---
for cp in SCENARIO_CPS:
    scen = cp["scenario"]
    lo, hi = SCEN_CFCOUNT[scen]
    for _ in range(10):
        emit_email(cp, rand_dt(), scen, "Y", random.randint(lo, hi), scenario_render=SCENARIO[scen])
 
# --- 35 SSI + 100 noise ---
for _ in range(35):
    emit_ssi_email(rand_dt())
for _ in range(100):
    emit_noise(rand_dt())
 
add_orphan_cashflows(20)
 
# folder2: freshest 40 affirmations + 20 SSI
aff = [e for e in emails if e["email_type"] == "SettlementAffirmation"]
ssi = [e for e in emails if e["email_type"] == "SSIVerification"]
f2ids = {e["email_id"] for e in sorted(aff, key=lambda e: e["_dt"], reverse=True)[:40] + sorted(ssi, key=lambda e: e["_dt"], reverse=True)[:20]}
for e in emails:
    e["folder2"] = "Y" if e["email_id"] in f2ids else "N"
f2map = {e["email_id"]: e["folder2"] for e in emails}
for r in ground_truth:
    r["folder2"] = f2map[r["email_id"]]; r["folder1"] = "Y"
 
# ------------------------------------------------------------------ write .eml
def write_eml(folder, e):
    text, html, images = e["_render"]
    msg = EmailMessage()
    msg["From"] = f'{e["sender_name"]} <{e["sender_email"]}>'
    msg["To"] = BANK_TO
    if e.get("cc"):
        msg["Cc"] = e["cc"]
    msg["Date"] = e["received_date"]; msg["Subject"] = e["subject"]
    msg["Message-ID"] = make_msgid(domain=e["sender_domain"])
    msg["X-Email-ID"] = e["email_id"]; msg["X-Email-Type"] = e["email_type"]; msg["X-Email-Format"] = e["email_format"]
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
        shutil.copyfile(os.path.join(F1, f'{e["email_id"]}.eml'), os.path.join(F2, f'{e["email_id"]}.eml'))
 
# ------------------------------------------------------------------ xlsx
HFILL = PatternFill("solid", fgColor="1F3A5F"); HFONT = Font(color="FFFFFF", bold=True)
AFILL = PatternFill("solid", fgColor="FCE4E4")
def col_letter(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s
def style_header(ws, n):
    for c in range(1, n + 1):
        ws.cell(row=1, column=c).fill = HFILL; ws.cell(row=1, column=c).font = HFONT
    ws.freeze_panes = "A2"
 
# email_ground_truth.xlsx
wb = Workbook(); ws = wb.active; ws.title = "ground_truth"
GCOLS = ["email_id", "cashflow_index", "folder1", "folder2", "sender_name", "sender_domain",
         "from_email", "to_email", "cc_email", "derived_counterparty_org_name", "sender_category", "email_type",
         "email_format", "format_is_primary", "body_mime", "subject", "received_date", "counterparty_reference",
         "product", "currency", "amount", "direction", "value_date", "ssi_bank_name",
         "ssi_bank_bic", "ssi_account_number", "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank",
         "internal_match_status", "matched_cashflow_id", "match_basis", "is_allege", "action"]
ws.append(GCOLS)
for r in ground_truth:
    ws.append([r.get(c, "") for c in GCOLS])
    if r.get("is_allege") == "Y":
        for c in range(1, len(GCOLS) + 1):
            ws.cell(row=ws.max_row, column=c).fill = AFILL
style_header(ws, len(GCOLS))
for i in range(1, len(GCOLS) + 1):
    ws.column_dimensions[col_letter(i)].width = 16
 
# folder2 sheet
wf = wb.create_sheet("folder2")
FCOLS = ["email_id", "cashflow_index", "sender_name", "email_type", "email_format", "subject", "received_date",
         "scenario_type", "scenario_outcome", "is_allege", "action"]
wf.append(FCOLS)
for r in sorted([r for r in ground_truth if r["folder2"] == "Y"], key=lambda r: (r["email_id"], r.get("cashflow_index", ""))):
    if r["email_type"] == "SettlementAffirmation":
        st, so = "Cash-flow affirmation", r["internal_match_status"]
    else:
        st, so = "SSI Confirmation", f'{r.get("_ssi_outcome","Confirmed")} (not actioned)'
    wf.append([r["email_id"], r.get("cashflow_index", ""), r["sender_name"], r["email_type"], r["email_format"],
               r["subject"], r["received_date"], st, so, r["is_allege"], r.get("action", "")])
    if r.get("is_allege") == "Y":
        for c in range(1, len(FCOLS) + 1):
            wf.cell(row=wf.max_row, column=c).fill = AFILL
style_header(wf, len(FCOLS))
for i, w in enumerate([10, 12, 26, 20, 14, 44, 30, 20, 26, 8, 24], 1):
    wf.column_dimensions[col_letter(i)].width = w
 
# formats catalog sheet
wc = wb.create_sheet("formats")
wc.append(["format_code", "family", "used_by", "description"])
FORMAT_CATALOG = [
    ("SemiStructured", "generic", "CP1-10", "inline Label: value, inconsistent order"),
    ("FreeForm", "generic", "CP1-10", "unstructured prose, no field labels"),
    ("TabularPlain", "generic", "CP1-10", "SWIFT-style tag block"),
    ("TabularHTML", "generic", "CP1-10", "pasted HTML table"),
    ("MixedProseTable", "generic", "CP1-10", "prose intro + table"),
    ("ForwardedThread", "generic", "CP1-10", "forwarded internal mail, quoted"),
    ("ReplyChain", "generic", "CP1-10", "reply on top, quoted bank content below"),
    ("ImageEmbedded", "generic", "CP1-10", "inline PNG screenshot of a blotter"),
    ("S1", "scenario", "CP11", "terse 'please agree' option premium"),
    ("S2", "scenario", "CP12", "fund settlement, single HTML table, FFC"),
    ("S3", "scenario", "CP13", "multiple cash-flow tables + escalation block"),
    ("S4", "scenario", "CP14", "bilingual Korean, forwarded original, ELS"),
    ("S5", "scenario", "CP13", "NDF close-out, external-caution + CONFIDENTIAL, yellow table"),
    ("S6", "scenario", "CP15", "wide payment-affirmation table, missing-SSI prompts"),
    ("S7", "scenario", "CP16", "EQ cashflows, on-behalf-of, gross note, UnMatched status"),
    ("S8", "scenario", "CP17", "high-volume flat table, deem-good, info-classification banner"),
    ("S9", "scenario", "CP18", "upfront-fee confirmation; SSI as PDF attachment; asks for trade ref + UTI"),
    ("S10", "scenario", "CP11", "system-generated settlement summary; recipient-fill response columns (Reply-All)"),
    ("S11", "scenario", "CP11", "equity total-return-swap settlement with component breakdown (equity/div/int/wht)"),
    ("S12", "scenario", "CP19", "high-value EQ pre-confirmation; signed PAY/RECV table + detailed multi-entity SSI blocks"),
    ("SSI", "other", "CP1-10", "plain-text SSI verification / update note (not a cash-flow affirmation)"),
    ("Noise", "other", "-", "non-counterparty mail (out of scope)"),
]
for r in FORMAT_CATALOG:
    wc.append(list(r))
style_header(wc, 4)
for i, w in enumerate([18, 12, 12, 60], 1):
    wc.column_dimensions[col_letter(i)].width = w
 
# summary sheet
sm = wb.create_sheet("summary")
def cnt(p): return sum(1 for r in ground_truth if p(r))
aff_rows = [r for r in ground_truth if r["email_type"] == "SettlementAffirmation"]
oc = Counter(r["internal_match_status"] for r in aff_rows)
CFMAP = {c["cashflow_id"]: c for c in cashflows}
_refm = [r for r in aff_rows if r["internal_match_status"] != "Unmatched"]
_refsame = sum(1 for r in _refm if str(r["counterparty_reference"]) == str(CFMAP[r["matched_cashflow_id"]]["counterparty_trade_ref"]))
srows = [("Sr No", "Metric", "Value"),
         ("1", "Total emails", len(emails)),
         ("1.1", "Counterparty affirmations", len(aff)),
         ("1.1.1", "generic (CP1-10)", sum(1 for e in aff if e["email_format"] in GENERIC_FORMATS)),
         ("1.1.2", "scenario (CP11-17, S1-S8)", sum(1 for e in aff if e["email_format"] in SCENARIO)),
         ("1.2", "SSI verifications", len(ssi)),
         ("1.3", "Noise", sum(1 for e in emails if e["email_type"] == "Noise")),
         ("", "", ""),
         ("2", "Total cash flows (ground-truth rows)", len(aff_rows)),
         ("2.1", "Agreed", oc["Agreed"]), ("2.2", "AmountMismatch", oc["AmountMismatch"]),
         ("2.3", "SSIMissing", oc["SSIMissing"]), ("2.4", "Unmatched", oc["Unmatched"]),
         ("2.5", "AlreadySettled", oc["AlreadySettled"]),
         ("2.6", "of matched: cpty ref == internal ref", f"{_refsame} / {len(_refm)}"),
         ("3", "Allege cash flows (is_allege=Y)", cnt(lambda r: r["is_allege"] == "Y")),
         ("4", "Internal cash-flow records", len(cashflows)),
         ("4.1", "status Settled", sum(1 for c in cashflows if c["status"] == "Settled")),
         ("5", "Total ground-truth rows", len(ground_truth))]
for row in srows:
    depth = row[0].count(".") if row[0] and row[0] != "Sr No" else 0
    sm.append([row[0], ("    " * depth) + str(row[1]), row[2]])
    if row[0] and "." not in row[0] and row[0] != "Sr No":
        for c in (1, 2, 3):
            sm.cell(row=sm.max_row, column=c).font = Font(bold=True)
style_header(sm, 3)
sm.column_dimensions["A"].width = 10; sm.column_dimensions["B"].width = 40; sm.column_dimensions["C"].width = 12
wb.save(os.path.join(ROOT, "email_ground_truth.xlsx"))
 
# internal_bookings.xlsx (cashflows + product_systems)
wb2 = Workbook(); ws2 = wb2.active; ws2.title = "cashflows"
CFCOLS = ["cashflow_id", "booking_ref", "counterparty_org_name", "counterparty_entity", "product_type", "trade_system",
          "cashflow_type", "currency", "amount", "direction", "value_date", "status", "bank_trade_ref",
          "counterparty_trade_ref", "product_ref", "ssi_bank_name", "ssi_bank_bic", "ssi_account_number",
          "ssi_beneficiary_name", "ssi_beneficiary_bic", "ssi_intermediary_bank"]
ws2.append(CFCOLS)
SFILL = PatternFill("solid", fgColor="FFF2CC")
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
style_header(wp, 2); wp.column_dimensions["A"].width = 24; wp.column_dimensions["B"].width = 14
wb2.save(os.path.join(ROOT, "internal_bookings.xlsx"))
 
# counterparty_directory.xlsx
wb3 = Workbook(); ws3 = wb3.active; ws3.title = "counterparty_directory"
DCOLS = ["email_address", "counterparty_org_name", "counterparty_entity", "product_scope", "dl_role", "notes"]
ws3.append(DCOLS)
for r in DIRECTORY_ROWS:
    ws3.append(list(r))
style_header(ws3, len(DCOLS))
for i, w in enumerate([40, 20, 30, 24, 10, 24], 1):
    ws3.column_dimensions[col_letter(i)].width = w
wb3.save(os.path.join(ROOT, "counterparty_directory.xlsx"))
 
print(f"Emails: {len(emails)}  (affirmations {len(aff)} / SSI {len(ssi)} / noise {sum(1 for e in emails if e['email_type']=='Noise')})")
print(f"  generic {sum(1 for e in aff if e['email_format'] in GENERIC_FORMATS)} / scenario {sum(1 for e in aff if e['email_format'] in SCENARIO)}")
print(f"Cash-flow rows: {len(aff_rows)}  | outcomes: {dict(oc)}")
print(f"Internal cash flows: {len(cashflows)} (settled {sum(1 for c in cashflows if c['status']=='Settled')})")
print(f"Total ground-truth rows: {len(ground_truth)}")
