#!/usr/bin/env python3
"""
new_scenarios.py  —  generates 5 .eml files per new scenario (S13-S18).
Drop-in extension of sadhya.py logic; produces EML only, no ground-truth xlsx.

Masking rules applied throughout:
  - No real counterparty names in email body or domain
  - No CP abbreviations in domain
  - Counterparty names replaced with fictional coined brands (web-safe .example TLD)
  - Bank (recipient) is always "the bank" / BANK_MAILBOX
"""

import io
import os
import random
import shutil
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

random.seed(20260901)

# ── output folder ──────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.join(ROOT, "folder_new_scenarios")
if os.path.isdir(OUT):
    shutil.rmtree(OUT)
os.makedirs(OUT)

# ── bank (recipient) — unchanged from sadhya.py ───────────────────────────────
BANK_TO       = "Settlements <bank@settlements.com>"
BANK_MAILBOX  = "bank@settlements.com"

# ── date helpers ──────────────────────────────────────────────────────────────
START = datetime(2026, 6, 1, 8, 0, 0)

def rand_dt(span=24):
    return START + timedelta(
        days=random.randint(0, span),
        hours=random.randint(0, 9),
        minutes=random.randint(0, 59),
        seconds=random.randint(0, 59),
    )

def fmt_date(d, style="iso"):
    if isinstance(d, datetime):
        d = d.date()
    return {
        "iso":   d.isoformat(),
        "dmy":   d.strftime("%d/%m/%Y"),
        "dmon":  d.strftime("%d-%b-%Y"),
        "dmony": d.strftime("%d %b %Y"),
        # Windows-safe M/D/YYYY — no %-m or %-d
        "slash": f"{d.month}/{d.day}/{d.year}",
    }.get(style, d.isoformat())

def dnum(n):
    return f"{n:,.2f}"

# ── fictional counterparty pool (no real names, no CP abbreviations in domain) ─
#    coined nonsense tokens on RFC-reserved .example TLD
FICTIONAL_CPS = [
    # (display_name, sender_domain, sender_local, entity_name, bic)
    ("Veltrix Markets",      "veltrix.settlements.example",   "derivatives.settlements",
     "Veltrix Markets AG",                                    "VLTXDE2L"),
    ("Orvane Finance",       "orvane.ops.example",            "otc.settlements",
     "Orvane Finance N.V.",                                   "ORVNNL2L"),
    ("Quendal Securities",   "quendal.ops.example",           "otc.settlements",
     "Quendal Securities N.V.",                               "QNDLNL2L"),
    ("Stroven Bank",         "stroven.markets.example",       "ird.settlements",
     "Stroven Bank PLC",                                      "STRVGB2L"),
    ("Braxen Global",        "braxen.global.example",         "paris.settlements",
     "Braxen Global S.A.",                                    "BRXNFRPP"),
    ("Kesvin Capital",       "kesvin.capital.example",        "settlement.ops",
     "Kesvin Capital Ltd",                                    "KSVNGB2L"),
]

# ── currencies & SSI agent pool ───────────────────────────────────────────────
CCYS = ["USD", "EUR", "GBP", "JPY", "CHF", "AUD", "SGD", "HKD", "CNH"]

SSI_AGENTS = [
    ("AGENT BANK 1",  "AGABUS33", "AGABUS33XXX"),
    ("AGENT BANK 2",  "AGBBGB2L", "AGBBGB2LXXX"),
    ("AGENT BANK 3",  "AGCBDEFF", "AGCBDEFFXXX"),
    ("AGENT BANK 4",  "AGDBFRPP", "AGDBFRPPXXX"),
    ("AGENT BANK 5",  "AGEBGB22", "AGEBGB22XXX"),
    ("AGENT BANK 6",  "AGFBSG22", "AGFBSG22XXX"),
    ("AGENT BANK 7",  "AGGBUS3N", "AGGBUS3NXXX"),
    ("AGENT BANK 8",  "AGHBGB21", "AGHBGB21XXX"),
]

INTERMEDIARIES = [
    ("INTERMEDIARY BANK 1", "INTAUS3N"),
    ("INTERMEDIARY BANK 2", "INTBCHZ8"),
    ("INTERMEDIARY BANK 3", "INTCFRPP"),
    ("INTERMEDIARY BANK 4", "INTDJPJT"),
]

# ── product pools per scenario ────────────────────────────────────────────────
S13_PRODUCTS = ["Cross-Currency Swap", "Interest Rate Swap"]
S14_PRODUCTS = ["Interest Rate Swap"]
S15_PRODUCTS = ["Interest Rate Swap", "Note/MTN"]
S16_PRODUCTS = ["Basis Swap", "Interest Rate Swap"]
S17_PRODUCTS = ["Equity Option"]
S18_PRODUCTS = ["Equity Swap"]

# ── DB entity pool (S13) ──────────────────────────────────────────────────────
DB_ENTITIES = ["DB_LN", "DB_AG", "DB_NY", "DB_TK", "DB_SG"]

# ── equity underlyings pool (S18) ─────────────────────────────────────────────
EQ_UNDERLYINGS = [
    ("ROLLS-ROYCE HOLDINGS PLC",   "RR/ LN",  14.402),
    ("HSBC HOLDINGS PLC",          "HSBA LN", 15.618),
    ("BRITISH AMERICAN TOBACCO PLC","BATS LN", 45.8),
    ("VODAFONE GROUP PLC",         "VOD LN",   1.1805),
    ("ASTRAZENECA PLC",            "AZN LN",  127.26),
    ("BAE SYSTEMS PLC",            "BA/ LN",   19.82),
    ("ANGLO AMERICAN PLC",         "AAL LN",   37.86),
    ("UNILEVER PLC",               "ULVR LN",  46.24),
    ("SHELL PLC",                  "SHEL LN",  28.45),
    ("BP PLC",                     "BP/ LN",    4.72),
]

# ── xlsx style helpers ────────────────────────────────────────────────────────
HDR_FILL  = PatternFill("solid", fgColor="1F3A5F")
HDR_FONT  = Font(color="FFFFFF", bold=True, size=10)
BODY_FONT = Font(size=10)
THIN      = Side(style="thin", color="CCCCCC")
THIN_BDR  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_header_row(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 28

def auto_width(ws, min_w=10, max_w=40):
    for col in ws.columns:
        length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(
            max(length + 2, min_w), max_w
        )

def write_xlsx_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  S13 — DB-style pre-confirmation
#  Email:  external-caution banner + prose + SSI-change notice
#  Attach: multi-sheet XLSX (Cashflows | Your SSI | DB SSI | DB Contacts)
# ══════════════════════════════════════════════════════════════════════════════

def _s13_cashflow_rows(cp, vdate, n_rows):
    """Generate n_rows of synthetic cashflow data matching DB Cashflows sheet."""
    products = S13_PRODUCTS
    cf_types  = {
        "Cross-Currency Swap": ["FLOATING_RATE_RE", "FIXED_RATE_RETUR", "MULTIPLE", "PRINCIPAL"],
        "Interest Rate Swap":  ["FLOATING_RATE_RE", "FIXED_RATE_RETUR"],
    }
    prod_codes = {
        "Cross-Currency Swap": "IRXCcySwapFixFlt",
        "Interest Rate Swap":  "IRSwapFixFlt",
    }
    sys_codes = {
        "Cross-Currency Swap": "XCY_SWAP",
        "Interest Rate Swap":  "IR Swap",
    }
    rows = []
    for _ in range(n_rows):
        prod    = random.choice(products)
        ccy     = random.choice(["USD", "CNH", "JPY", "AUD", "EUR"])
        amount  = round(random.uniform(-2_000_000, 2_000_000), 2)
        trade_id = (
            random.choice(["AY", "AP", "AW", "Z", "SL", "TY"])
            + str(random.randint(100000, 999999))
            + random.choice(["M", "L"])
        )
        netting_id = str(random.randint(1_100_000, 1_200_000))
        mw_id      = str(random.randint(1_100_000, 1_200_000))
        trade_date = (vdate - timedelta(days=random.randint(2, 10))).strftime("%d/%m/%Y")
        db_entity  = random.choice(DB_ENTITIES)
        rows.append([
            trade_id,
            "Settlement Scheduled",
            "Please advise if any discrepancy",
            prod_codes[prod],
            netting_id,
            mw_id,
            vdate.strftime("%d/%m/%Y"),
            ccy,
            amount,
            cp["entity"],          # masked entity name
            db_entity,
            sys_codes[prod],
            random.choice(cf_types[prod]),
            0, 0,
            "STP",
            trade_date,
            "y",
        ])
    return rows

def _s13_ssi_rows_client(cp, ccys):
    """Your SSI (client's payment instructions to DB)."""
    rows = []
    for ccy in ccys:
        agent = random.choice(SSI_AGENTS)
        inter = random.choice(INTERMEDIARIES)
        acct  = str(random.randint(10**9, 10**10 - 1))
        rows.append([
            ccy,
            agent[1],          # Account with Bank Swift
            cp["bic"],         # Bene Swift
            agent[0],          # Account with Bank Name
            cp["entity"],      # Bene Name
            cp["entity"],      # CounterpartyId
            acct,              # Bene Account
            inter[1],          # Intermediary Swift
            inter[0],          # Intermediary Name
            "",                # Account with Bank Account
            "",                # Bank to Bank Info
        ])
    return rows

def _s13_ssi_rows_db(ccys):
    """DB SSI (DB's payment instructions)."""
    db_agents = [
        ("DB AG New York",    "DEUTUS33",   "DEUTGB2L", "DB AG London"),
        ("DB AG Hong Kong",   "DEUTHKHH",   "DEUTGB2L", "DB AG London"),
        ("DB AG Tokyo",       "DEUTJPJT",   "DEUTGB2L", "DB AG London"),
        ("DB AG Frankfurt",   "DEUTDEFF",   "DEUTDEFFSIP", "DB AG Frankfurt"),
        ("DB AG Sydney",      "DEUTAU2S",   "DEUTDEFFSIP", "DB AG Frankfurt"),
    ]
    rows = []
    for ccy in ccys:
        ag = random.choice(db_agents)
        acct = str(random.randint(10**7, 10**8 - 1))
        rows.append([
            ccy,
            ag[0],   # Account with Bank Name
            ag[1],   # Account with Bank Swift
            ag[3],   # Bene Name
            ag[2],   # Bene Swift
            "",      # CounterpartyId placeholder
            acct,    # Bene Account
            "", "", "", "",
        ])
    return rows

def build_s13_xlsx(cp, vdate, n_rows=8):
    wb = Workbook()

    # ── Sheet 1: Credit & Rates Cashflows ────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Credit & Rates Cashflows"

    title_font = Font(bold=True, size=12, color="1F3A5F")
    ws1["A1"] = "Derivative Settlements - Cashflow Confirmation"
    ws1["A1"].font = title_font
    ws1["A2"] = "Cashflows"
    ws1["A2"].font = Font(bold=True, size=11)
    ws1.merge_cells("A1:R1")
    ws1.merge_cells("A2:R2")

    cf_headers = [
        "Trade ID", "Cash Flow Status", "Response Required",
        "Product Features", "Netting Id", "MarkitWire Id",
        "Value Date", "Currency", "Amount",
        "Counterparty Legal Name", "Legal Entity",
        "Product Type", "Cashflow Type",
        "Notional Currency", "Original Notional Amount",
        "Rate", "Trade Date", "Status",
    ]
    ws1.append(cf_headers)
    style_header_row(ws1, len(cf_headers), row=ws1.max_row)

    cf_rows = _s13_cashflow_rows(cp, vdate, n_rows)
    for r in cf_rows:
        ws1.append(r)
        for c in range(1, len(cf_headers) + 1):
            ws1.cell(row=ws1.max_row, column=c).font = BODY_FONT
            ws1.cell(row=ws1.max_row, column=c).border = THIN_BDR

    ws1.freeze_panes = "A4"
    auto_width(ws1)

    # ── Sheet 2: Your SSI ─────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Settlements Instructions")
    ws2["A1"] = "Derivative Settlements - Cashflow Confirmation"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:K1")
    ws2["A2"] = "Settlements Instructions"
    ws2["A2"].font = Font(bold=True, size=11)
    ws2.merge_cells("A2:K2")

    ssi_headers = [
        "Currency", "Account with Bank Swift", "Bene Swift",
        "Account with Bank Name", "Bene Name", "CounterpartyId",
        "Bene Account", "Intermediary Swift", "Intermediary Name",
        "Account with Bank Account", "Bank to Bank Info",
    ]
    ws2.append(ssi_headers)
    style_header_row(ws2, len(ssi_headers), row=ws2.max_row)

    used_ccys = list({r[7] for r in cf_rows})
    for r in _s13_ssi_rows_client(cp, used_ccys):
        ws2.append(r)
        for c in range(1, len(ssi_headers) + 1):
            ws2.cell(row=ws2.max_row, column=c).font = BODY_FONT
            ws2.cell(row=ws2.max_row, column=c).border = THIN_BDR
    ws2.freeze_panes = "A4"
    auto_width(ws2)

    # ── Sheet 3: DB SSI ───────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Settlements Instructions_DB")
    ws3["A1"] = "Derivative Settlements - Cashflow Confirmation"
    ws3["A1"].font = title_font
    ws3.merge_cells("A1:K1")
    ws3["A2"] = "Settlements Instructions"
    ws3["A2"].font = Font(bold=True, size=11)
    ws3.merge_cells("A2:K2")
    ws3.append(ssi_headers)
    style_header_row(ws3, len(ssi_headers), row=ws3.max_row)
    for r in _s13_ssi_rows_db(used_ccys):
        ws3.append(r)
        for c in range(1, len(ssi_headers) + 1):
            ws3.cell(row=ws3.max_row, column=c).font = BODY_FONT
            ws3.cell(row=ws3.max_row, column=c).border = THIN_BDR
    ws3.freeze_panes = "A4"
    auto_width(ws3)

    # ── Sheet 4: DB Contacts ──────────────────────────────────────────────────
    ws4 = wb.create_sheet("DB Contacts")
    ws4["A1"] = "Derivative Settlements - Cashflow Confirmation"
    ws4["A1"].font = title_font
    ws4.merge_cells("A1:D1")
    ws4.append(["Name", "Mail", "Telephone", "Group Inbox"])
    style_header_row(ws4, 4, row=ws4.max_row)
    ws4.append(["", "settlements.connect@" + cp["domain"], "44-207-338-" + str(random.randint(1000, 9999)), ""])
    auto_width(ws4)

    return write_xlsx_bytes(wb)


def s13_email(cp, vdate):
    """S13 email body (text + html) — DB pre-confirmation style."""
    vd_str = vdate.strftime("%d %B %Y")
    req_num = random.randint(500000, 599999)
    banner = (
        "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n"
        "Do not click any links, open attachments or reply if you do not believe "
        "that the email is legitimate. If in doubt, report using the report email "
        "button or email to SPAM SUBMISSIONS.\n\n"
    )
    body_text = (
        f"{banner}"
        "Dear Client,\n\n"
        "Please be advised of below upcoming settlement amount(s) and settlement "
        "instructions. If there is a query with the cash flow(s) please advise and "
        "we will investigate and respond accordingly.\n\n"
        "Please refer to all the tabs in the attachment.\n\n\n\n"
        "Our Payments will be made to existing SSI details unless specified otherwise.\n\n"
        "PS: Contents\n\n"
        "1.  First Sheet : cashflow details\n"
        "2.  Second sheet : Your SSI\n"
        "3.  Third sheet : Our SSI\n"
        "4.  Contacts\n\n\n"
        "Kind Regards\n"
        "GBS Derivative Settlements\n"
        f"settlements.connect@{cp['domain']}\n"
    )
    html_body = f"""<html><body>
<pre style='font-family:Arial;font-size:11px;color:#c00'>{banner}</pre>
<p style='font-family:Arial;font-size:12px'>Dear Client,</p>
<p style='font-family:Arial;font-size:12px'>
Please be advised of below upcoming settlement amount(s) and settlement instructions.
If there is a query with the cash flow(s) please advise and we will investigate and respond accordingly.
</p>
<p style='font-family:Arial;font-size:12px'>Please refer to all the tabs in the attachment.</p>
<p style='font-family:Arial;font-size:12px'>
Our Payments will be made to existing SSI details unless specified otherwise.
</p>
<p style='font-family:Arial;font-size:12px'><b>PS: Contents</b><br>
1. First Sheet : cashflow details<br>
2. Second sheet : Your SSI<br>
3. Third sheet : Our SSI<br>
4. Contacts</p>
<p style='font-family:Arial;font-size:12px'>
Kind Regards<br>
GBS Derivative Settlements<br>
settlements.connect@{cp['domain']}
</p>
</body></html>"""
    subject = (
        f"Derivative Settlements Pre-Confirmation VD - {vd_str} - "
        f"{vd_str} Request#{req_num} - Auto"
    )
    return subject, body_text, html_body


# ══════════════════════════════════════════════════════════════════════════════
#  S14 — NEF/NGFP payment confirmation
#  Email:  short prose "please confirm agreement with attached CF"
#  Attach: Confirmation sheet (net summary + breakdown rows)
# ══════════════════════════════════════════════════════════════════════════════

def _s14_breakdown_rows(entity_code, vdate, settlement_no, n_legs):
    """Generate breakdown rows for the BreakDown_NN sheet."""
    rows = []
    total = 0
    for i in range(1, n_legs + 1):
        uss_id   = str(random.randint(50000, 70000))
        trade_ref = str(random.randint(5_000_000, 6_000_000)) + "_GMSW"
        amount   = round(random.uniform(-500_000, 1_500_000), 0)
        total   += amount
        rows.append([
            1, "N", entity_code, str(vdate), settlement_no,
            uss_id, "Swap - FI", "Swap", trade_ref, "JPY", amount,
        ])
    return rows, int(total)

def build_s14_xlsx(cp, vdate, entity_code, entity_full, n_legs=4):
    """Build Confirmation sheet XLSX matching NEF/NGFP style."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Confirmation"

    # ── header block ──────────────────────────────────────────────────────────
    title_font = Font(bold=True, size=12)
    ws["A1"] = "Settlement Confirmation :"
    ws["D1"] = entity_full
    ws["A1"].font = title_font
    ws["D1"].font = title_font

    gen_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    ws["A2"] = "Generated"
    ws["B2"] = gen_ts
    ws["C2"] = "Value Date"
    ws["D2"] = str(vdate)

    # ── net summary row ───────────────────────────────────────────────────────
    settlement_no = f"{entity_code}_{vdate.strftime('%Y%m%d')}_000{random.randint(1,9)}"
    agent = random.choice(SSI_AGENTS)
    inter = random.choice(INTERMEDIARIES)
    bene_acct = str(random.randint(100_000_000, 999_999_999))

    breakdown_rows, net_total = _s14_breakdown_rows(entity_code, vdate, settlement_no, n_legs)

    summary_headers = [
        "No.", "CCY", "Settlement Amount (Net)", "Pay from", "To",
        "Beneficiary BIC Code", "Beneficiary name", "Beneficiary a/c number",
        "Beneficiary Bank BIC Code", "Beneficiary Bank name",
        "Beneficiary Bank a/c number", "Intermediate Bank BIC code",
        "Intermediate Bank name", "Product-Event",
    ]
    ws.append(summary_headers)
    style_header_row(ws, len(summary_headers), row=ws.max_row)

    ws.append([
        1, "JPY", net_total,
        entity_code, cp["code"],
        cp["bic"],
        cp["entity"],
        bene_acct,
        inter[1],
        inter[0],
        "",
        "",
        "",
        "Swap - FI",
    ])
    for c in range(1, len(summary_headers) + 1):
        ws.cell(row=ws.max_row, column=c).font = BODY_FONT
        ws.cell(row=ws.max_row, column=c).border = THIN_BDR

    # ── notice text ───────────────────────────────────────────────────────────
    ws.append([])
    ws.append([f"Please be advised that {cp['entity']} will arrange the instruction(s) in accordance with above."])
    ws.append(["For the details, please see the sheet(s) labeled \"BreakDown_NN\""])
    ws.append(["Kindly confirm your agreement to the above by return e-mail to us."])
    ws.append(["Regards"])
    ws.append([cp["entity"]])
    ws.append([])

    # ── breakdown rows ────────────────────────────────────────────────────────
    bd_headers = [
        "No.", "Affirmed Flag", "Counterparty", "Value Date", "Settlement No",
        "USS", "Isin-Code", "Product-Event", "Trade Reference", "CCY", "Gross Amount",
    ]
    ws.append(bd_headers)
    style_header_row(ws, len(bd_headers), row=ws.max_row)
    for r in breakdown_rows:
        ws.append(r)
        for c in range(1, len(bd_headers) + 1):
            ws.cell(row=ws.max_row, column=c).font = BODY_FONT
            ws.cell(row=ws.max_row, column=c).border = THIN_BDR

    ws.append([])
    ws.append([net_total])

    ws.freeze_panes = "A5"
    auto_width(ws)
    return write_xlsx_bytes(wb), settlement_no, net_total


def s14_email(cp, vdate, settlement_no):
    """S14/S15 email body — short NEF-style prose."""
    vd_str = vdate.strftime("%Y%m%d")
    subject = f"Payment Confirmation for value date {vd_str}"
    body = (
        "Dear all,\n\n"
        "Please confirm your agreement with attached CF with title value date.\n\n"
        "Regards,\n"
        f"{cp['entity']}\n"
        "Operations Dept\n"
        f"Group e-mail address: settlement@{cp['domain']}\n"
    )
    return subject, body, None   # plain text only


# ══════════════════════════════════════════════════════════════════════════════
#  S15 — NEF/NIP arrangement fees  (same email template as S14, different entity)
#  Reuses build_s14_xlsx with entity_code="NIP1", entity_full="COUNTERPARTY INTL PLC"
# ══════════════════════════════════════════════════════════════════════════════
# (no separate renderer needed — s14_email + build_s14_xlsx cover it)


# ══════════════════════════════════════════════════════════════════════════════
#  S16 — NWM IRD daily spreadsheet
#  Email:  external caution + system-generated disclaimer + legal footer
#  Attach: CashFlowConfirmations sheet with STP/Non-STP column
# ══════════════════════════════════════════════════════════════════════════════

RESPONSE_ACTIONS = [
    "Agree",
    "Trade Date Mismatch",
    "Disagree Settlement date Mismatch",
    "DK ISIN Mismatch",
    "Quantity Mismatch",
    "Cash Mismatch",
    "Price Mismatch",
    "Direction Mismatch",
    "CCY Mismatch",
    "SSI Mismatch",
    "DK Trade",
    "Multiple Mismatch",
]

def _s16_cf_rows(cp, vdate, n_rows):
    rows = []
    for _ in range(n_rows):
        ccy    = random.choice(# ══════════════════════════════════════════════════════════════════════════════
#  (continuing _s16_cf_rows)
# ══════════════════════════════════════════════════════════════════════════════

        ["AUD", "GBP", "USD", "JPY", "EUR", "CHF"][random.randint(0, 5)]
    )
        amount = round(random.uniform(-5_000_000, 5_000_000), 2)
        trade_ref = f"HK0{random.randint(10**11, 10**12-1)}_REP{random.randint(1,3)}"
        settle_id = str(random.randint(300_000_000, 400_000_000))
        stp_flag  = random.choice(["STP", "STP", "STP", "Non STP"])
        direction = random.choice(["Pay", "Receive"])
        product   = random.choice(S16_PRODUCTS)
        agent     = random.choice(SSI_AGENTS)
        cpty_instr = (
            f"[ Creditor Bic:- ({cp['bic']}) | "
            f"Creditor Agent Bank:- ({agent[1]}) | "
            f"Account Name:- ({cp['entity']}) | "
            f"Account Number:- ({str(random.randint(10**12, 10**13-1))}) ]"
        )
        rows.append([
            "",                  # Response Action (blank — recipient fills)
            "",                  # CPTY Comments
            cp["entity"],        # Counter Party Name  (masked)
            "SETTLEMENTS PLC",   # NWM Entity  (generic)
            trade_ref,
            settle_id,
            stp_flag,
            ccy,
            amount,
            vdate.strftime("%m/%d/%Y"),
            direction,
            "",                  # Swap wire Id
            "",                  # Structure Id
            product,
            "",                  # NWM Payment Instructions
            cpty_instr,          # Customer Payment Instructions
        ])
    return rows


def build_s16_xlsx(cp, vdate, n_rows=6):
    wb = Workbook()
    ws = wb.active
    ws.title = "CashFlowConfirmations"

    title_font = Font(bold=True, size=12, color="1F3A5F")
    ws["A1"] = "Cash flow Confirmation"
    ws["A1"].font = title_font
    ws.merge_cells("A1:P1")

    headers = [
        "Response Action", "CPTY Comments", "Counter Party Name", "NWM Entity",
        "Trade Ref", "Settlement Id", "STP/Non STP", "CCY", "Amount",
        "Settlement Date", "Direction", "Swap wire Id", "Structure Id",
        "Product Type", "NWM Payment Instructions", "Customer Payment Instructions",
    ]
    ws.append(headers)
    style_header_row(ws, len(headers), row=ws.max_row)

    cf_rows = _s16_cf_rows(cp, vdate, n_rows)
    for r in cf_rows:
        ws.append(r)
        for c in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row, column=c).font = BODY_FONT
            ws.cell(row=ws.max_row, column=c).border = THIN_BDR

    # ── response action legend block ──────────────────────────────────────────
    ws.append([])
    ws.append(["Valid Response Actions:"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    for action in RESPONSE_ACTIONS:
        ws.append([action])

    ws.freeze_panes = "A3"
    auto_width(ws)
    return write_xlsx_bytes(wb), cf_rows


def s16_email(cp, vdate, cf_rows):
    """S16 email body — NWM IRD system-generated style."""
    vd_str   = vdate.strftime("%d %b %Y")
    ref_ids  = ",".join(str(random.randint(1000, 9999)) for _ in range(random.randint(3, 6)))
    date_str = vdate.strftime("%d %b %Y")
    send_date = (vdate - timedelta(days=1)).strftime("%d%m%Y")
    subject  = (
        f"Settlement Confirmation - {vd_str} - {ref_ids} - "
        f"{date_str} to {date_str} - {send_date}"
    )
    banner = (
        "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n"
        "Do not click any links, open attachments or reply if you do not believe "
        "that the email is legitimate. If in doubt, report using the report email "
        "button or email to SPAM SUBMISSIONS.\n\n"
    )
    body_text = (
        f"{banner}"
        "Hi Team,\n\n"
        "The enclosed spreadsheet contains cash flows\n\n"
        "Please provide your confirmation on this email for all Non STP cashflows "
        "as per column \"STP/Non STP\" by replying to all and attach the updated spreadsheet.\n\n"
        "In case of Pay, We will pay on the SSIs as mentioned against the cash flow\n\n"
        "This is a system generated email. In case of discrepancy please mail your "
        f"calculations to settlements@{cp['domain']} or call hotline: "
        f"+{random.randint(10,99)} {random.randint(100,999)} {random.randint(100,999)} "
        f"{random.randint(1000,9999)}\n\n"
        "Disclaimer\n"
        "Kindly respond to the emails as per expected format to provide a better service "
        "to our valuable clients. Also if the trade does not pertain to your area please "
        "forward the same to the relevant team for smooth settlement of the trade.\n"
    )
    html_body = f"""<html><body>
<pre style='font-family:Arial;font-size:11px;color:#c00'>{banner}</pre>
<p style='font-family:Arial;font-size:12px'>Hi Team,</p>
<p style='font-family:Arial;font-size:12px'>The enclosed spreadsheet contains cash flows</p>
<p style='font-family:Arial;font-size:12px'>
Please provide your confirmation on this email for all Non STP cashflows
as per column "STP/Non STP" by replying to all and attach the updated spreadsheet.
</p>
<p style='font-family:Arial;font-size:12px'>
In case of Pay, We will pay on the SSIs as mentioned against the cash flow
</p>
<p style='font-family:Arial;font-size:12px'>
This is a system generated email. In case of discrepancy please mail your calculations to
settlements@{cp['domain']}
</p>
</body></html>"""
    return subject, body_text, html_body


# ══════════════════════════════════════════════════════════════════════════════
#  S17 — BNPP settlement notice
#  Email:  external caution + "Agreed/Reject link" prose + legal disclaimer
#  Attach: "Details proposal for unit CF" sheet — SETTLEMENT NOTICE header
#          + flow rows + OUR/YOUR payment instructions block
# ══════════════════════════════════════════════════════════════════════════════

def _s17_flow_rows(cp, vdate, n_rows):
    rows = []
    for _ in range(n_rows):
        flow_ref   = str(random.randint(900_000_000, 999_999_999))
        strategy   = "OP" + str(random.randint(10**7, 10**8 - 1))
        caf_id     = str(random.randint(1_300_000_000, 1_400_000_000))
        amount     = round(random.uniform(-3_000_000, -500_000), 2)   # BNPP pays (negative)
        mw_ref     = str(random.randint(900_000_000, 999_999_999))
        cp_ref     = ""
        rows.append([
            flow_ref, strategy, caf_id, "PREM",
            cp["entity"],   # CounterParty (masked)
            "CPTYREF",      # Crds (generic)
            vdate.strftime("%d-%b-%Y"),
            amount, "JPY",
            mw_ref, cp_ref,
        ])
    return rows


def build_s17_xlsx(cp, vdate, n_rows=3):
    wb = Workbook()
    ws = wb.active
    ws.title = "Details proposal for unit CF"

    title_font = Font(bold=True, size=12)
    sn_ref = f"SN{random.randint(10**9, 10**10-1)}"
    ts_str = (vdate - timedelta(days=random.randint(1, 3))).strftime("%d-%b-%Y %H:%M:%S")

    ws["A1"] = f"SETTLEMENT NOTICE\nRef: {sn_ref}"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(wrap_text=True)
    ws["G1"] = ts_str
    ws["G1"].font = Font(size=10)
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 36

    # ── flow rows ─────────────────────────────────────────────────────────────
    flow_headers = [
        "FLOW REF", "STRATEGY", "MO CAF ID", "Cfw Typ Desc",
        "CounterParty", "Crds", "Theo Val Date", "Amount", "Ccy",
        "MW Ref", "Counterparty Ref",
    ]
    ws.append(flow_headers)
    style_header_row(ws, len(flow_headers), row=ws.max_row)

    flow_rows = _s17_flow_rows(cp, vdate, n_rows)
    for r in flow_rows:
        ws.append(r)
        for c in range(1, len(flow_headers) + 1):
            ws.cell(row=ws.max_row, column=c).font = BODY_FONT
            ws.cell(row=ws.max_row, column=c).border = THIN_BDR

    ws.append([])
    ws.append(["Each Amount will be settled independently "
               "(Negative amount is payment for sender / Positive amount sender is receiving)"])
    ws.cell(row=ws.max_row, column=1).font = Font(italic=True, size=9)
    ws.merge_cells(f"A{ws.max_row}:K{ws.max_row}")

    # ── OUR payment instructions ──────────────────────────────────────────────
    ws.append([])
    ws.append(["OUR PAYMENT INSTRUCTIONS"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=11)
    ws.merge_cells(f"A{ws.max_row}:K{ws.max_row}")

    our_agent = random.choice(SSI_AGENTS)
    our_acct  = str(random.randint(10**9, 10**10 - 1))
    ws.append(["F57 : CORRESPONDENT BANK", "", "", "", "F57 : BENEFICIARY CUSTOMER"])
    ws.append(["Name", our_agent[0], "", "", "Name", cp["entity"]])
    ws.append(["Account Number", "", "", "", "Account Number", our_acct])
    ws.append(["BIC SWIFT", our_agent[1], "", "", "BIC SWIFT", cp["bic"]])

    # ── YOUR payment instructions ─────────────────────────────────────────────
    ws.append([])
    ws.append(["YOUR PAYMENT INSTRUCTIONS"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=11)
    ws.merge_cells(f"A{ws.max_row}:K{ws.max_row}")

    inter      = random.choice(INTERMEDIARIES)
    your_agent = random.choice(SSI_AGENTS)
    your_acct  = str(random.randint(100_000_000, 999_999_999))
    ws.append(["F56: INTERMEDIARY BANK", "", "", "F57: CORRESPONDENT BANK", "", "", "", "F58: BENEFICIARY CUSTOMER"])
    ws.append(["Name", inter[0], "", "Name", your_agent[0], "", "", "Name", cp["entity"]])
    ws.append(["Account Number", "", "", "Account Number", "", "", "", "Account Number", your_acct])
    ws.append(["BIC SWIFT", inter[1], "", "BIC SWIFT", your_agent[1], "", "", "BIC SWIFT", cp["bic"]])

    # ── contacts ──────────────────────────────────────────────────────────────
    ws.append([])
    ws.append(["For any queries regarding settlements, please contact:"])
    ws.append(["GROUP EMAIL ADDRESS FOR ALL PRODUCTS"])
    ws.append([f"For Settlement Notices : settlements@{cp['domain']}"])

    auto_width(ws)
    return write_xlsx_bytes(wb), sn_ref, flow_rows


def s17_email(cp, vdate, sn_ref, flow_rows):
    """S17 email body — BNPP settlement notice style."""
    vd_str   = vdate.strftime("%d-%b-%Y")
    strategy = flow_rows[0][1] if flow_rows else "OP00000000"
    tech_ref = "".join(random.choices("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", k=10))
    subject  = (
        f"SettlementNotice ref {sn_ref} {strategy} CPTYREF JPY "
        f"{vd_str} - Tech-Ref: {tech_ref}"
    )
    banner = (
        "< CAUTION: THIS EMAIL HAS BEEN SENT FROM AN EXTERNAL SENDER >\n"
        "Do not click any links, open attachments or reply if you do not believe "
        "that the email is legitimate. If in doubt, report using the report email "
        "button or email to SPAM SUBMISSIONS.\n\n"
    )
    disclaimer_en = (
        "This message and any attachments (the \"message\") is intended solely for "
        "the addressees and is confidential. If you receive this message in error, "
        "please delete it and immediately notify the sender. Any use not in accord "
        "with its purpose, any dissemination or disclosure, either whole or partial, "
        "is prohibited except formal approval."
    )
    body_text = (
        f"{banner}"
        "Sir, Madam,\n\n"
        "Please find attached our invoice.\n\n"
        "Please kindly respond by clicking on the 'Agreed' or 'Reject' link, "
        "if 'Reject' please state reason.\n\n"
        f"Please do not modify subject line of this email.\n\n"
        "For any query you can contact us at the group email address below.\n\n"
        "Thanks in advance,\n\n"
        "Best regards,\n"
        "OTC EQD Settlement department\n\n"
        "***\n\n"
        f"{disclaimer_en}\n"
    )
    html_body = f"""<html><body>
<pre style='font-family:Arial;font-size:11px;color:#c00'>{banner}</pre>
<p style='font-family:Arial;font-size:12px'>Sir, Madam,</p>
<p style='font-family:Arial;font-size:12px'>Please find attached our invoice.</p>
<p style='font-family:Arial;font-size:12px'>
Please kindly respond by clicking on the 'Agreed' or 'Reject' link,
if 'Reject' please state reason.
</p>
<p style='font-family:Arial;font-size:12px'>
Please do not modify subject line of this email.
</p>
<p style='font-family:Arial;font-size:12px'>
Thanks in advance,<br>Best regards,<br>OTC EQD Settlement department
</p>
<hr/>
<p style='font-family:Arial;font-size:10px;color:#666'>{disclaimer_en}</p>
</body></html>"""
    return subject, body_text, html_body


# ══════════════════════════════════════════════════════════════════════════════
#  S18 — Scotia-style GBP equity swap
#  Email:  short inline table in body + "calculations attached"
#  Attach: Sheet1 — wide equity swap breakdown (interest notional + trading gain)
# ══════════════════════════════════════════════════════════════════════════════

def _s18_eq_rows(cp, vdate, swap_id, n_underlyings=3):
    """Generate equity swap breakdown rows per underlying."""
    rows = []
    underlyings = random.sample(EQ_UNDERLYINGS, min(n_underlyings, len(EQ_UNDERLYINGS)))
    fund_name   = f"{cp['entity']} {random.randint(1,4)} Wk Short"
    eff_start   = vdate - timedelta(days=random.randint(3, 10))

    for name, ticker, price in underlyings:
        qty      = random.choice([250_000, 500_000, 750_000, 1_000_000,
                                  1_500_000, 2_000_000, 3_000_000, 6_000_000])
        notional = round(qty * price, 0)
        borrow   = round(random.uniform(60, 120), 0)
        rate     = round(random.uniform(3.70, 3.75), 4)
        spread   = round(random.uniform(80, 120), 0)

        # interest notional rows (one per day in range)
        for day_offset in range(random.randint(1, 4)):
            fin_start = eff_start + timedelta(days=day_offset)
            days      = 1 if day_offset < 3 else 3
            int_amt   = round(-notional * rate / 100 / 365 * days, 2)
            rows.append([
                cp["entity"], fund_name, vdate.strftime("%d/%m/%y") + " GBP 1 Wk Short",
                cp["entity"],
                f"{name} AT LONDON - UNITED KINGDOM(EQUITY)",
                ticker,
                "Interest Notional",
                vdate.strftime("%Y-%m-%d 00:00:00"),
                "GBP", "Value-dated",
                int_amt,
                "BNS London",
                fin_start.strftime("%Y-%m-%d 00:00:00"),
                fin_start.strftime("%Y-%m-%d 00:00:00"),
                "Unchecked", 1, "No FX required",
                vdate.strftime("%Y-%m-%d 00:00:00"),
                vdate.strftime("%Y-%m-%d 00:00:00"),
                "ALL",
                eff_start.strftime("%Y-%m-%d 00:00:00"),
                days,
                -notional, -qty,
                spread, borrow, price, price,
                "GBP", swap_id,
                int_amt, -notional, -notional,
                "LNLF", "Short",
            ])

        # trading gain row
        new_price = round(price * random.uniform(0.95, 1.06), 3)
        tg_amt    = round((new_price - price) * qty, 2)
        new_notl  = round(new_price * qty, 0)
        rows.append([
            cp["entity"], fund_name, vdate.strftime("%d/%m/%y") + " GBP 1 Wk Short",
            cp["entity"],
            f"{name} AT LONDON - UNITED KINGDOM(EQUITY)",
            ticker,
            "Trading Gain",
            vdate.strftime("%Y-%m-%d 00:00:00"),
            "GBP", "Trade-dated",
            tg_amt,
            "BNS London",
            eff_start.strftime("%Y-%m-%d 00:00:00"),
            eff_start.strftime("%Y-%m-%d 00:00:00"),
            "Unchecked", 1, "No FX required",
            vdate.strftime("%Y-%m-%d 00:00:00"),
            vdate.strftime("%Y-%m-%d 00:00:00"),
            "ALL",
            eff_start.strftime("%Y-%m-%d 00:00:00"),
            1,
            new_notl, qty,
            price, new_price,
            "GBP", swap_id,
            tg_amt, notional, new_notl,
            "LNLF", "Short",
        ])
    return rows


def build_s18_xlsx(cp, vdate, swap_id, n_underlyings=4):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = [
        "Client name", "Swap", "Fund name", "Instrument Name", "BLOOMBERG",
        "Cashflow type", "Cashflow subtype", "Payment date", "Payment currency",
        "Position type", "Cashflow amount", "Entity name",
        "Effective start date", "Effective end date",
        "Fx required", "Fx rate", "Fx type",
        "Financing start date", "Financing end date", "Payment type",
        "Confirm date", "Days", "Notional", "Quantity",
        "Spread", "Borrow cost", "Interest rate", "Distribution percentage",
        "Distribution rate", "Withholding tax", "Distribution country code",
        "Cost Price", "Underlying currency", "Swap id",
        "Underlying amount", "Cost notional", "Price notional",
        "Hedge account", "Market fee rate", "Market fee description",
        "Comment", "Side",
    ]
    ws.append(headers)
    style_header_row(ws, len(headers), row=1)

    eq_rows = _s18_eq_rows(cp, vdate, swap_id, n_underlyings)
    for r in eq_rows:
        ws.append(r)
        for c in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row, column=c).font = BODY_FONT

    ws.freeze_panes = "A2"
    auto_width(ws, min_w=8, max_w=30)
    return write_xlsx_bytes(wb), eq_rows


def s18_email(cp, vdate, swap_id, eq_rows):
    """S18 email body — Scotia-style short inline table."""
    # Windows-safe date formatting (no %-m / %-d)
    vd_str  = f"{vdate.month}/{vdate.day}/{vdate.year}"
    net_amt = sum(r[10] for r in eq_rows)
    ccy     = "GBP"
    subject = f"{cp['entity']} / Please confirm {ccy} settlement / {vd_str}"

    inline_table = (
        f"{'Client name':<30} {'Payment date':<14} {'Payment currency':<18} "
        f"{'Entity name':<12} {'Swap id':<10} {'Cashflow amount':>18}\n"
        f"{cp['entity']:<30} {vd_str:<14} {ccy:<18} "
        f"{'SIDAC':<12} {swap_id:<10} {dnum(net_amt):>18}"
    )
    body_text = (
        "Hello,\n\n"
        f"We see to settle the following {ccy} flow for VD {vd_str}. "
        "Please review and share your SSIs.\n\n"
        f"{inline_table}\n\n"
        "Please find our calculations attached.\n\n"
        "Thank you.\n\n"
        "Regards,\n\n"
        f"Analyst | Derivative Settlement Operations\n"
        f"Tel: +{random.randint(1,9)} ({random.randint(100,999)})-"
        f"{random.randint(100,999)}-{random.randint(1000,9999)} "
        f"Email: settlement.ops@{cp['domain']}\n"
        "_______________________________________________\n"
        f"Global Wholesale Operations\n"
        f"Escalation: Level 1: Manager | Email: escalation1@{cp['domain']}\n"
        f"            Level 2: Senior Manager | Email: escalation2@{cp['domain']}\n"
    )
    html_body = f"""<html><body>
<p style='font-family:Arial;font-size:12px'>Hello,</p>
<p style='font-family:Arial;font-size:12px'>
We see to settle the following {ccy} flow for VD {vd_str}.
Please review and share your SSIs.
</p>
<table style='border-collapse:collapse;font-family:Arial;font-size:11px'>
<thead><tr>
{''.join(f"<th style='background:#1F3A5F;color:#fff;padding:4px 8px'>{h}</th>" for h in
 ['Client name','Payment date','Payment currency','Entity name','Swap id','Cashflow amount'])}
</tr></thead>
<tbody><tr>
<td style='padding:3px 8px;border:1px solid #ccc'>{cp['entity']}</td>
<td style='padding:3px 8px;border:1px solid #ccc'>{vd_str}</td>
<td style='padding:3px 8px;border:1px solid #ccc'>{ccy}</td>
<td style='padding:3px 8px;border:1px solid #ccc'>SIDAC</td>
<td style='padding:3px 8px;border:1px solid #ccc'>{swap_id}</td>
<td style='padding:3px 8px;border:1px solid #ccc'>{dnum(net_amt)}</td>
</tr></tbody>
</table>
<p style='font-family:Arial;font-size:12px'>Please find our calculations attached.</p>
<p style='font-family:Arial;font-size:12px'>Thank you.<br>Regards,<br>
Analyst | Derivative Settlement Operations<br>
settlement.ops@{cp['domain']}</p>
</body></html>"""
    return subject, body_text, html_body


# ══════════════════════════════════════════════════════════════════════════════
#  EML writer  (mirrors sadhya.py write_eml exactly)
# ══════════════════════════════════════════════════════════════════════════════

def write_eml(folder, email_id, sender_name, sender_email, sender_domain,
              subject, received_dt, text_body, html_body,
              attachments=None, email_type="SettlementAffirmation", fmt=""):
    """
    attachments: list of (filename, bytes_data)
    """
    msg = EmailMessage()
    msg["From"]       = f"{sender_name} <{sender_email}>"
    msg["To"]         = BANK_TO
    msg["Date"]       = format_datetime(received_dt)
    msg["Subject"]    = subject
    msg["Message-ID"] = make_msgid(domain=sender_domain)
    msg["X-Email-ID"] = email_id
    msg["X-Email-Type"]   = email_type
    msg["X-Email-Format"] = fmt

    msg.set_content(text_body, cte="8bit")

    if html_body:
        msg.add_alternative(html_body, subtype="html", cte="8bit")

    for fname, fdata in (attachments or []):
        msg.add_attachment(
            fdata,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=fname,
        )

    out_path = os.path.join(folder, f"{email_id}.eml")
    with open(out_path, "wb") as fh:
        fh.write(bytes(msg))
    return out_path


# ══════════════════════════════════════════════════════════════════════════════
#  Main generation loop — 5 EML per scenario (S13-S18)
# ══════════════════════════════════════════════════════════════════════════════

_idx = 0
def next_id(prefix="EML"):
    global _idx
    _idx += 1
    return f"{prefix}-{_idx:04d}"


SCENARIOS = {
    "S13": {
        "cp_idx": 0,          # Veltrix Markets
        "n_cf_range": (5, 20),
        "products": S13_PRODUCTS,
    },
    "S14": {
        "cp_idx": 1,          # Orvane Finance  (NGFP-style)
        "n_cf_range": (3, 8),
        "products": S14_PRODUCTS,
    },
    "S15": {
        "cp_idx": 2,          # Quendal Securities  (NIP-style)
        "n_cf_range": (1, 4),
        "products": S15_PRODUCTS,
    },
    "S16": {
        "cp_idx": 3,          # Stroven Bank
        "n_cf_range": (4, 12),
        "products": S16_PRODUCTS,
    },
    "S17": {
        "cp_idx": 4,          # Braxen Global
        "n_cf_range": (2, 5),
        "products": S17_PRODUCTS,
    },
    "S18": {
        "cp_idx": 5,          # Kesvin Capital
        "n_cf_range": (3, 6),
        "products": S18_PRODUCTS,
    },
}

generated = []

for scen_id, cfg in SCENARIOS.items():
    cp = FICTIONAL_CPS[cfg["cp_idx"]]
    # build a dict matching the shape used in renderers
    cp_dict = {
        "name":   cp[0],
        "domain": cp[1],
        "email":  f"{cp[2]}@{cp[1]}",
        "entity": cp[3],
        "bic":    cp[4],
        "code":   scen_id,
    }

    for i in range(5):
        eid      = next_id(scen_id)
        recv_dt  = rand_dt(span=30)
        vdate    = (recv_dt + timedelta(days=random.choice([1, 2, 2, 3]))).date()
        lo, hi   = cfg["n_cf_range"]
        n_rows   = random.randint(lo, hi)

        # ── per-scenario build ────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════════════
#  (continuing main generation loop from where Part B left off)
# ══════════════════════════════════════════════════════════════════════════════

        if scen_id == "S13":
            xlsx_bytes = build_s13_xlsx(cp_dict, vdate, n_rows=n_rows)
            subject, text_body, html_body = s13_email(cp_dict, vdate)
            vd_str = vdate.strftime("%d%m%Y")
            attach_name = f"Cashflows_{vd_str}_{eid}.xlsx"
            attachments = [(attach_name, xlsx_bytes)]

        elif scen_id == "S14":
            entity_code = "NGFP1"
            entity_full = cp_dict["entity"].upper()
            xlsx_bytes, settlement_no, net_total = build_s14_xlsx(
                cp_dict, vdate, entity_code=entity_code,
                entity_full=entity_full, n_legs=n_rows
            )
            subject, text_body, html_body = s14_email(cp_dict, vdate, settlement_no)
            attach_name = f"Settlement_Affirmation_{settlement_no}.xlsx"
            attachments = [(attach_name, xlsx_bytes)]

        elif scen_id == "S15":
            entity_code = "NIP1"
            entity_full = cp_dict["entity"].upper()
            xlsx_bytes, settlement_no, net_total = build_s14_xlsx(
                cp_dict, vdate, entity_code=entity_code,
                entity_full=entity_full, n_legs=n_rows
            )
            subject, text_body, html_body = s14_email(cp_dict, vdate, settlement_no)
            subject = subject.replace(
                "Payment Confirmation",
                "Arrangement Fees and swaps"
            )
            attach_name = f"Settlement_Affirmation_{settlement_no}.xlsx"
            attachments = [(attach_name, xlsx_bytes)]

        elif scen_id == "S16":
            xlsx_bytes, cf_rows = build_s16_xlsx(cp_dict, vdate, n_rows=n_rows)
            subject, text_body, html_body = s16_email(cp_dict, vdate, cf_rows)
            vd_str = vdate.strftime("%d%m%Y")
            attach_name = f"DailySpreadSheet_{vd_str}_{eid}.xlsx"
            attachments = [(attach_name, xlsx_bytes)]

        elif scen_id == "S17":
            xlsx_bytes, sn_ref, flow_rows = build_s17_xlsx(
                cp_dict, vdate, n_rows=n_rows
            )
            subject, text_body, html_body = s17_email(
                cp_dict, vdate, sn_ref, flow_rows
            )
            vd_str = vdate.strftime("%d-%b-%Y").replace(" ", "")
            attach_name = (
                f"BNPP_Payment_{flow_rows[0][1] if flow_rows else 'OP00000000'}"
                f"_CPTYREF_{vd_str}_{eid}.xlsx"
            )
            attachments = [(attach_name, xlsx_bytes)]

        elif scen_id == "S18":
            swap_id = str(random.randint(400_000, 600_000))
            n_und   = random.randint(3, 6)
            xlsx_bytes, eq_rows = build_s18_xlsx(
                cp_dict, vdate, swap_id=swap_id, n_underlyings=n_und
            )
            subject, text_body, html_body = s18_email(
                cp_dict, vdate, swap_id, eq_rows
            )
            vd_str = vdate.strftime("%Y%m%d")
            attach_name = f"{cp_dict['entity'].replace(' ','_')}_{vd_str}.xlsx"
            attachments = [(attach_name, xlsx_bytes)]

        else:
            continue

        # ── write .eml ────────────────────────────────────────────────────────
        out_path = write_eml(
            folder       = OUT,
            email_id     = eid,
            sender_name  = f"{cp_dict['name']} Settlements",
            sender_email = cp_dict["email"],
            sender_domain= cp_dict["domain"],
            subject      = subject,
            received_dt  = recv_dt,
            text_body    = text_body,
            html_body    = html_body,
            attachments  = attachments,
            email_type   = "SettlementAffirmation",
            fmt          = scen_id,
        )

        generated.append({
            "email_id":    eid,
            "scenario":    scen_id,
            "cp_name":     cp_dict["name"],
            "cp_domain":   cp_dict["domain"],
            "vdate":       str(vdate),
            "received_dt": str(recv_dt),
            "subject":     subject,
            "attachment":  attach_name,
            "eml_path":    out_path,
        })

        print(
            f"  [{scen_id}] {eid}  vdate={vdate}  "
            f"attach={attach_name}"
        )


# ══════════════════════════════════════════════════════════════════════════════
#  Summary
# ══════════════════════════════════════════════════════════════════════════════

from collections import Counter
scen_counts = Counter(r["scenario"] for r in generated)

print("\n" + "=" * 60)
print(f"Total EML files generated : {len(generated)}")
print(f"Output folder             : {OUT}")
print("-" * 60)
for scen_id in sorted(scen_counts):
    print(f"  {scen_id} : {scen_counts[scen_id]} emails")
print("=" * 60)



