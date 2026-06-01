"""
Finance Bot v3 — Multi-user with Google Sheets
Each user gets their own sheet tab named by their Telegram user ID.
Data is fully isolated between users.
"""

import os, re, csv, io, logging
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    XLSX_OK = True
except ImportError:
    XLSX_OK = False
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8698682076:AAGa2VWg3MN0IdJcQ64Rtuegg4Mt9GvvCYE")
WEBAPP_URL = os.getenv("WEBAPP_URL", "")
GOOGLE_CREDS_FILE = os.getenv("GOOGLE_CREDS_FILE", "google_creds.json")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Finance Bot Data")

EXPENSE_CATS = [
    "🛒 Groceries","🏠 Rent","🚌 Transport","📱 Subscriptions","🍽 Dining",
    "👕 Clothing","💊 Health","🎮 Entertainment","🏋 Sports","✈️ Travel",
    "📚 Education","🔧 Home","💰 Investments","📦 Other"
]
INCOME_CATS = ["💼 Salary","🖥 Freelance","📈 Dividends","🎁 Gift","🏠 Rental income","💡 Other income"]
HEADERS = ["Date","Amount","Category","Description","Type","Month","UserID"]

pending = {}

# ─── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
_gs_client = None
_spreadsheet = None

def get_spreadsheet():
    global _gs_client, _spreadsheet
    if _spreadsheet:
        return _spreadsheet
    try:
        import gspread, json
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
        # Try env variable first, then file
        creds_json = os.getenv("GOOGLE_CREDS_JSON")
        if creds_json:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
        elif os.path.exists(GOOGLE_CREDS_FILE):
            creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
        else:
            logger.warning("No Google credentials found")
            return None
        _gs_client = gspread.authorize(creds)
        _spreadsheet = _gs_client.open(SPREADSHEET_NAME)
        return _spreadsheet
    except Exception as e:
        logger.error(f"Google Sheets error: {e}")
        return None

def get_user_sheet(user_id: int):
    """Get or create a sheet tab for this user."""
    sp = get_spreadsheet()
    if not sp:
        return None
    sheet_name = f"user_{user_id}"
    try:
        sheet = sp.worksheet(sheet_name)
    except Exception:
        try:
            import gspread
            sheet = sp.add_worksheet(title=sheet_name, rows=5000, cols=8)
            sheet.append_row(HEADERS)
        except Exception as e:
            logger.error(f"Create sheet error: {e}")
            return None
    return sheet

def read_user_tx(user_id: int):
    sheet = get_user_sheet(user_id)
    if not sheet:
        return []
    try:
        return sheet.get_all_records()
    except Exception as e:
        logger.error(f"Read error: {e}")
        return []

def write_user_tx(user_id: int, date, amount, category, description, tx_type):
    sheet = get_user_sheet(user_id)
    if not sheet:
        logger.warning(f"No sheet for user {user_id}, using local fallback")
        _write_local(user_id, date, amount, category, description, tx_type)
        return
    try:
        sheet.append_row([
            date.strftime("%d.%m.%Y"), amount, category,
            description, tx_type, date.strftime("%Y-%m"), str(user_id)
        ])
    except Exception as e:
        logger.error(f"Write error: {e}")
        _write_local(user_id, date, amount, category, description, tx_type)

# ─── LOCAL FALLBACK (if no Google Sheets) ─────────────────────────────────────
def _local_path(user_id): return f"/tmp/{user_id}.csv"

def _ensure_local(user_id):
    p = _local_path(user_id)
    if not os.path.exists(p):
        with open(p,"w",newline="",encoding="utf-8") as f:
            csv.writer(f).writerow(HEADERS)

def _write_local(user_id, date, amount, category, description, tx_type):
    _ensure_local(user_id)
    with open(_local_path(user_id),"a",newline="",encoding="utf-8") as f:
        csv.writer(f).writerow([date.strftime("%d.%m.%Y"),amount,category,description,tx_type,date.strftime("%Y-%m"),str(user_id)])

def _read_local(user_id):
    _ensure_local(user_id)
    with open(_local_path(user_id),"r",encoding="utf-8") as f:
        return list(csv.DictReader(f))

def get_tx(user_id):
    """Read from Google Sheets or local fallback."""
    rows = read_user_tx(user_id)
    if not rows:
        rows = _read_local(user_id)
    return rows

# ─── STATS ─────────────────────────────────────────────────────────────────────
def get_month_stats(user_id, month):
    rows = get_tx(user_id)
    txs = [r for r in rows if r.get("Month") == month]
    exp = sum(float(r["Amount"]) for r in txs if r["Type"] == "expense")
    inc = sum(float(r["Amount"]) for r in txs if r["Type"] == "income")
    by_cat = {}
    for r in txs:
        if r["Type"] == "expense":
            by_cat[r["Category"]] = by_cat.get(r["Category"],0) + float(r["Amount"])
    return exp, inc, by_cat

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def cat_keyboard(tx_type="expense"):
    cats = EXPENSE_CATS if tx_type=="expense" else INCOME_CATS
    buttons, row = [], []
    for cat in cats:
        row.append(InlineKeyboardButton(cat, callback_data=f"cat:{cat}"))
        if len(row)==2: buttons.append(row); row=[]
    if row: buttons.append(row)
    buttons.append([InlineKeyboardButton("❌ Cancel", callback_data="cat:cancel")])
    return InlineKeyboardMarkup(buttons)

def main_keyboard():
    buttons = []
    if WEBAPP_URL:
        buttons.append([InlineKeyboardButton("📊 Open Dashboard", web_app=WebAppInfo(url=WEBAPP_URL))])
    buttons.append([
        InlineKeyboardButton("📋 /stats", callback_data="cmd:stats"),
        InlineKeyboardButton("📈 /compare", callback_data="cmd:compare"),
    ])
    buttons.append([
        InlineKeyboardButton("📤 /export", callback_data="cmd:export"),
        InlineKeyboardButton("❓ /help", callback_data="cmd:help"),
    ])
    return InlineKeyboardMarkup(buttons)

# ─── COMMANDS ─────────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    name = update.effective_user.first_name or "there"
    # Create their sheet on first use
    sheet = get_user_sheet(uid)
    storage = "Google Sheets ✅" if sheet else "local storage ⚠️"
    await update.message.reply_text(
        f"👋 Hi {name}! Welcome to *Finance Bot*\n\n"
        f"Your data storage: {storage}\n"
        f"Your data is *private* — only you can see it.\n\n"
        "📝 *How to add transactions:*\n"
        "`coffee 4.5` — expense\n"
        "`salary +3411` — income\n"
        "`coffee 4.5 25.05` — with custom date\n\n"
        "📋 *Commands:*\n"
        "/stats — monthly breakdown\n"
        "/compare — vs last month\n"
        "/export — download CSV\n"
        "/exportxls — download Excel report\n"
        "/find coffee — search transactions\n"
        "/last — last 10 transactions\n"
        "/addq coffee 4.5 Dining — save quick template\n"
        "/q coffee — use quick template\n"
        "/budget Groceries 400 — set spending limit\n"
        "/deletedata — delete all your data\n\n"
        "Your data belongs to you and is never shared.",
        parse_mode="Markdown",
        reply_markup=main_keyboard()
    )

async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or (update.callback_query.message if update.callback_query else None)
    if msg:
        await msg.reply_text(
            "📖 *Finance Bot Help*\n\n"
            "*Adding transactions:*\n"
            "`coffee 4.5` → expense\n"
            "`salary +3411` → income\n"
            "`coffee 4.5 25.05` → with date\n\n"
            "*Commands:*\n"
            "/stats `[YYYY-MM]` — monthly stats\n"
            "/compare — this month vs last month\n"
            "/export — download CSV\n"
            "/find `query` — search\n"
            "/last — last 10 transactions\n"
            "/addq `name amount category` — save template\n"
            "/q `name` — use template\n"
            "/budget `category limit` — set budget\n"
            "/deletedata — delete all your data\n\n"
            "*Privacy:*\n"
            "Your data is stored in a private sheet tab.\n"
            "No other user can access it.",
            parse_mode="Markdown"
        )

async def stats_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    month = ctx.args[0] if ctx.args else datetime.now().strftime("%Y-%m")
    exp, inc, by_cat = get_month_stats(uid, month)
    lines = [f"📊 *Stats for {month}*\n"]
    if by_cat:
        for cat, amt in sorted(by_cat.items(), key=lambda x: x[1], reverse=True):
            bar = "▓" * min(int(amt/50),8)
            lines.append(f"{cat}: *{amt:.2f}€* {bar}")
    else:
        lines.append("_No expenses yet_")
    lines.append(f"\n💸 Total expenses: *{exp:.2f}€*")
    if inc > 0:
        lines.append(f"💚 Total income: *{inc:.2f}€*")
        lines.append(f"📈 Balance: *{inc-exp:.2f}€*")
    msg = update.message or (update.callback_query.message if update.callback_query else None)
    if msg: await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def compare_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    now = datetime.now()
    cur = now.strftime("%Y-%m")
    prev = (now.replace(day=1)-timedelta(days=1)).strftime("%Y-%m")
    exp_c,inc_c,by_c = get_month_stats(uid, cur)
    exp_p,inc_p,by_p = get_month_stats(uid, prev)
    diff = exp_c - exp_p
    sign = "+" if diff>=0 else ""
    pct = (diff/exp_p*100) if exp_p>0 else 0
    lines = [f"📊 *{cur} vs {prev}*\n"]
    lines.append(f"Expenses: *{exp_c:.2f}€* vs {exp_p:.2f}€ ({sign}{diff:.2f}€, {sign}{pct:.0f}%)")
    lines.append(f"Income: *{inc_c:.2f}€* vs {inc_p:.2f}€\n")
    all_cats = set(list(by_c.keys())+list(by_p.keys()))
    for cat in sorted(all_cats, key=lambda c: by_c.get(c,0), reverse=True)[:6]:
        c=by_c.get(cat,0); p=by_p.get(cat,0); d=c-p
        arrow="↑" if d>0 else ("↓" if d<0 else "→")
        lines.append(f"{cat}: {c:.0f}€ {arrow} ({'+' if d>=0 else ''}{d:.0f}€)")
    msg = update.message or (update.callback_query.message if update.callback_query else None)
    if msg: await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def export_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    rows = get_tx(uid)
    msg = update.message or (update.callback_query.message if update.callback_query else None)
    if not rows:
        if msg: await msg.reply_text("No transactions to export yet.")
        return
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=HEADERS)
    writer.writeheader()
    writer.writerows(rows)
    output.seek(0)
    filename = f"my_finances_{datetime.now().strftime('%Y%m')}.csv"
    if msg:
        await msg.reply_document(
            document=io.BytesIO(output.getvalue().encode("utf-8")),
            filename=filename,
            caption=f"📤 Your transactions — {len(rows)} records\n_Only you received this file._",
            parse_mode="Markdown"
        )

async def find_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not ctx.args:
        await update.message.reply_text("Usage: `/find coffee`", parse_mode="Markdown")
        return
    query = " ".join(ctx.args).lower()
    rows = get_tx(uid)
    found = [r for r in rows if query in r.get("Description","").lower() or query in r.get("Category","").lower()]
    if not found:
        await update.message.reply_text(f"Nothing found for «{query}»")
        return
    found = list(reversed(found[-10:]))
    lines = [f"🔍 *Results for «{query}»*\n"]
    for r in found:
        sign = "+" if r["Type"]=="income" else "-"
        lines.append(f"{r['Date']} | {sign}{r['Amount']}€ | {r['Category']} | {r['Description']}")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def last_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    rows = get_tx(uid)
    last = list(reversed(rows[-10:])) if rows else []
    if not last:
        await update.message.reply_text("No transactions yet.")
        return
    lines = ["📋 *Last 10 transactions:*\n"]
    for r in last:
        sign = "+" if r["Type"]=="income" else "-"
        lines.append(f"{r['Date']} | {sign}{r['Amount']}€ | {r['Category']} | {r['Description']}")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def budget_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if ctx.args and len(ctx.args)>=2:
        cat_q = " ".join(ctx.args[:-1]).lower()
        try:
            limit = float(ctx.args[-1])
            matched = next((c for c in EXPENSE_CATS if cat_q in c.lower()), " ".join(ctx.args[:-1]))
            # Store budget as special row
            write_user_tx(uid, datetime.now(), limit, matched, "__budget__", "budget")
            await update.message.reply_text(f"✅ Budget set: *{matched}* → {limit:.2f}€/month", parse_mode="Markdown")
            return
        except: pass
    month = datetime.now().strftime("%Y-%m")
    _, _, by_cat = get_month_stats(uid, month)
    rows = get_tx(uid)
    budgets = {r["Category"]: float(r["Amount"]) for r in rows if r.get("Type")=="budget"}
    if not budgets:
        await update.message.reply_text("No budgets set.\nUse: `/budget Groceries 400`", parse_mode="Markdown")
        return
    lines = ["🎯 *Monthly Budgets*\n"]
    for cat, lim in budgets.items():
        spent = by_cat.get(cat,0)
        pct = spent/lim*100 if lim>0 else 0
        bar = "█"*int(pct/10)+"░"*(10-int(pct/10))
        status = "🔴" if pct>=100 else ("🟡" if pct>=80 else "🟢")
        lines.append(f"{status} {cat}\n`{bar}` {pct:.0f}%\n{spent:.2f}€ / {lim:.2f}€\n")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def addq_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not ctx.args or len(ctx.args)<3:
        await update.message.reply_text("Usage: `/addq coffee 4.50 Dining`", parse_mode="Markdown")
        return
    name = ctx.args[0].lower()
    try: amount = float(ctx.args[1])
    except: await update.message.reply_text("Invalid amount."); return
    cat_q = " ".join(ctx.args[2:]).lower()
    cat = next((c for c in EXPENSE_CATS+INCOME_CATS if cat_q in c.lower()), " ".join(ctx.args[2:]))
    write_user_tx(uid, datetime.now(), amount, cat, f"__template__{name}", "template")
    await update.message.reply_text(f"✅ Template saved: `{name}` = {amount:.2f}€ ({cat})", parse_mode="Markdown")

async def q_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    rows = get_tx(uid)
    templates = {r["Description"].replace("__template__",""): {"amount":float(r["Amount"]),"category":r["Category"]} for r in rows if r.get("Type")=="template"}
    if not ctx.args:
        if not templates:
            await update.message.reply_text("No templates. Use `/addq coffee 4.50 Dining`", parse_mode="Markdown")
            return
        lines = ["⚡ *Quick Templates:*\n"]
        for name,v in templates.items():
            lines.append(f"`/q {name}` — {v['amount']:.2f}€ ({v['category']})")
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
        return
    name = ctx.args[0].lower()
    if name not in templates:
        await update.message.reply_text(f"Template `{name}` not found.", parse_mode="Markdown")
        return
    t = templates[name]
    write_user_tx(uid, datetime.now(), t["amount"], t["category"], name, "expense")
    await update.message.reply_text(f"⚡ *Quick add:* -{t['amount']:.2f}€ ({t['category']})", parse_mode="Markdown")

async def deletedata_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Yes, delete everything", callback_data="deleteconfirm:yes"),
        InlineKeyboardButton("❌ Cancel", callback_data="deleteconfirm:no"),
    ]])
    await update.message.reply_text(
        "⚠️ *Are you sure?*\n\nThis will permanently delete ALL your transactions and data. This cannot be undone.",
        parse_mode="Markdown", reply_markup=keyboard
    )


async def exportxls_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    rows = get_tx(uid)
    rows = [r for r in rows if r.get("Type") not in ("budget","template")]
    if not rows:
        await update.message.reply_text("No transactions to export yet.")
        return
    if not XLSX_OK:
        await export_cmd(update, ctx)
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    # Styles
    gold = "B8860B"
    dark = "1A1A1A"
    light_gold = "FEF9EC"
    light_gray = "F5F5F5"
    red_fill = "FFF0F0"
    green_fill = "F0FFF4"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=dark)
    header_align = Alignment(horizontal="center", vertical="center")

    title_font = Font(bold=True, color=gold, size=14)
    ws.merge_cells("A1:G1")
    ws["A1"] = f"💰 Finance Report — {update.effective_user.first_name or 'User'}"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = PatternFill("solid", fgColor="1A1A1A")

    ws.merge_cells("A2:G2")
    from datetime import datetime as dt
    ws["A2"] = f"Generated: {dt.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A2"].font = Font(color="888888", size=10, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")
    ws["A2"].fill = PatternFill("solid", fgColor="1A1A1A")

    ws.append([])  # row 3 empty

    # Headers
    headers = ["Date", "Amount (€)", "Category", "Description", "Type", "Month"]
    ws.append(headers)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Data
    for row in rows:
        tx_type = row.get("Type","expense")
        amount = float(row.get("Amount",0))
        ws.append([
            row.get("Date",""),
            amount if tx_type=="income" else -amount,
            row.get("Category",""),
            row.get("Description",""),
            tx_type.capitalize(),
            row.get("Month",""),
        ])
        r = ws.max_row
        fill_color = green_fill if tx_type=="income" else red_fill
        for col in range(1,7):
            ws.cell(r,col).fill = PatternFill("solid", fgColor=fill_color)
            ws.cell(r,col).alignment = Alignment(vertical="center")
        ws.cell(r,2).number_format = '#,##0.00 "€"'

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Summary by Month"
    ws2["A1"].font = Font(bold=True, color=gold, size=13)
    ws2["A1"].fill = PatternFill("solid", fgColor=dark)

    by_month = {}
    for row in rows:
        m = row.get("Month","")
        t = row.get("Type","expense")
        amt = float(row.get("Amount",0))
        if m not in by_month: by_month[m]={"income":0,"expense":0}
        by_month[m][t] = by_month[m].get(t,0)+amt

    ws2.append([])
    ws2.append(["Month","Income (€)","Expenses (€)","Balance (€)"])
    for col in range(1,5):
        ws2.cell(3,col).font = header_font
        ws2.cell(3,col).fill = header_fill
        ws2.cell(3,col).alignment = header_align

    for month in sorted(by_month.keys()):
        inc = by_month[month].get("income",0)
        exp = by_month[month].get("expense",0)
        bal = inc - exp
        ws2.append([month, inc, exp, bal])
        r = ws2.max_row
        ws2.cell(r,2).font = Font(color="16A34A", bold=True)
        ws2.cell(r,3).font = Font(color="DC2626", bold=True)
        ws2.cell(r,4).font = Font(color=gold if bal>=0 else "DC2626", bold=True)
        for col in range(1,5):
            ws2.cell(r,col).fill = PatternFill("solid", fgColor=light_gray)
            ws2.cell(r,col).number_format = '#,##0.00 "€"' if col>1 else "General"

    # Column widths
    for ws_sheet in [ws, ws2]:
        for col in ws_sheet.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws_sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_len+4, 30)

    # Save
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"finances_{dt.now().strftime('%Y%m%d')}.xlsx"
    await update.message.reply_document(
        document=output,
        filename=filename,
        caption=f"*Your Finance Report*
{len(rows)} transactions exported
Sheet 1: All transactions
Sheet 2: Monthly summary",
        parse_mode="Markdown"
    )

# ─── MESSAGE HANDLER ──────────────────────────────────────────────────────────
async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    uid = update.effective_user.id
    date_pattern = r"(\d{1,2}[./]\d{1,2}(?:[./]\d{2,4})?)"
    custom_date = None
    date_match = re.search(date_pattern, text)
    if date_match:
        date_str = date_match.group(1).replace("/",".")
        text_nd = text.replace(date_match.group(1),"").strip()
        try:
            parts = date_str.split(".")
            if len(parts)==2:
                custom_date = datetime.now().replace(day=int(parts[0]),month=int(parts[1]))
            elif len(parts)==3:
                y=int(parts[2]); y=y+2000 if y<100 else y
                custom_date = datetime(y,int(parts[1]),int(parts[0]))
            if custom_date: text=text_nd
        except: custom_date=None
    pattern = r"^([+]?\d+(?:[.,]\d+)?)\s+(.+)$|^(.+?)\s+([+]?\d+(?:[.,]\d+)?)$|^([+]?\d+(?:[.,]\d+)?)$"
    match = re.match(pattern, text, re.IGNORECASE)
    if not match:
        await update.message.reply_text("Didn't understand 🤔 Try: `coffee 4.50` or `salary +3411`", parse_mode="Markdown")
        return
    g = match.groups()
    if g[0] and g[1]: raw,desc=g[0],g[1]
    elif g[2] and g[3]: desc,raw=g[2],g[3]
    else: raw,desc=g[4],"—"
    raw=raw.replace(",",".")
    is_income=raw.startswith("+")
    amount=float(raw.lstrip("+"))
    tx_type="income" if is_income else "expense"
    pending[uid]={"amount":amount,"description":desc,"type":tx_type,"date":custom_date or datetime.now()}
    sign="+" if is_income else "-"
    date_info=f" ({custom_date.strftime('%d.%m.%Y')})" if custom_date else ""
    await update.message.reply_text(
        f"{'💚' if is_income else '💸'} *{sign}{amount:.2f}€* — {desc}{date_info}\n\nChoose category:",
        reply_markup=cat_keyboard(tx_type), parse_mode="Markdown"
    )

# ─── CALLBACK ─────────────────────────────────────────────────────────────────
async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    uid = query.from_user.id

    if query.data.startswith("cmd:"):
        cmd = query.data.split(":")[1]
        if cmd=="stats": await stats_cmd(update,ctx)
        elif cmd=="compare": await compare_cmd(update,ctx)
        elif cmd=="export": await export_cmd(update,ctx)
        elif cmd=="help": await help_cmd(update,ctx)
        return

    if query.data.startswith("deleteconfirm:"):
        ans = query.data.split(":")[1]
        if ans=="yes":
            sheet = get_user_sheet(uid)
            if sheet:
                try:
                    sp = get_spreadsheet()
                    sp.del_worksheet(sheet)
                except: pass
            try:
                import os as _os
                _os.remove(f"/tmp/{uid}.csv")
            except: pass
            await query.edit_message_text("✅ All your data has been deleted.")
        else:
            await query.edit_message_text("❌ Cancelled. Your data is safe.")
        return

    if query.data=="cat:cancel":
        pending.pop(uid,None)
        await query.edit_message_text("❌ Cancelled.")
        return

    category = query.data.replace("cat:","")
    info = pending.pop(uid,None)
    if not info:
        await query.edit_message_text("Something went wrong, try again.")
        return
    write_user_tx(uid, info["date"], info["amount"], category, info["description"], info["type"])
    sign="+" if info["type"]=="income" else "-"
    await query.edit_message_text(
        f"✅ Saved!\n*{sign}{info['amount']:.2f}€* — {info['description']}\n"
        f"Category: {category}\nDate: {info['date'].strftime('%d.%m.%Y %H:%M')}",
        parse_mode="Markdown"
    )

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("stats", stats_cmd))
    app.add_handler(CommandHandler("compare", compare_cmd))
    app.add_handler(CommandHandler("export", export_cmd))
    app.add_handler(CommandHandler("find", find_cmd))
    app.add_handler(CommandHandler("last", last_cmd))
    app.add_handler(CommandHandler("budget", budget_cmd))
    app.add_handler(CommandHandler("addq", addq_cmd))
    app.add_handler(CommandHandler("q", q_cmd))
    app.add_handler(CommandHandler("deletedata", deletedata_cmd))
    app.add_handler(CommandHandler("exportxls", exportxls_cmd))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("🤖 Finance Bot v3 — Multi-user mode started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
