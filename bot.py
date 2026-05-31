"""
Телеграм-бот для учёта расходов с Mini App
"""

import os, re, csv
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

TELEGRAM_TOKEN = "8698682076:AAGa2VWg3MN0IdJcQ64Rtuegg4Mt9GvvCYE"
WEBAPP_URL = os.getenv("WEBAPP_URL", "")
LOCAL_CSV = "expenses.csv"
ALLOWED_USER_ID = 0

CATEGORIES = [
    "🛒 Продукты","🏠 Квартплата","🚌 Транспорт","📱 Связь/Подписки",
    "🍽 Рестораны/Кафе","👕 Одежда","💊 Медицина","🎮 Развлечения",
    "🏋 Спорт","✈️ Путешествия","📚 Учёба","🔧 Быт","💰 Инвестиции","📦 Другое",
]

pending = {}

def ensure_csv():
    if not os.path.exists(LOCAL_CSV):
        with open(LOCAL_CSV, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(["Дата","Сумма","Категория","Описание","Тип","Месяц"])

def read_csv():
    ensure_csv()
    with open(LOCAL_CSV, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def write_csv(date, amount, category, description, tx_type):
    ensure_csv()
    with open(LOCAL_CSV, "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([date.strftime("%d.%m.%Y"), amount, category, description, tx_type, date.strftime("%Y-%m")])

def is_allowed(update: Update) -> bool:
    return ALLOWED_USER_ID == 0 or update.effective_user.id == ALLOWED_USER_ID

def category_keyboard():
    buttons, row = [], []
    for cat in CATEGORIES:
        row.append(InlineKeyboardButton(cat, callback_data=f"cat:{cat}"))
        if len(row) == 2:
            buttons.append(row)
            row = []
    if row: buttons.append(row)
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="cat:cancel")])
    return InlineKeyboardMarkup(buttons)

def main_keyboard():
    buttons = []
    if WEBAPP_URL:
        buttons.append([InlineKeyboardButton("📊 Открыть дашборд", web_app=WebAppInfo(url=WEBAPP_URL))])
    buttons.append([
        InlineKeyboardButton("📋 /stats", callback_data="cmd:stats"),
        InlineKeyboardButton("📝 /last", callback_data="cmd:last"),
    ])
    return InlineKeyboardMarkup(buttons)

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    await update.message.reply_text(
        "👋 *Бот учёта расходов*\n\n"
        "Напиши трату:\n`кофе 4.5` или `продукты 47`\n\n"
        "Доход с плюсом:\n`зарплата +3411`\n\n"
        "/stats — статистика\n/last — последние записи",
        parse_mode="Markdown",
        reply_markup=main_keyboard()
    )

async def stats_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    month = ctx.args[0] if ctx.args else datetime.now().strftime("%Y-%m")
    rows = read_csv()
    expenses = [r for r in rows if r.get("Месяц") == month and r.get("Тип") == "расход"]
    income = [r for r in rows if r.get("Месяц") == month and r.get("Тип") == "доход"]
    by_cat, total_exp = {}, 0
    for r in expenses:
        amt = float(r["Сумма"] or 0)
        by_cat[r["Категория"]] = by_cat.get(r["Категория"], 0) + amt
        total_exp += amt
    total_inc = sum(float(r["Сумма"] or 0) for r in income)
    lines = [f"📊 *Статистика за {month}*\n"]
    if by_cat:
        for cat, amt in sorted(by_cat.items(), key=lambda x: x[1], reverse=True):
            lines.append(f"{cat}: *{amt:.2f} €* {'▓' * min(int(amt/50), 10)}")
    else:
        lines.append("_Нет расходов_")
    lines.append(f"\n💸 Расходы: *{total_exp:.2f} €*")
    if total_inc > 0:
        lines.append(f"💚 Доходы: *{total_inc:.2f} €*")
        lines.append(f"📈 Остаток: *{total_inc - total_exp:.2f} €*")
    msg = update.message or (update.callback_query.message if update.callback_query else None)
    if msg: await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def last_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    rows = read_csv()
    last5 = list(reversed(rows[-5:])) if rows else []
    if not last5:
        await update.message.reply_text("Записей пока нет.")
        return
    lines = ["📋 *Последние записи:*\n"]
    for r in last5:
        sign = "+" if r.get("Тип") == "доход" else "-"
        lines.append(f"{r['Дата']} | {sign}{r['Сумма']} € | {r['Категория']} | {r['Описание']}")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    text = update.message.text.strip()
    user_id = update.effective_user.id
    pattern = r"^([+]?\d+(?:[.,]\d+)?)\s+(.+)$|^(.+?)\s+([+]?\d+(?:[.,]\d+)?)$|^([+]?\d+(?:[.,]\d+)?)$"
    match = re.match(pattern, text, re.IGNORECASE)
    if not match:
        await update.message.reply_text("Не понял 🤔 Напиши: `кофе 4.50`", parse_mode="Markdown")
        return
    g = match.groups()
    if g[0] and g[1]: raw_amount, description = g[0], g[1]
    elif g[2] and g[3]: description, raw_amount = g[2], g[3]
    else: raw_amount, description = g[4], "—"
    raw_amount = raw_amount.replace(",", ".")
    is_income = raw_amount.startswith("+")
    amount = float(raw_amount.lstrip("+"))
    tx_type = "доход" if is_income else "расход"
    pending[user_id] = {"amount": amount, "description": description, "type": tx_type}
    sign = "+" if is_income else "-"
    await update.message.reply_text(
        f"{'💚' if is_income else '💸'} *{sign}{amount:.2f} €* — {description}\n\nВыбери категорию:",
        reply_markup=category_keyboard(), parse_mode="Markdown"
    )

async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data.startswith("cmd:"):
        cmd = query.data.split(":")[1]
        if cmd == "stats": await stats_cmd(update, ctx)
        return

    if query.data == "cat:cancel":
        pending.pop(user_id, None)
        await query.edit_message_text("❌ Отменено.")
        return

    category = query.data.replace("cat:", "")
    info = pending.pop(user_id, None)
    if not info:
        await query.edit_message_text("Попробуй ещё раз.")
        return
    write_csv(datetime.now(), info["amount"], category, info["description"], info["type"])
    sign = "+" if info["type"] == "доход" else "-"
    await query.edit_message_text(
        f"✅ Записано!\n*{sign}{info['amount']:.2f} €* — {info['description']}\n"
        f"Категория: {category}\nДата: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        parse_mode="Markdown"
    )

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(CommandHandler("stats", stats_cmd))
    app.add_handler(CommandHandler("last", last_cmd))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("🤖 Бот запущен!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
