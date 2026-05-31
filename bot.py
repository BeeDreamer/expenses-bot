"""
Телеграм-бот для учёта расходов
Токен: уже вшит
Google Sheets: подключается если есть google_creds.json, иначе сохраняет в expenses.csv
"""

import os
import re
import csv
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)

# ─── НАСТРОЙКИ ───────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = "8698682076:AAGa2VWg3MN0IdJcQ64Rtuegg4Mt9GvvCYE"
GOOGLE_CREDS_FILE = "google_creds.json"
SPREADSHEET_NAME = "Расходы и доходы"
SHEET_NAME = "Транзакции"
LOCAL_CSV = "expenses.csv"
ALLOWED_USER_ID = 0  # 0 = доступ для всех, впиши свой ID чтобы закрыть

# Проверяем есть ли Google Sheets
USE_GOOGLE = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    if os.path.exists(GOOGLE_CREDS_FILE):
        USE_GOOGLE = True
        print("✅ Google Sheets найден — используем таблицу")
    else:
        print("⚠️  google_creds.json не найден — сохраняем в expenses.csv")
except ImportError:
    print("⚠️  gspread не установлен — сохраняем в expenses.csv")

# Категории
CATEGORIES = [
    "🛒 Продукты",
    "🏠 Квартплата",
    "🚌 Транспорт",
    "📱 Связь/Подписки",
    "🍽 Рестораны/Кафе",
    "👕 Одежда",
    "💊 Медицина",
    "🎮 Развлечения",
    "🏋 Спорт",
    "✈️ Путешествия",
    "📚 Учёба",
    "🔧 Быт",
    "💰 Инвестиции",
    "📦 Другое",
]

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────
def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open(SPREADSHEET_NAME)
    try:
        sheet = spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=6)
        sheet.append_row(["Дата", "Сумма (€)", "Категория", "Описание", "Тип", "Месяц"])
    return sheet

# ─── ЛОКАЛЬНЫЙ CSV ────────────────────────────────────────────────────────────
def ensure_csv():
    if not os.path.exists(LOCAL_CSV):
        with open(LOCAL_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Дата", "Сумма (€)", "Категория", "Описание", "Тип", "Месяц"])

def read_csv():
    ensure_csv()
    rows = []
    with open(LOCAL_CSV, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def write_csv(date, amount, category, description, tx_type):
    ensure_csv()
    month = date.strftime("%Y-%m")
    with open(LOCAL_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([date.strftime("%d.%m.%Y"), amount, category, description, tx_type, month])

# ─── СОХРАНЕНИЕ ───────────────────────────────────────────────────────────────
def add_row(date, amount, category, description, tx_type="расход"):
    if USE_GOOGLE:
        sheet = get_sheet()
        month = date.strftime("%Y-%m")
        sheet.append_row([date.strftime("%d.%m.%Y"), amount, category, description, tx_type, month])
    else:
        write_csv(date, amount, category, description, tx_type)

def get_all_rows():
    if USE_GOOGLE:
        return get_sheet().get_all_records()
    else:
        return read_csv()

# ─── СТАТИСТИКА ───────────────────────────────────────────────────────────────
def get_stats(month=None):
    if not month:
        month = datetime.now().strftime("%Y-%m")
    rows = get_all_rows()
    expenses = [r for r in rows if r.get("Месяц") == month and r.get("Тип") == "расход"]
    income   = [r for r in rows if r.get("Месяц") == month and r.get("Тип") == "доход"]

    by_cat = {}
    total_exp = 0
    for r in expenses:
        cat = r["Категория"]
        amt = float(r["Сумма (€)"] or 0)
        by_cat[cat] = by_cat.get(cat, 0) + amt
        total_exp += amt

    total_inc = sum(float(r["Сумма (€)"] or 0) for r in income)
    return total_exp, total_inc, by_cat, month

# ─── СОСТОЯНИЕ ────────────────────────────────────────────────────────────────
pending = {}

# ─── ХЕЛПЕРЫ ─────────────────────────────────────────────────────────────────
def is_allowed(update: Update) -> bool:
    if ALLOWED_USER_ID == 0:
        return True
    return update.effective_user.id == ALLOWED_USER_ID

def category_keyboard():
    buttons = []
    row = []
    for cat in CATEGORIES:
        row.append(InlineKeyboardButton(cat, callback_data=f"cat:{cat}"))
        if len(row) == 2:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="cat:cancel")])
    return InlineKeyboardMarkup(buttons)

# ─── КОМАНДЫ ─────────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    storage = "Google Sheets ✅" if USE_GOOGLE else "локальный файл expenses.csv ⚠️"
    text = (
        f"👋 *Бот учёта расходов*\n"
        f"Хранение: {storage}\n\n"
        "Просто напиши трату:\n"
        "`кофе 4.5`\n"
        "`продукты 47`\n"
        "`такси 12.80`\n\n"
        "Доход — добавь `+`:\n"
        "`зарплата +3411`\n\n"
        "📋 Команды:\n"
        "/stats — статистика за месяц\n"
        "/last — последние 5 записей\n"
        "/help — помощь"
    )
    await update.message.reply_text(text, parse_mode="Markdown")

async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await start(update, ctx)

async def stats_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    month = ctx.args[0] if ctx.args else None
    try:
        total_exp, total_inc, by_cat, month_str = get_stats(month)
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {e}")
        return

    lines = [f"📊 *Статистика за {month_str}*\n"]
    if by_cat:
        for cat, amt in sorted(by_cat.items(), key=lambda x: x[1], reverse=True):
            bar = "▓" * min(int(amt / 50), 10)
            lines.append(f"{cat}: *{amt:.2f} €* {bar}")
    else:
        lines.append("_Нет расходов_")

    lines.append(f"\n💸 Расходы: *{total_exp:.2f} €*")
    if total_inc > 0:
        lines.append(f"💚 Доходы: *{total_inc:.2f} €*")
        lines.append(f"📈 Остаток: *{total_inc - total_exp:.2f} €*")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def last_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    try:
        rows = get_all_rows()
        last5 = list(reversed(rows[-5:])) if rows else []
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {e}")
        return

    if not last5:
        await update.message.reply_text("Записей пока нет.")
        return

    lines = ["📋 *Последние записи:*\n"]
    for r in last5:
        sign = "+" if r.get("Тип") == "доход" else "-"
        lines.append(f"{r['Дата']} | {sign}{r['Сумма (€)']} € | {r['Категория']} | {r['Описание']}")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

# ─── ОБРАБОТКА СООБЩЕНИЙ ─────────────────────────────────────────────────────
async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return

    text = update.message.text.strip()
    user_id = update.effective_user.id

    pattern = r"^([+]?\d+(?:[.,]\d+)?)\s+(.+)$|^(.+?)\s+([+]?\d+(?:[.,]\d+)?)$|^([+]?\d+(?:[.,]\d+)?)$"
    match = re.match(pattern, text, re.IGNORECASE)

    if not match:
        await update.message.reply_text(
            "Не понял 🤔 Напиши например:\n`кофе 4.50` или `продукты 45`",
            parse_mode="Markdown"
        )
        return

    g = match.groups()
    if g[0] and g[1]:
        raw_amount, description = g[0], g[1]
    elif g[2] and g[3]:
        description, raw_amount = g[2], g[3]
    else:
        raw_amount, description = g[4], "—"

    raw_amount = raw_amount.replace(",", ".")
    is_income = raw_amount.startswith("+")
    amount = float(raw_amount.lstrip("+"))
    tx_type = "доход" if is_income else "расход"

    pending[user_id] = {"amount": amount, "description": description, "type": tx_type}

    emoji = "💚" if is_income else "💸"
    sign = "+" if is_income else "-"
    await update.message.reply_text(
        f"{emoji} *{sign}{amount:.2f} €* — {description}\n\nВыбери категорию:",
        reply_markup=category_keyboard(),
        parse_mode="Markdown"
    )

# ─── CALLBACK ─────────────────────────────────────────────────────────────────
async def handle_category(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "cat:cancel":
        pending.pop(user_id, None)
        await query.edit_message_text("❌ Отменено.")
        return

    category = query.data.replace("cat:", "")
    info = pending.pop(user_id, None)

    if not info:
        await query.edit_message_text("Что-то пошло не так, попробуй ещё раз.")
        return

    try:
        add_row(datetime.now(), info["amount"], category, info["description"], info["type"])
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка записи: {e}")
        return

    sign = "+" if info["type"] == "доход" else "-"
    await query.edit_message_text(
        f"✅ Записано!\n\n"
        f"*{sign}{info['amount']:.2f} €* — {info['description']}\n"
        f"Категория: {category}\n"
        f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        parse_mode="Markdown"
    )

# ─── ЗАПУСК ───────────────────────────────────────────────────────────────────
def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("stats", stats_cmd))
    app.add_handler(CommandHandler("last", last_cmd))
    app.add_handler(CallbackQueryHandler(handle_category, pattern="^cat:"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("🤖 Бот запущен! Открой @ExpensesDT_bot в Telegram")
    app.run_polling()

if __name__ == "__main__":
    main()
