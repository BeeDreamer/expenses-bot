"""
Finance Bot API — Flask server for dashboard integration
Runs alongside the Telegram bot on Railway
"""

import os, json, hmac, hashlib, time
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app, origins=["*"])  # Allow all origins for dashboard access

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Finance Bot Data")

# ── COLUMN MAPPING (supports both Russian and English headers) ─────────────
# Russian headers used by the bot: Дата, Сумма (€), Категория, Описание, Тип, Месяц, UserID
# English headers created by API: Date, Amount, Category, Description, Type, Month, UserID

def normalize_row(row: dict) -> dict:
    """Normalize a row dict to always use English keys, regardless of header language."""
    mapping = {
        # Date
        'дата': 'Date', 'date': 'Date',
        # Amount
        'сумма (€)': 'Amount', 'сумма': 'Amount', 'amount': 'Amount',
        # Category
        'категория': 'Category', 'category': 'Category',
        # Description
        'описание': 'Description', 'description': 'Description',
        # Type
        'тип': 'Type', 'type': 'Type',
        # Month
        'месяц': 'Month', 'month': 'Month',
        # UserID
        'userid': 'UserID', 'user_id': 'UserID',
    }
    normalized = {}
    for k, v in row.items():
        key_norm = mapping.get(k.lower().strip(), k)
        normalized[key_norm] = v
    return normalized

def normalize_type(raw_type: str) -> str:
    """Normalize type field — supports Russian and English."""
    t = (raw_type or '').lower().strip()
    if t in ('расход', 'expense', '-'): return 'expense'
    if t in ('доход', 'income', '+'): return 'income'
    if t in ('бюджет', 'budget'): return 'budget'
    if t in ('шаблон', 'template'): return 'template'
    if t == 'account': return 'account'
    return t  # pass through unknown types

# ── GOOGLE SHEETS ──────────────────────────────────────────────────────────────
_spreadsheet = None

def get_spreadsheet():
    global _spreadsheet
    if _spreadsheet:
        return _spreadsheet
    try:
        import gspread, json as _json
        from google.oauth2.service_account import Credentials
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds_json = os.getenv("GOOGLE_CREDS_JSON")
        if creds_json:
            creds = Credentials.from_service_account_info(_json.loads(creds_json), scopes=scopes)
        else:
            return None
        client = gspread.authorize(creds)
        _spreadsheet = client.open(SPREADSHEET_NAME)
        return _spreadsheet
    except Exception as e:
        print(f"Sheets error: {e}")
        return None

def get_user_sheet(user_id):
    """Get or create a sheet for the user. Detects existing Russian-header sheets."""
    sp = get_spreadsheet()
    if not sp: return None

    sheet_name = f"user_{user_id}"
    try:
        sheet = sp.worksheet(sheet_name)
        return sheet
    except:
        pass

    # Also try to find the sheet by checking all worksheets
    try:
        for ws in sp.worksheets():
            if ws.title == sheet_name:
                return ws
    except:
        pass

    # Create new with English headers
    try:
        sheet = sp.add_worksheet(title=sheet_name, rows=5000, cols=8)
        sheet.append_row(["Date", "Amount", "Category", "Description", "Type", "Month", "UserID"])
        return sheet
    except Exception as e:
        print(f"Create sheet error: {e}")
        return None

def get_all_rows(sheet):
    """Get all rows, normalizing Russian/English headers automatically."""
    try:
        raw = sheet.get_all_records()
        return [normalize_row(r) for r in raw]
    except Exception as e:
        print(f"Read error: {e}")
        return []

# ── TELEGRAM AUTH ──────────────────────────────────────────────────────────────
def verify_telegram_data(init_data: str) -> dict | None:
    """Verify Telegram WebApp initData and return user dict if valid."""
    if not TELEGRAM_TOKEN:
        return None
    try:
        from urllib.parse import parse_qs, unquote
        parsed = parse_qs(init_data, keep_blank_values=True)
        hash_val = parsed.get('hash', [''])[0]

        data_pairs = []
        for key, values in sorted(parsed.items()):
            if key != 'hash':
                data_pairs.append(f"{key}={values[0]}")
        check_string = '\n'.join(data_pairs)

        secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
        expected = hmac.new(secret_key, check_string.encode(), hashlib.sha256).hexdigest()

        if not hmac.compare_digest(expected, hash_val):
            return None

        auth_date = int(parsed.get('auth_date', ['0'])[0])
        if time.time() - auth_date > 86400:
            return None

        user_str = parsed.get('user', ['{}'])[0]
        return json.loads(unquote(user_str))
    except Exception as e:
        print(f"Auth error: {e}")
        return None

def get_user_id_from_request():
    """Extract user ID from request — supports Telegram WebApp and direct user_id param."""
    init_data = request.headers.get('X-Telegram-Init-Data', '')
    if init_data:
        user = verify_telegram_data(init_data)
        if user:
            return user.get('id')

    uid = request.args.get('user_id')
    if not uid and request.is_json:
        uid = request.json.get('user_id')
    if uid:
        return int(uid)

    return None

# ── ROUTES ─────────────────────────────────────────────────────────────────────
@app.route('/health')
def health():
    sp = get_spreadsheet()
    return jsonify({
        'status': 'ok',
        'time': datetime.now().isoformat(),
        'sheets_connected': sp is not None
    })

@app.route('/api/transactions', methods=['GET'])
def get_transactions():
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized — provide user_id parameter'}), 401

    sheet = get_user_sheet(uid)
    if not sheet:
        return jsonify({'transactions': [], 'accounts': get_default_accounts()})

    try:
        rows = get_all_rows(sheet)
        # Filter out meta rows using normalized type
        txs = []
        for r in rows:
            t = normalize_type(r.get('Type', ''))
            if t not in ('budget', 'template', 'account'):
                r['Type'] = t  # store normalized type back
                txs.append(r)
        accs = get_accounts(uid, sheet)
        return jsonify({'transactions': txs, 'accounts': accs})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/transactions', methods=['POST'])
def add_transaction():
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.json
    sheet = get_user_sheet(uid)
    if not sheet:
        return jsonify({'error': 'Storage unavailable'}), 500

    now = datetime.now()
    try:
        # Detect whether sheet uses Russian or English headers
        existing = sheet.get_all_values()
        headers = [h.lower().strip() for h in (existing[0] if existing else [])]
        uses_russian = any('дата' in h or 'сумма' in h or 'тип' in h for h in headers)

        if uses_russian:
            # Write in Russian format to match existing data
            type_ru = 'доход' if data['type'] == 'income' else 'расход'
            sheet.append_row([
                now.strftime("%d.%m.%Y"),
                float(data['amount']),
                data['category'],
                data.get('description', data['category']),
                type_ru,
                now.strftime("%Y-%m"),
                str(uid)
            ])
        else:
            sheet.append_row([
                now.strftime("%d.%m.%Y"),
                float(data['amount']),
                data['category'],
                data.get('description', data['category']),
                data['type'],
                now.strftime("%Y-%m"),
                str(uid)
            ])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/transactions/<tx_id>', methods=['DELETE'])
def delete_transaction(tx_id):
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized'}), 401

    sheet = get_user_sheet(uid)
    if not sheet:
        return jsonify({'error': 'Storage unavailable'}), 500

    # tx_id format: "date|amount|description" e.g. "01.06.2026|5.0|Coffee"
    # Fall back to matching row[0] == tx_id for legacy calls
    try:
        parts = tx_id.split('|')
        match_date = parts[0] if len(parts) > 0 else None
        match_amount = parts[1] if len(parts) > 1 else None
        match_desc = parts[2] if len(parts) > 2 else None

        rows = sheet.get_all_values()
        headers = [h.lower().strip() for h in (rows[0] if rows else [])]

        # find column indices
        date_col = next((i for i,h in enumerate(headers) if 'дата' in h or h=='date'), 0)
        amt_col = next((i for i,h in enumerate(headers) if 'сумма' in h or h=='amount'), 1)
        desc_col = next((i for i,h in enumerate(headers) if 'описан' in h or h=='description'), 3)

        for i, row in enumerate(rows[1:], 2):
            if not row: continue
            # try compound key match first
            if match_date and match_amount:
                row_date = row[date_col] if len(row) > date_col else ''
                row_amt = str(row[amt_col]) if len(row) > amt_col else ''
                row_desc = row[desc_col] if len(row) > desc_col else ''
                # normalize amount for comparison
                try:
                    amt_match = abs(float(row_amt) - float(match_amount)) < 0.01
                except:
                    amt_match = row_amt == match_amount
                if row_date == match_date and amt_match:
                    if not match_desc or match_desc == '—' or row_desc == match_desc or match_desc == '':
                        sheet.delete_rows(i)
                        return jsonify({'success': True})
            # legacy: match by first column
            elif row[0] == tx_id:
                sheet.delete_rows(i)
                return jsonify({'success': True})

        return jsonify({'error': 'Not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/accounts', methods=['GET'])
def get_accounts_route():
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized'}), 401
    sheet = get_user_sheet(uid)
    return jsonify({'accounts': get_accounts(uid, sheet)})

@app.route('/api/accounts', methods=['POST'])
def save_accounts_route():
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized'}), 401

    sheet = get_user_sheet(uid)
    if not sheet:
        return jsonify({'error': 'Storage unavailable'}), 500

    accounts = request.json.get('accounts', [])
    try:
        rows = sheet.get_all_values()
        to_delete = [i+2 for i, r in enumerate(rows[1:]) if r and r[4] == 'account']
        for i in reversed(to_delete):
            sheet.delete_rows(i)
        for acc in accounts:
            sheet.append_row([
                acc.get('id', ''), acc.get('balance', 0), acc.get('type', 'bank'),
                acc.get('name', ''), 'account',
                'primary' if acc.get('primary') else '',
                str(uid)
            ])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_accounts(uid, sheet):
    if not sheet:
        return []
    try:
        rows = get_all_rows(sheet)
        acc_rows = [r for r in rows if normalize_type(r.get('Type', '')) == 'account']
        if not acc_rows:
            return []  # return empty so dashboard keeps its own accounts
        return [{
            'id': r['Date'],
            'balance': float(r.get('Amount') or 0),
            'type': r.get('Category', 'bank'),
            'name': r.get('Description', ''),
            'primary': r.get('Month', '') == 'primary'
        } for r in acc_rows]
    except:
        return []

def get_default_accounts():
    return [
        {'id': 'a1', 'name': 'Trade Republic', 'type': 'invest', 'balance': 0, 'primary': True},
        {'id': 'a2', 'name': 'Cash', 'type': 'cash', 'balance': 0},
        {'id': 'a3', 'name': 'Revolut', 'type': 'card', 'balance': 0},
    ]

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
