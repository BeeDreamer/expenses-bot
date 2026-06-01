"""
Finance Bot API — Flask server for dashboard integration
Runs alongside the Telegram bot on Railway
"""

import os, json, hmac, hashlib, time
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app, origins=["https://beedreamer.github.io"])

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "8698682076:AAGa2VWg3MN0IdJcQ64Rtuegg4Mt9GvvCYE")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Finance Bot Data")

# ── GOOGLE SHEETS ──────────────────────────────────────────────────────────────
_spreadsheet = None

def get_spreadsheet():
    global _spreadsheet
    if _spreadsheet:
        return _spreadsheet
    try:
        import gspread, json as _json
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
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
    sp = get_spreadsheet()
    if not sp: return None
    sheet_name = f"user_{user_id}"
    try:
        return sp.worksheet(sheet_name)
    except:
        try:
            import gspread
            sheet = sp.add_worksheet(title=sheet_name, rows=5000, cols=8)
            sheet.append_row(["Date","Amount","Category","Description","Type","Month","UserID"])
            return sheet
        except:
            return None

# ── TELEGRAM AUTH ──────────────────────────────────────────────────────────────
def verify_telegram_data(init_data: str) -> dict | None:
    """Verify Telegram WebApp initData and return user dict if valid"""
    try:
        from urllib.parse import parse_qs, unquote
        parsed = parse_qs(init_data, keep_blank_values=True)
        hash_val = parsed.get('hash', [''])[0]
        
        # Build check string
        data_pairs = []
        for key, values in sorted(parsed.items()):
            if key != 'hash':
                data_pairs.append(f"{key}={values[0]}")
        check_string = '\n'.join(data_pairs)
        
        # Verify
        secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
        expected = hmac.new(secret_key, check_string.encode(), hashlib.sha256).hexdigest()
        
        if not hmac.compare_digest(expected, hash_val):
            return None
        
        # Check not expired (24h)
        auth_date = int(parsed.get('auth_date', ['0'])[0])
        if time.time() - auth_date > 86400:
            return None
            
        user_str = parsed.get('user', ['{}'])[0]
        return json.loads(unquote(user_str))
    except Exception as e:
        print(f"Auth error: {e}")
        return None

def get_user_id_from_request():
    """Extract user ID from request — supports Telegram WebApp and direct user_id param"""
    # Try Telegram WebApp initData first
    init_data = request.headers.get('X-Telegram-Init-Data', '')
    if init_data:
        user = verify_telegram_data(init_data)
        if user:
            return user.get('id')
    
    # Allow direct user_id param (for API access and testing)
    uid = request.args.get('user_id')
    if not uid and request.is_json:
        uid = request.json.get('user_id')
    if uid:
        return int(uid)
    
    return None

# ── ROUTES ─────────────────────────────────────────────────────────────────────
@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'time': datetime.now().isoformat()})

@app.route('/api/transactions', methods=['GET'])
def get_transactions():
    uid = get_user_id_from_request()
    if not uid:
        return jsonify({'error': 'Unauthorized'}), 401
    
    sheet = get_user_sheet(uid)
    if not sheet:
        return jsonify({'transactions': [], 'accounts': get_default_accounts()})
    
    try:
        rows = sheet.get_all_records()
        # Filter out meta rows
        txs = [r for r in rows if r.get('Type') not in ('budget', 'template', 'account')]
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
    
    try:
        rows = sheet.get_all_values()
        for i, row in enumerate(rows[1:], 2):  # skip header
            if row and row[0] == tx_id:
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
        # Remove existing account rows
        rows = sheet.get_all_values()
        to_delete = [i+2 for i, r in enumerate(rows[1:]) if r and r[4] == 'account']
        for i in reversed(to_delete):
            sheet.delete_rows(i)
        # Save new
        for acc in accounts:
            sheet.append_row([
                acc.get('id',''), acc.get('balance', 0), acc.get('type','bank'),
                acc.get('name',''), 'account',
                'primary' if acc.get('primary') else '',
                str(uid)
            ])
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_accounts(uid, sheet):
    if not sheet:
        return get_default_accounts()
    try:
        rows = sheet.get_all_records()
        acc_rows = [r for r in rows if r.get('Type') == 'account']
        if not acc_rows:
            return get_default_accounts()
        return [{'id': r['Date'], 'balance': float(r['Amount'] or 0), 'type': r['Category'], 'name': r['Description'], 'primary': r['Month'] == 'primary'} for r in acc_rows]
    except:
        return get_default_accounts()

def get_default_accounts():
    return [
        {'id': 'a1', 'name': 'Trade Republic', 'type': 'invest', 'balance': 0, 'primary': True},
        {'id': 'a2', 'name': 'Cash', 'type': 'cash', 'balance': 0},
        {'id': 'a3', 'name': 'Revolut', 'type': 'card', 'balance': 0},
    ]

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(host='0.0.0.0', port=port)
