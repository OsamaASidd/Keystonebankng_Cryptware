"""
Keystone Bank FIRS e-Invoice Compliance Dashboard
"""
from flask import (Flask, render_template, jsonify, send_file, abort,
                   request, session, redirect, url_for, flash)
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
import pymssql
import json
import io
import re
import time
import logging
import threading
from pathlib import Path
from datetime import datetime, timedelta
from invoice_scheduler_pymssql import InvoiceScheduler, load_config
from apscheduler.schedulers.background import BackgroundScheduler
import openpyxl
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers as xl_numbers)
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import urllib.request
import qrcode
import hmac as _hmac
import hashlib as _hashlib
import requests as _req

app = Flask(__name__, static_folder='assets', static_url_path='/assets')

with open('config.json') as f:
    cfg = json.load(f)

app.secret_key = cfg.get('secret_key', 'kb-firs-2026-xK9mP2qRsT7vWz')
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=15)

# ── Per-worker logging ──────────────────────────────────────────────────────────
_log_dir = cfg.get('logging', {}).get('log_dir', 'logs')
Path(_log_dir).mkdir(parents=True, exist_ok=True)
_log_fmt = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

_fh = logging.FileHandler(Path(_log_dir) / 'dashboard.log', encoding='utf-8')
_fh.setFormatter(_log_fmt)
_sh = logging.StreamHandler()
_sh.setFormatter(_log_fmt)

logging.getLogger().setLevel(logging.INFO)
logging.getLogger().handlers.clear()
logging.getLogger().addHandler(_fh)
logging.getLogger().addHandler(_sh)

app.logger.setLevel(logging.INFO)
# Flask's logger inherits from root — no extra handlers needed

_scheduler_cfg      = load_config('config.json')
_invoice_scheduler  = InvoiceScheduler(_scheduler_cfg)

# ── Auto-post background scheduler ─────────────────────────────────────────────
_interval_minutes = _scheduler_cfg.get('interval_minutes', 5)
_bg_scheduler = BackgroundScheduler(daemon=True)
_bg_scheduler.add_job(
    _invoice_scheduler.process_invoices,
    'interval',
    minutes=_interval_minutes,
    id='auto_post_job',
    name='Auto-post pending FT invoices to NRS',
    max_instances=1,        # never overlap
    coalesce=True,          # skip missed runs
)
_bg_scheduler.start()

_ft_cfg = cfg.get('fund_transfer', {})


def _ft_ref(table_name: str) -> str:
    """Return full 3-part or 4-part table reference for a FUNDS_TRANSFER table."""
    ls = _ft_cfg.get('linked_server')
    db = _ft_cfg.get('database', 'Transactions')
    if ls:
        return f'[{ls}].[{db}].[dbo].[{table_name}]'
    return f'[{db}].[dbo].[{table_name}]'


def _ft_tables(months_back: int = 2):
    """Return list of FUNDS_TRANSFER_YYYYMM names for current + prior months."""
    now = datetime.now()
    result = []
    for i in range(months_back):
        m = now.month - i
        y = now.year
        if m <= 0:
            m += 12
            y -= 1
        result.append(f'FUNDS_TRANSFER_{y}{m:02d}')
    return result


def _parse_commission_amount(raw) -> float:
    """Parse COMMISSION_AMOUNT values like 'NGN0.45' or 'NGN1.00]NGN2.00'."""
    if not raw:
        return 0.0
    raw = str(raw).strip()
    if 'WAIVE' in raw.upper():
        return 0.0
    total = 0.0
    for part in raw.split(']'):
        cleaned = part.replace('NGN', '').replace(',', '').strip()
        try:
            total += float(cleaned)
        except (ValueError, TypeError):
            pass
    return total

# ── Role definitions ───────────────────────────────────────────────────────────
ROLES = {
    'System Admin':  {'label': 'System Admin',  'color': '#4C1D95', 'bg': '#EDE9FE',
                      'can_post': True, 'can_report': True, 'can_export': True, 'can_admin': True},
    'Administrator': {'label': 'Administrator', 'color': '#991B1B', 'bg': '#FEE2E2',
                      'can_post': True, 'can_report': True, 'can_export': True, 'can_admin': True},
    'Manager':       {'label': 'Manager',       'color': '#065F46', 'bg': '#D1FAE5',
                      'can_post': False, 'can_report': True, 'can_export': True, 'can_admin': False},
    'Operator':      {'label': 'Operator',      'color': '#1E40AF', 'bg': '#DBEAFE',
                      'can_post': True, 'can_report': False, 'can_export': True, 'can_admin': False},
    'Viewer':        {'label': 'Viewer',        'color': '#374151', 'bg': '#F3F4F6',
                      'can_post': False, 'can_report': False, 'can_export': False, 'can_admin': False},
}


# ── 2FA helpers ──────────────────────────────────────────────────────────────

def _twofa_enabled() -> bool:
    return bool(cfg.get('twofa', {}).get('enabled', False))

def _twofa_signature(username: str, password: str, accesstoken: str) -> str:
    tfa = cfg.get('twofa', {})
    secret = tfa.get('secret_key', '')
    raw = username.upper() + password.upper() + accesstoken
    return _hmac.new(secret.encode('utf-8'), raw.encode('utf-8'), _hashlib.sha256).hexdigest()

def _call_2fa(username: str, password: str, accesstoken: str):
    """Returns (ok: bool, message: str). Bypasses if 2FA disabled in config."""
    tfa = cfg.get('twofa', {})
    if not tfa.get('enabled', False):
        return True, 'disabled'
    base_url   = tfa.get('base_url', '').rstrip('/')
    client_id  = tfa.get('client_id', '')
    client_key = tfa.get('client_key', '')
    sig = _twofa_signature(username, password, accesstoken)
    url = f'{base_url}/auth/login'
    payload = {'username': username, 'password': '***', 'accesstoken': accesstoken, 'signature': sig}
    logging.info(f'[2FA] POST {url} | payload={json.dumps(payload)}')
    try:
        resp = _req.post(
            url,
            json={'username': username, 'password': password,
                  'accesstoken': accesstoken, 'signature': sig},
            headers={
                'x-consumer-client-id': client_id,
                'x-consumer-client-key': client_key,
                'Accept': 'application/json',
                'Content-Type': 'application/json',
            },
            timeout=10,
            verify=False,
        )
        logging.info(f'[2FA] HTTP {resp.status_code} | body={resp.text[:500]}')
        try:
            data = resp.json()
        except Exception:
            data = {}
        if resp.status_code in (200, 201):
            return True, data.get('message', 'OK')
        return False, data.get('message', f'2FA rejected (HTTP {resp.status_code})')
    except Exception as e:
        logging.error(f'[2FA] Exception: {e}')
        return False, f'2FA service unavailable: {e}'


# ── Security: rate limiting ───────────────────────────────────────────────────
_login_attempts: dict = {}   # {ip: {'count': N, 'lockout_until': float}}
_attempts_lock = threading.Lock()

MAX_ATTEMPTS   = 5           # lock after N failures
LOCKOUT_SECS   = [0, 300, 900, 1800, 3600]  # progressive: 5m, 15m, 30m, 60m

def _check_lockout(ip: str) -> tuple[bool, int]:
    """Returns (is_locked, seconds_remaining)."""
    with _attempts_lock:
        entry = _login_attempts.get(ip)
        if not entry:
            return False, 0
        now = time.time()
        # Purge stale entries older than 2 hours
        _login_attempts.update(
            {k: v for k, v in _login_attempts.items() if now - v.get('first', now) < 7200}
        )
        if entry.get('lockout_until', 0) > now:
            return True, int(entry['lockout_until'] - now)
        return False, 0

def _record_failure(ip: str):
    with _attempts_lock:
        entry = _login_attempts.setdefault(ip, {'count': 0, 'first': time.time(), 'lockout_until': 0})
        entry['count'] += 1
        tier = min(entry['count'] // MAX_ATTEMPTS, len(LOCKOUT_SECS) - 1)
        if entry['count'] >= MAX_ATTEMPTS:
            entry['lockout_until'] = time.time() + LOCKOUT_SECS[tier]

def _clear_attempts(ip: str):
    with _attempts_lock:
        _login_attempts.pop(ip, None)


# ── Security: password complexity ─────────────────────────────────────────────
def _validate_password(pw: str) -> str | None:
    """Returns None if valid, or an error string."""
    if len(pw) < 8:
        return 'Password must be at least 8 characters.'
    if not re.search(r'[A-Z]', pw):
        return 'Password must contain at least one uppercase letter (A–Z).'
    if not re.search(r'[a-z]', pw):
        return 'Password must contain at least one lowercase letter (a–z).'
    if not re.search(r'\d', pw):
        return 'Password must contain at least one number (0–9).'
    if not re.search(r'[!@#$%^&*()\-_=+\[\]{};:\'",.<>/?\\|`~]', pw):
        return 'Password must contain at least one special character (!@#$%^&* …).'
    return None


# ── Security: session inactivity timeout ──────────────────────────────────────
SESSION_TIMEOUT_SECS = 15 * 60  # 15 minutes

@app.before_request
def _check_session_timeout():
    if request.endpoint in ('login', 'logout', 'static') or not request.endpoint:
        return None
    if 'user_id' not in session:
        return None
    last = session.get('_last_activity')
    now  = datetime.now().timestamp()
    if last and (now - last) > SESSION_TIMEOUT_SECS:
        session.clear()
        if (request.is_json or
                request.headers.get('X-Requested-With') == 'XMLHttpRequest'):
            return jsonify({'error': 'session_expired'}), 401
        flash('Your session expired due to inactivity. Please sign in again.', 'warning')
        return redirect(url_for('login'))
    session['_last_activity'] = now
    return None


def _get_session_user():
    if 'user_id' in session:
        return {'id': session['user_id'], 'email': session['email'],
                'role': session['role'], 'full_name': session['full_name']}
    return None


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            if request.is_json or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return jsonify({'error': 'auth_required'}), 401
            return redirect(url_for('login', next=request.full_path))
        return f(*args, **kwargs)
    return decorated


def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('login', next=request.full_path))
            if session.get('role') not in roles:
                flash('Access denied — you do not have permission for this page.', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated
    return decorator


@app.context_processor
def inject_globals():
    user = _get_session_user()
    role_info = ROLES.get(user['role'], {}) if user else {}
    return {'now': datetime.now(), 'current_user': user,
            'current_role': role_info, 'ROLES': ROLES}


def _clean(val):
    if val is None or str(val).strip().upper() == 'NULL':
        return ''
    return str(val).strip()


def get_conn():
    db = cfg['database']
    return pymssql.connect(
        server=db['server'], database=db['name'],
        user=db['user'], password=db['password'], timeout=120
    )


def fetch_invoice(trans_ref):
    """
    Build invoice dict from dbo.RESPONSES (API result) + FUNDS_TRANSFER (source data).
    RESPONSES supplies IRN / QR / status; FUNDS_TRANSFER supplies customer / amount / date.
    Falls back to RESPONSES extended fields when FT row is unavailable.
    """
    conn = get_conn()
    cur  = conn.cursor(as_dict=True)

    # ── Step 1: check RESPONSES ────────────────────────────────────────────────
    cur.execute("""
        SELECT BOOKING_DATE, HTTP_STATUS, IRN, QR_CODE, ERROR_MESSAGE,
               ENVIRONMENT, LAST_UPDATED, RESPONSE_CODE,
               CUSTOMER_NAME, CURRENCY, AMOUNT, CHARGE_DESCRIPTION,
               DEBIT_ACCOUNT_NO, COMMISSION_TYPE
        FROM [dbo].[RESPONSES] WHERE TRANS_REF = %s
    """, (trans_ref,))
    resp = cur.fetchone()

    # ── Step 2: find FT row ────────────────────────────────────────────────────
    ft = None
    if resp:
        # BOOKING_DATE is char(8) YYYYMMDD → derive month table
        yyyymm = str(resp['BOOKING_DATE'])[:6]
        tables_to_try = [f'FUNDS_TRANSFER_{yyyymm}'] + _ft_tables(2)
    else:
        tables_to_try = _ft_tables(2)

    for tbl in dict.fromkeys(tables_to_try):   # deduplicated, order preserved
        try:
            cur.execute(f"""
                SELECT TOP 1
                    TRAN_REFERENCE, DEBIT_VALUE_DATE, DEBIT_CURRENCY, DEBIT_AMOUNT,
                    DEBIT_ACCOUNT_NO, CONTR_NAME, TIN, SUPPL_ADDR,
                    COMMISSION_AMOUNT, COMMISSION_TYPE, CHARGE_CODE, COMMISSION_CODE,
                    ORDERING_CUSTOMER, DEBIT_CUSTOMER, TRAN_TYPE, DATED, PROCESSING_DATE
                FROM {_ft_ref(tbl)} WHERE TRAN_REFERENCE = %s
            """, (trans_ref,))
            ft = cur.fetchone()
            if ft:
                break
        except Exception:
            pass

    conn.close()

    if not resp and not ft:
        return None

    # ── Step 3: build unified invoice dict ────────────────────────────────────
    # Date — prefer FT DEBIT_VALUE_DATE (YYYYMMDD char)
    date_str = ''
    if ft:
        raw = str(ft.get('DEBIT_VALUE_DATE') or ft.get('PROCESSING_DATE') or '').replace('-', '')[:8]
        if len(raw) == 8 and raw.isdigit():
            date_str = f"{raw[:4]}-{raw[4:6]}-{raw[6:8]}"
        elif ft.get('DATED') and hasattr(ft['DATED'], 'strftime'):
            date_str = ft['DATED'].strftime('%Y-%m-%d')
    if not date_str and resp:
        bd = str(resp['BOOKING_DATE'] or '')
        if len(bd) == 8 and bd.isdigit():
            date_str = f"{bd[:4]}-{bd[4:6]}-{bd[6:8]}"

    # Customer details — prefer FT, fall back to RESPONSES extended fields
    if ft:
        customer  = _clean(ft.get('CONTR_NAME')) or _clean(ft.get('ORDERING_CUSTOMER')) or _clean(ft.get('DEBIT_CUSTOMER')) or 'Unknown'
        currency  = _clean(ft.get('DEBIT_CURRENCY')) or 'NGN'
        amount    = _parse_commission_amount(ft.get('COMMISSION_AMOUNT'))
        account_no= _clean(ft.get('DEBIT_ACCOUNT_NO')) or ''
        tin       = _clean(ft.get('TIN')) or ''
        street    = _clean(ft.get('SUPPL_ADDR')) or ''
        comm_type = _clean(ft.get('COMMISSION_TYPE') or '')
        tax_cat   = 'STANDARD_VAT' if 'VAT' in comm_type.upper() else 'ZERO_VAT'
        svc_name  = _clean(ft.get('CHARGE_CODE')) or _clean(ft.get('COMMISSION_CODE')) or ''
        description = svc_name or comm_type or 'Bank Charge'
    else:
        customer   = _clean(resp.get('CUSTOMER_NAME')) or 'Unknown'
        currency   = _clean(resp.get('CURRENCY')) or 'NGN'
        amount     = float(resp.get('AMOUNT') or 0)
        account_no = _clean(resp.get('DEBIT_ACCOUNT_NO')) or ''
        tin        = ''
        street     = ''
        comm_type  = _clean(resp.get('COMMISSION_TYPE') or '')
        tax_cat    = 'STANDARD_VAT' if 'VAT' in comm_type.upper() else 'ZERO_VAT'
        svc_name   = _clean(resp.get('CHARGE_DESCRIPTION')) or ''
        description = svc_name or 'Bank Charge'

    # Status: derive from HTTP_STATUS in RESPONSES
    if resp:
        http_st = resp.get('HTTP_STATUS')
        if http_st in (200, 201):
            status = 1
        elif http_st is not None:
            status = 0
        else:
            status = None
        irn         = _clean(resp.get('IRN')) or ''
        qr_code     = _clean(resp.get('QR_CODE')) or ''
        error       = _clean(resp.get('ERROR_MESSAGE')) or ''
        environment = _clean(resp.get('ENVIRONMENT')) or ''
        last_upd    = str(resp['LAST_UPDATED'])[:19] if resp.get('LAST_UPDATED') else ''
    else:
        status = None; irn = qr_code = error = environment = last_upd = ''

    return {
        'trans_ref':      trans_ref,
        'date':           date_str,
        'invoice_type':   '381',
        'currency':       currency,
        'amount':         amount,
        'account_no':     account_no,
        'customer':       customer,
        'tin':            tin,
        'email':          'noreply@na.ng',
        'street':         street,
        'city':           'Unknown',
        'country':        'NG',
        'postal_zone':    '100001',
        'phone':          '',
        'original_amount': amount,
        'service_fee_name': svc_name,
        'hsn_code':       '6499.00',
        'tax_category':   tax_cat,
        'description':    description,
        'status':         status,
        'irn':            irn,
        'qr_code':        qr_code,
        'error':          error,
        'environment':    environment,
        'last_updated':   last_upd,
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/api/admin/sync', methods=['POST'])
@role_required('System Admin', 'Administrator')
def admin_sync():
    """Immediately run the invoice scheduler job and return updated RESPONSES stats."""
    def _get_stats():
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT
                SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END),
                SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END),
                COUNT(*)
            FROM [dbo].[RESPONSES]
        """)
        r = conn.cursor().fetchone() if False else cur.fetchone()
        conn.close()
        return {'submitted': r[0] or 0, 'failed': r[1] or 0, 'total': r[2] or 0}

    try:
        before = _get_stats()
    except Exception as e:
        return jsonify({'success': False, 'error': f'DB read failed: {e}'}), 500

    try:
        _invoice_scheduler.process_invoices()
    except Exception as e:
        return jsonify({'success': False, 'error': f'Sync error: {e}'}), 500

    try:
        after = _get_stats()
    except Exception as e:
        return jsonify({'success': False, 'error': f'DB read after sync failed: {e}'}), 500

    newly_submitted = after['submitted'] - before['submitted']
    newly_failed    = after['failed']    - before['failed']

    return jsonify({
        'success':         True,
        'processed':       newly_submitted + newly_failed,
        'newly_submitted': newly_submitted,
        'newly_failed':    newly_failed,
        'total':           after['total'],
        'pending':         0,
        'submitted':       after['submitted'],
        'failed':          after['failed'],
    })


@app.route('/')
@login_required
def index():
    conn = get_conn()
    cur  = conn.cursor()

    # Stats from RESPONSES
    cur.execute("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END) as submitted,
            SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END) as failed
        FROM [dbo].[RESPONSES]
    """)
    r = cur.fetchone()
    stats = {'total': r[0] or 0, 'submitted': r[1] or 0,
             'failed': r[2] or 0, 'pending': 0}

    # 10 most-recently processed invoices
    cur.execute("""
        SELECT TOP 10
            TRANS_REF,
            SUBSTRING(BOOKING_DATE,1,4)+'-'+SUBSTRING(BOOKING_DATE,5,2)+'-'+SUBSTRING(BOOKING_DATE,7,2),
            ISNULL(CUSTOMER_NAME,''), ISNULL(AMOUNT,0),
            HTTP_STATUS, ISNULL(IRN,''),
            ISNULL(ERROR_MESSAGE,''), LAST_UPDATED,
            ISNULL(CURRENCY,'NGN')
        FROM [dbo].[RESPONSES]
        ORDER BY LAST_UPDATED DESC
    """)
    recent = []
    for r in cur.fetchall():
        http_st = r[4]
        status  = 1 if http_st in (200, 201) else (0 if http_st is not None else None)
        recent.append({
            'trans_ref':   r[0], 'date': r[1], 'customer': r[2],
            'amount':      r[3], 'status': status, 'irn': r[5],
            'error':       r[6][:80] if r[6] else '',
            'last_updated': str(r[7]) if r[7] else '',
            'currency':    r[8] or 'NGN'
        })

    # Daily submission chart — last 14 days
    cur.execute("""
        SELECT CONVERT(varchar,LAST_UPDATED,23) as day, COUNT(*) as cnt
        FROM [dbo].[RESPONSES]
        WHERE HTTP_STATUS IN (200,201)
          AND LAST_UPDATED >= DATEADD(day,-14,GETDATE())
        GROUP BY CONVERT(varchar,LAST_UPDATED,23)
        ORDER BY day
    """)
    chart_data = [{'day': r[0], 'count': r[1]} for r in cur.fetchall()]

    # Top 5 errors
    cur.execute("""
        SELECT TOP 5 ISNULL(ERROR_MESSAGE,'Unknown'), COUNT(*) as cnt
        FROM [dbo].[RESPONSES]
        WHERE HTTP_STATUS NOT IN (200,201) AND ERROR_MESSAGE IS NOT NULL
        GROUP BY ERROR_MESSAGE
        ORDER BY cnt DESC
    """)
    top_errors = [{'error': r[0][:60], 'count': r[1]} for r in cur.fetchall()]

    conn.close()

    # Scheduler next-run info for dashboard display
    job = _bg_scheduler.get_job('auto_post_job')
    sched_info = {
        'interval_minutes': _interval_minutes,
        'next_run': job.next_run_time.strftime('%d %b %Y %H:%M:%S') if job and job.next_run_time else '—',
    }

    return render_template('index.html', stats=stats, recent=recent,
                           chart_data=chart_data, top_errors=top_errors,
                           sched_info=sched_info)


@app.route('/api/invoices')
@login_required
def api_invoices():
    """DataTables server-side processing endpoint — sources from dbo.RESPONSES (+ FT for pending)."""
    draw      = request.args.get('draw', 1, type=int)
    start     = request.args.get('start', 0, type=int)
    length    = request.args.get('length', 25, type=int)
    search    = request.args.get('search[value]', '').strip()
    status    = request.args.get('status', 'all')
    currency  = request.args.get('currency', 'all')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    # Convert YYYY-MM-DD → YYYYMMDD for BOOKING_DATE char(8) comparisons
    bd_from = date_from.replace('-', '') if date_from else ''
    bd_to   = date_to.replace('-', '')   if date_to   else ''

    try:
        conn = get_conn()
        cur  = conn.cursor()

        # ── Pending filter: FT rows not yet in RESPONSES ─────────────────────
        if status == 'pending':
            rows           = []
            total_filtered = 0
            total_records  = 0

            # Default date range to current month when no filter set
            # (avoids full 22M-row scan on linked server)
            now = datetime.now()
            default_from = f"{now.year}{now.month:02d}01"
            effective_from = bd_from.replace('-', '') if bd_from else default_from
            effective_to   = bd_to.replace('-', '')   if bd_to   else ''

            for tbl in _ft_tables(2):   # current + prior month
                ft_ref = _ft_ref(tbl)
                try:
                    p_where: list[str] = [
                        "ft.COMMISSION_AMOUNT IS NOT NULL",
                        "ft.COMMISSION_AMOUNT != ''",
                        "ft.COMMISSION_AMOUNT NOT LIKE 'WAIVE%'",
                        "ft.TRAN_REFERENCE IS NOT NULL",
                        f"ft.DEBIT_VALUE_DATE >= '{effective_from}'",
                    ]
                    page_params: list = []
                    if effective_to:
                        p_where.append(f"ft.DEBIT_VALUE_DATE <= '{effective_to}'")
                    if currency and currency != 'all':
                        p_where.append("ft.DEBIT_CURRENCY = %s"); page_params.append(currency)
                    if search:
                        p_where.append("ft.TRAN_REFERENCE LIKE %s"); page_params.append(f'%{search}%')
                    p_where_sql = 'WHERE ' + ' AND '.join(p_where)

                    # TOP 500 with no ORDER BY — fast on linked server (no sort)
                    cur.execute(f"""
                        SELECT TOP 500
                            ft.TRAN_REFERENCE,
                            ft.DEBIT_VALUE_DATE,
                            ISNULL(ft.CONTR_NAME, ISNULL(ft.ORDERING_CUSTOMER,'')),
                            ISNULL(ft.DEBIT_CURRENCY,'NGN'),
                            ft.COMMISSION_AMOUNT,
                            ISNULL(ft.CHARGE_CODE, ft.COMMISSION_CODE)
                        FROM {ft_ref} ft {p_where_sql}
                    """, page_params)

                    candidates = cur.fetchall()
                    total_filtered = len(candidates)
                    total_records  = len(candidates)

                    page_slice = candidates[start: start + length]
                    for r in page_slice:
                        ref = r[0]
                        raw_date = str(r[1] or '')
                        dt_str = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:8]}" if len(raw_date) == 8 and raw_date.isdigit() else raw_date
                        amt = _parse_commission_amount(r[4])
                        cur_code = r[3] or 'NGN'
                        rows.append([
                            f'<a href="/invoices/{ref}" class="fw-semibold text-decoration-none">{ref}</a>',
                            dt_str, r[2] or '',
                            f'{cur_code} {amt:,.2f}' if amt else '',
                            '<span class="badge bg-warning text-dark">Pending</span>',
                            '<span class="text-muted">—</span>', '', '', '',
                            f'<button onclick="postInvoice(\'{ref}\',this)" class="btn btn-sm btn-outline-warning me-1" data-bs-toggle="tooltip" title="Post to FIRS now"><i class="bi bi-send"></i></button>'
                            f'<a href="/invoices/{ref}" class="btn btn-sm btn-outline-secondary me-1" data-bs-toggle="tooltip" title="View detail"><i class="bi bi-eye"></i></a>'
                        ])
                    if candidates:
                        break   # found data in this month, stop
                except Exception as e:
                    app.logger.warning(f'Pending query failed for {tbl}: {e}')

            conn.close()
            return jsonify({'draw': draw, 'recordsTotal': total_records,
                            'recordsFiltered': total_filtered, 'data': rows})

        # ── Submitted / Failed / All: query dbo.RESPONSES ─────────────────────
        where, params = [], []

        if currency and currency != 'all':
            where.append("ISNULL(CURRENCY,'NGN') = %s"); params.append(currency)

        if status == 'submitted':
            where.append('HTTP_STATUS IN (200,201)')
        elif status == 'failed':
            where.append('HTTP_STATUS NOT IN (200,201)')

        if search:
            where.append("(TRANS_REF LIKE %s OR ISNULL(CUSTOMER_NAME,'') LIKE %s OR ISNULL(IRN,'') LIKE %s)")
            params += [f'%{search}%', f'%{search}%', f'%{search}%']

        if bd_from:
            where.append('BOOKING_DATE >= %s'); params.append(bd_from)
        if bd_to:
            where.append('BOOKING_DATE <= %s'); params.append(bd_to)

        where_sql = ('WHERE ' + ' AND '.join(where)) if where else ''

        count_params = list(params)
        if count_params:
            cur.execute(f"SELECT COUNT(*) FROM [dbo].[RESPONSES] {where_sql}", count_params)
        else:
            cur.execute(f"SELECT COUNT(*) FROM [dbo].[RESPONSES] {where_sql}")
        total_filtered = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM [dbo].[RESPONSES]")
        total_records = cur.fetchone()[0]

        page_params = list(params) + [start + 1, start + length]
        cur.execute(f"""
            SELECT TRANS_REF,
                   SUBSTRING(BOOKING_DATE,1,4)+'-'+SUBSTRING(BOOKING_DATE,5,2)+'-'+SUBSTRING(BOOKING_DATE,7,2),
                   ISNULL(CUSTOMER_NAME,''), ISNULL(AMOUNT,0),
                   HTTP_STATUS, ISNULL(IRN,''),
                   ISNULL(ERROR_MESSAGE,''), ISNULL(ENVIRONMENT,''),
                   LAST_UPDATED, ISNULL(CURRENCY,'NGN'),
                   ISNULL(ERROR_DETAIL,'')
            FROM (
                SELECT *, ROW_NUMBER() OVER (ORDER BY LAST_UPDATED DESC) AS rn
                FROM [dbo].[RESPONSES] {where_sql}
            ) t WHERE rn BETWEEN %s AND %s
        """, page_params)

        rows = []
        for r in cur.fetchall():
            http_st  = r[4]
            err_msg  = r[6] or ''
            err_det  = r[10] or ''
            if http_st in (200, 201):
                badge    = '<span class="badge" style="background:#dbeafe;color:#1e40af">Submitted</span>'
                can_post = False
            else:
                badge    = '<span class="badge bg-danger">Failed</span>'
                can_post = True

            amt_display = f'{r[9] or "NGN"} {float(r[3]):,.2f}' if r[3] else ''

            # Error cell: short preview + expand button if there's an error
            import html as _html
            if err_msg:
                err_cell = (
                    f'<button class="btn btn-sm btn-danger nrs-err-btn" '
                    f'style="font-size:.72rem;padding:.2rem .5rem" '
                    f'data-msg="{_html.escape(err_msg)}" '
                    f'data-det="{_html.escape(err_det)}">'
                    f'<i class="bi bi-exclamation-circle me-1"></i>Error</button>'
                )
            else:
                err_cell = ''

            rows.append([
                f'<a href="/invoices/{r[0]}" class="fw-semibold text-decoration-none">{r[0]}</a>',
                r[1] or '', r[2] or '', amt_display, badge,
                r[5] or '<span class="text-muted">—</span>',
                err_cell, r[7] or '',
                str(r[8])[:16] if r[8] else '',
                f'<a href="/invoices/{r[0]}" class="btn btn-sm btn-outline-secondary me-1" data-bs-toggle="tooltip" title="View invoice detail"><i class="bi bi-eye"></i></a>'
                + (f'<button onclick="postInvoice(\'{r[0]}\',this)" class="btn btn-sm btn-outline-warning me-1" data-bs-toggle="tooltip" title="Retry — post to FIRS now"><i class="bi bi-send"></i></button>'
                   if can_post else '')
                + f'<a href="/invoices/{r[0]}/print" target="_blank" class="btn btn-sm btn-outline-primary me-1" data-bs-toggle="tooltip" title="Open print-ready invoice with QR code"><i class="bi bi-printer"></i></a>'
                + f'<a href="/invoices/{r[0]}/pdf" class="btn btn-sm btn-outline-success" data-bs-toggle="tooltip" title="Download PDF invoice"><i class="bi bi-file-earmark-pdf"></i></a>'
            ])

        conn.close()
        return jsonify({'draw': draw, 'recordsTotal': total_records,
                        'recordsFiltered': total_filtered, 'data': rows})

    except Exception as e:
        return jsonify({'draw': draw, 'recordsTotal': 0, 'recordsFiltered': 0,
                        'data': [], 'error': str(e)}), 200


@app.route('/invoices/<trans_ref>/post', methods=['POST'])
@login_required
def invoice_post(trans_ref):
    """Manually post a single FUNDS_TRANSFER invoice to the FIRS API immediately."""
    if session.get('role') not in ('System Admin', 'Administrator', 'Operator'):
        return jsonify({'success': False, 'error': 'Permission denied — Operator or Administrator role required.'}), 403
    try:
        scheduler = InvoiceScheduler(_scheduler_cfg)

        # Search for the FT record across recent months
        ft_row = None
        conn = get_conn()
        cur  = conn.cursor(as_dict=True)
        for tbl in _ft_tables(2):
            try:
                cur.execute(f"""
                    SELECT TOP 1
                        RECID, TRAN_REFERENCE, TRAN_TYPE,
                        DEBIT_ACCOUNT_NO, DEBIT_CURRENCY, DEBIT_AMOUNT, DEBIT_VALUE_DATE,
                        COMMISSION_TYPE, COMMISSION_CODE, COMMISSION_AMOUNT,
                        CHARGE_CODE, TIN, CONTR_NAME, SUPPL_ADDR, DATED,
                        DEBIT_CUSTOMER, ORDERING_CUSTOMER, PROCESSING_DATE
                    FROM {_ft_ref(tbl)} WHERE TRAN_REFERENCE = %s
                """, (trans_ref,))
                ft_row = cur.fetchone()
                if ft_row:
                    break
            except Exception:
                pass
        conn.close()

        if not ft_row:
            return jsonify({'success': False, 'error': 'Invoice not found in FUNDS_TRANSFER tables'}), 404

        ft_row['_parsed_amount'] = scheduler._parse_commission_amount(ft_row.get('COMMISSION_AMOUNT', ''))
        payload      = scheduler.map_ft_to_payload(ft_row)
        status_code, response_data = scheduler.submit_invoice(payload)
        booking_date = ft_row.get('DEBIT_VALUE_DATE') or ft_row.get('PROCESSING_DATE') or ft_row.get('DATED')
        scheduler.write_response(trans_ref, booking_date, status_code, response_data, ft_row)

        if status_code in (200, 201):
            irn = response_data.get('data', {}).get('irn', '')
            return jsonify({'success': True, 'irn': irn, 'status_code': status_code})
        else:
            msg = response_data.get('message', str(response_data))
            return jsonify({'success': False, 'error': msg, 'status_code': status_code})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/invoices')
@login_required
def invoices():
    return render_template('invoices.html')


# ── Report ────────────────────────────────────────────────────────────────────

@app.route('/report')
@role_required('System Admin', 'Administrator', 'Manager')
def report():
    conn = get_conn()
    cur  = conn.cursor()

    # Overall summary from RESPONSES
    cur.execute("""
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END) AS submitted,
            SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END) AS failed,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN ISNULL(AMOUNT,0) ELSE 0 END) AS total_submitted_amt,
            SUM(ISNULL(AMOUNT,0)) AS grand_total_amt
        FROM [dbo].[RESPONSES]
    """)
    r = cur.fetchone()
    total     = r[0] or 0
    submitted = r[1] or 0
    summary = {
        'total': total, 'submitted': submitted, 'failed': r[2] or 0, 'pending': 0,
        'total_submitted_amt': float(r[3] or 0),
        'grand_total_amt':     float(r[4] or 0),
        'success_rate': round(submitted / total * 100, 1) if total else 0,
    }

    # Currency breakdown
    cur.execute("""
        SELECT
            ISNULL(CURRENCY,'NGN') AS currency,
            COUNT(*) AS total,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END) AS submitted,
            SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END) AS failed,
            0 AS pending,
            SUM(ISNULL(AMOUNT,0)) AS total_amount,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN ISNULL(AMOUNT,0) ELSE 0 END) AS submitted_amount
        FROM [dbo].[RESPONSES]
        GROUP BY ISNULL(CURRENCY,'NGN')
        ORDER BY total_amount DESC
    """)
    currencies = [{'currency': r[0], 'total': r[1], 'submitted': r[2] or 0,
                   'failed': r[3] or 0, 'pending': 0,
                   'total_amount': float(r[5] or 0),
                   'submitted_amount': float(r[6] or 0)} for r in cur.fetchall()]

    # Tax summary — group by VAT vs non-VAT based on COMMISSION_TYPE
    cur.execute("""
        SELECT
            CASE WHEN COMMISSION_TYPE LIKE '%VAT%' THEN 'STANDARD_VAT' ELSE 'ZERO_VAT' END AS tax_cat,
            COUNT(*) AS invoice_count,
            SUM(ISNULL(AMOUNT,0)) AS base_amount,
            SUM(CASE WHEN COMMISSION_TYPE LIKE '%VAT%'
                     THEN ISNULL(AMOUNT,0)*0.075 ELSE 0 END) AS vat_amount,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) AND COMMISSION_TYPE LIKE '%VAT%'
                     THEN ISNULL(AMOUNT,0)*0.075 ELSE 0 END) AS vat_submitted
        FROM [dbo].[RESPONSES]
        GROUP BY CASE WHEN COMMISSION_TYPE LIKE '%VAT%' THEN 'STANDARD_VAT' ELSE 'ZERO_VAT' END
        ORDER BY invoice_count DESC
    """)
    tax_rows = [{'category': r[0], 'count': r[1],
                 'base_amount':  float(r[2] or 0),
                 'vat_amount':   float(r[3] or 0),
                 'vat_submitted': float(r[4] or 0)} for r in cur.fetchall()]
    total_vat = sum(t['vat_submitted'] for t in tax_rows)

    # Invoice type breakdown — all bank charges are code 381
    cur.execute("""
        SELECT
            381 AS invoice_type,
            COUNT(*) AS cnt,
            SUM(ISNULL(AMOUNT,0)) AS amount,
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END) AS submitted
        FROM [dbo].[RESPONSES]
    """)
    r = cur.fetchone()
    inv_types = [{'code': 381, 'label': 'Sales Invoice',
                  'count': r[1] or 0, 'amount': float(r[2] or 0), 'submitted': r[3] or 0}]

    # Environment breakdown
    cur.execute("""
        SELECT ISNULL(ENVIRONMENT,'UNKNOWN') AS env, COUNT(*) AS cnt,
               SUM(ISNULL(AMOUNT,0)) AS amount
        FROM [dbo].[RESPONSES] WHERE HTTP_STATUS IN (200,201)
        GROUP BY ISNULL(ENVIRONMENT,'UNKNOWN')
    """)
    environments = [{'env': r[0], 'count': r[1], 'amount': float(r[2] or 0)}
                    for r in cur.fetchall()]

    # Top 10 customers by submitted amount
    cur.execute("""
        SELECT TOP 10
            ISNULL(CUSTOMER_NAME,'Unknown') AS customer,
            COUNT(*) AS invoices,
            SUM(ISNULL(AMOUNT,0)) AS amount,
            ISNULL(CURRENCY,'NGN') AS currency
        FROM [dbo].[RESPONSES]
        WHERE HTTP_STATUS IN (200,201)
        GROUP BY ISNULL(CUSTOMER_NAME,'Unknown'), ISNULL(CURRENCY,'NGN')
        ORDER BY amount DESC
    """)
    top_customers = [{'customer': r[0], 'invoices': r[1],
                      'amount': float(r[2] or 0), 'currency': r[3]}
                     for r in cur.fetchall()]

    # Monthly trend (submitted) — BOOKING_DATE is char(8) YYYYMMDD
    cur.execute("""
        SELECT
            SUBSTRING(BOOKING_DATE,1,4)+'-'+SUBSTRING(BOOKING_DATE,5,2) AS month,
            COUNT(*) AS cnt,
            SUM(ISNULL(AMOUNT,0)) AS amount
        FROM [dbo].[RESPONSES]
        WHERE HTTP_STATUS IN (200,201)
        GROUP BY SUBSTRING(BOOKING_DATE,1,4)+'-'+SUBSTRING(BOOKING_DATE,5,2)
        ORDER BY month
    """)
    monthly = [{'month': r[0], 'count': r[1], 'amount': float(r[2] or 0)}
               for r in cur.fetchall()]

    conn.close()
    return render_template('report.html',
        summary=summary, currencies=currencies, tax_rows=tax_rows,
        total_vat=total_vat, inv_types=inv_types, environments=environments,
        top_customers=top_customers, monthly=monthly)


# ── Excel helpers ─────────────────────────────────────────────────────────────

def _xl_header_style():
    return {
        'font':  Font(bold=True, color='FFFFFF', size=10),
        'fill':  PatternFill('solid', fgColor='003087'),
        'align': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': Border(
            bottom=Side(style='thin', color='FFFFFF'),
            right=Side(style='thin', color='FFFFFF'),
        ),
    }

def _xl_apply(cell, font=None, fill=None, align=None, border=None, num_fmt=None):
    if font:   cell.font      = font
    if fill:   cell.fill      = fill
    if align:  cell.alignment = align
    if border: cell.border    = border
    if num_fmt: cell.number_format = num_fmt

def _xl_title_row(ws, text, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    c = ws.cell(1, 1, text)
    _xl_apply(c,
        font=Font(bold=True, size=13, color='003087'),
        align=Alignment(horizontal='center', vertical='center'))
    ws.row_dimensions[1].height = 22
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    ws.cell(2, 1, f'Generated: {datetime.now().strftime("%d %B %Y %H:%M")} | Keystone Bank FIRS Compliance')
    ws.cell(2, 1).font = Font(italic=True, size=9, color='6B7280')
    ws.cell(2, 1).alignment = Alignment(horizontal='center')
    ws.row_dimensions[2].height = 14

def _xl_headers(ws, row, headers):
    hs = _xl_header_style()
    for col, h in enumerate(headers, 1):
        c = ws.cell(row, col, h)
        _xl_apply(c, **hs)
    ws.row_dimensions[row].height = 16

def _xl_autowidth(ws, min_w=10, max_w=45):
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=min_w)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, min_w), max_w)


@app.route('/report/export')
@role_required('System Admin', 'Administrator', 'Manager')
def report_export():
    """Download CFO summary report as multi-sheet Excel workbook."""
    from collections import defaultdict
    sheets_raw = request.args.get('sheets', 'summary,currency,tax,invoices')
    cols_raw   = request.args.get('cols', '')
    active_sheets = set(s.strip() for s in sheets_raw.split(',') if s.strip())
    date_from_raw = request.args.get('date_from', '')
    date_to_raw   = request.args.get('date_to', '')
    bd_from = date_from_raw.replace('-', '') if date_from_raw else ''
    bd_to   = date_to_raw.replace('-', '')   if date_to_raw   else ''

    # Invoice detail column catalogue (same as invoices_export)
    ALL_COLS = [
        ('invoice_no',      'Invoice Number',        False),
        ('invoice_date',    'Invoice Date',          False),
        ('due_date',        'Due Date',              False),
        ('customer_name',   'Customer Name',         False),
        ('customer_tin',    'Customer TIN / Tax ID', False),
        ('customer_addr',   'Customer Address',      False),
        ('narration',       'Details / Narration',   False),
        ('amount',          'Amount',                True ),
        ('subtotal',        'Subtotal',              True ),
        ('vat',             'VAT @ 7.5%',            True ),
        ('total',           'Total (inc. VAT)',       True ),
        ('currency',        'Currency',              False),
        ('account_no',      'Account No.',           False),
        ('commission_type', 'Commission Type',       False),
        ('status',          'Status',                False),
        ('irn',             'IRN',                   False),
        ('environment',     'Environment',           False),
        ('error',           'Error Message',         False),
    ]
    DEFAULT_KEYS = {'invoice_no','invoice_date','due_date','customer_name',
                    'customer_tin','customer_addr','narration','amount','subtotal',
                    'vat','status','irn'}
    selected_keys = [k.strip() for k in cols_raw.split(',') if k.strip()] if cols_raw else list(DEFAULT_KEYS)
    sel_set      = set(selected_keys)
    active_cols  = [(k, lbl, money) for k, lbl, money in ALL_COLS if k in sel_set]
    needs_ft     = 'customer_tin' in sel_set or 'customer_addr' in sel_set

    conn = get_conn()
    cur  = conn.cursor()
    wb   = openpyxl.Workbook()

    GREY  = PatternFill('solid', fgColor='F3F4F6')
    BOLD  = Font(bold=True, size=10)
    MONEY = '#,##0.00'
    KB_BLUE = '003087'
    WHITE   = 'FFFFFF'
    ALT_FILL = PatternFill('solid', fgColor='EDF2FB')
    thin_s   = Side(style='thin', color='D1D5DB')
    cell_bdr = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)

    # Always fetch summary stats (needed for the summary sheet and totals)
    cur.execute("""
        SELECT COUNT(*),
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END),
            SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END),
            SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN ISNULL(AMOUNT,0) ELSE 0 END),
            SUM(ISNULL(AMOUNT,0)),
            MIN(LAST_UPDATED), MAX(LAST_UPDATED)
        FROM [dbo].[RESPONSES]
    """)
    r = cur.fetchone()
    total, submitted, failed = r[0] or 0, r[1] or 0, r[2] or 0
    submitted_amt, grand_amt = float(r[3] or 0), float(r[4] or 0)
    first_sub = str(r[5])[:16] if r[5] else '—'
    last_sub  = str(r[6])[:16] if r[6] else '—'

    cur.execute("""
        SELECT SUM(CASE WHEN COMMISSION_TYPE LIKE '%VAT%' AND HTTP_STATUS IN (200,201)
                        THEN ISNULL(AMOUNT,0)*0.075 ELSE 0 END)
        FROM [dbo].[RESPONSES]
    """)
    total_vat = float(cur.fetchone()[0] or 0)

    # ── Sheet 1: Executive Summary ──────────────────────────────────────────
    ws = wb.active
    if 'summary' in active_sheets:
        ws.title = 'Executive Summary'
        _xl_title_row(ws, 'FIRS e-Invoice — Executive Summary', 4)
        metrics = [
            ('Metric', 'Value', 'Notes', ''),
            ('Total Invoices Processed',   total,         'All records in RESPONSES table', ''),
            ('Successfully Submitted',     submitted,     'HTTP 200/201', ''),
            ('Failed / Exceptions',        failed,        'Non-200 HTTP response', ''),
            ('Submission Success Rate',    f'{round(submitted/total*100,1) if total else 0}%', '', ''),
            ('',)*4,
            ('Total Invoice Amount',       grand_amt,     'All processed invoices', MONEY),
            ('Submitted Invoice Amount',   submitted_amt, 'Successfully posted to FIRS', MONEY),
            ('Total VAT Reported (7.5%)',  total_vat,     'VAT-bearing invoices submitted', MONEY),
            ('',)*4,
            ('First Submission',           first_sub,     '', ''),
            ('Last Submission',            last_sub,      '', ''),
            ('Source Tables',              'FUNDS_TRANSFER_YYYYMM', 'T24 core banking tables', ''),
            ('FIRS Compliance Standard',   'Nigeria FIRS e-Invoice NRS', '', ''),
        ]
        _xl_headers(ws, 4, ['Metric', 'Value', 'Notes', ''])
        for i, row in enumerate(metrics[1:], 5):
            for col, val in enumerate(row, 1):
                c = ws.cell(i, col, val)
                if i % 2 == 0: c.fill = GREY
                if col == 2 and row[3] == MONEY and isinstance(val, float):
                    c.number_format = MONEY; c.alignment = Alignment(horizontal='right')
        _xl_autowidth(ws); ws.column_dimensions['B'].width = 20
    else:
        ws.title = '_remove'

    # ── Sheet 2: Currency Breakdown ─────────────────────────────────────────
    if 'currency' in active_sheets:
        ws2 = wb.create_sheet('Currency Breakdown')
        _xl_title_row(ws2, 'Invoice Breakdown by Currency', 6)
        cur.execute("""
            SELECT ISNULL(CURRENCY,'NGN'), COUNT(*),
                SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN 1 ELSE 0 END),
                SUM(CASE WHEN HTTP_STATUS NOT IN (200,201) THEN 1 ELSE 0 END),
                SUM(ISNULL(AMOUNT,0)),
                SUM(CASE WHEN HTTP_STATUS IN (200,201) THEN ISNULL(AMOUNT,0) ELSE 0 END)
            FROM [dbo].[RESPONSES]
            GROUP BY ISNULL(CURRENCY,'NGN') ORDER BY COUNT(*) DESC
        """)
        _xl_headers(ws2, 4, ['Currency','Total Invoices','Submitted','Failed',
                             'Total Amount','Submitted Amount'])
        for i, r in enumerate(cur.fetchall(), 5):
            row_data = [r[0], r[1], r[2] or 0, r[3] or 0, float(r[4] or 0), float(r[5] or 0)]
            for col, val in enumerate(row_data, 1):
                c = ws2.cell(i, col, val)
                if col in (5, 6): c.number_format = MONEY; c.alignment = Alignment(horizontal='right')
                if i % 2 == 0: c.fill = GREY
        _xl_autowidth(ws2)

    # ── Sheet 3: Tax Analysis ───────────────────────────────────────────────
    if 'tax' in active_sheets:
        ws3 = wb.create_sheet('Tax Analysis')
        _xl_title_row(ws3, 'VAT / Tax Analysis', 5)
        cur.execute("""
            SELECT
                CASE WHEN COMMISSION_TYPE LIKE '%VAT%' THEN 'STANDARD_VAT' ELSE 'ZERO_VAT' END,
                COUNT(*), SUM(ISNULL(AMOUNT,0)),
                SUM(CASE WHEN COMMISSION_TYPE LIKE '%VAT%' THEN ISNULL(AMOUNT,0)*0.075 ELSE 0 END),
                SUM(CASE WHEN HTTP_STATUS IN (200,201) AND COMMISSION_TYPE LIKE '%VAT%'
                         THEN ISNULL(AMOUNT,0)*0.075 ELSE 0 END)
            FROM [dbo].[RESPONSES]
            GROUP BY CASE WHEN COMMISSION_TYPE LIKE '%VAT%' THEN 'STANDARD_VAT' ELSE 'ZERO_VAT' END
            ORDER BY COUNT(*) DESC
        """)
        _xl_headers(ws3, 4, ['Tax Category','Invoice Count','Base Amount',
                             'Estimated VAT','VAT Reported to FIRS'])
        for i, r in enumerate(cur.fetchall(), 5):
            row_data = [r[0], r[1], float(r[2] or 0), float(r[3] or 0), float(r[4] or 0)]
            for col, val in enumerate(row_data, 1):
                c = ws3.cell(i, col, val)
                if col in (3, 4, 5): c.number_format = MONEY; c.alignment = Alignment(horizontal='right')
                if i % 2 == 0: c.fill = GREY
        last = ws3.max_row + 1
        ws3.cell(last, 1, 'TOTAL VAT REPORTED').font = BOLD
        ws3.cell(last, 5, total_vat).number_format = MONEY
        ws3.cell(last, 5).font = Font(bold=True, color='003087')
        _xl_autowidth(ws3)

    # ── Sheet 4: All Invoice Records (branded, column-selectable) ───────────
    if 'invoices' in active_sheets:
        from openpyxl.drawing.image import Image as XLImage
        ws4 = wb.create_sheet('All Invoice Records')
        n_cols = len(active_cols) or 1

        # Branding header (rows 1-5)
        logo_path = Path(__file__).parent / 'assets' / 'logo.png'
        if logo_path.exists():
            logo_img = XLImage(str(logo_path)); logo_img.width = 90; logo_img.height = 68
            ws4.add_image(logo_img, 'A1')
        ws4.merge_cells('A1:C5')
        for row_h, h in zip(range(1, 6), [20, 16, 14, 14, 10]):
            ws4.row_dimensions[row_h].height = h

        def _hc(row, col, val, bold=False, size=10, color=KB_BLUE, align='left'):
            c = ws4.cell(row, col, val)
            c.font = Font(bold=bold, size=size, color=color)
            c.alignment = Alignment(horizontal=align, vertical='center')

        _hc(1, 4, 'KEYSTONE BANK LIMITED', bold=True, size=13)
        _hc(2, 4, 'RC Number: 969956', size=9, color='374151')
        _hc(3, 4, '1 Keystone Crescent, Victoria Island, Lagos', size=9, color='374151')
        _hc(4, 4, 'Tel: 0700-KEYSTONE  |  www.keystonebankng.com', size=9, color='374151')
        _hc(1, max(n_cols, 6), 'FIRS ELECTRONIC INVOICE REPORT', bold=True, size=11, align='right')
        _hc(2, max(n_cols, 6), f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", size=8, color='6B7280', align='right')

        # Column headers row 6
        HDR_ROW = 6
        ws4.row_dimensions[HDR_ROW].height = 26
        hdr_fill = PatternFill('solid', fgColor=KB_BLUE)
        hdr_font = Font(bold=True, color=WHITE, size=10)
        for ci, (_, lbl, _m) in enumerate(active_cols, 1):
            c = ws4.cell(HDR_ROW, ci, lbl)
            c.font = hdr_font; c.fill = hdr_fill; c.border = cell_bdr
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Fetch invoice data + optional FT join (with optional date filter)
        inv_where, inv_params = [], []
        if bd_from: inv_where.append('BOOKING_DATE >= %s'); inv_params.append(bd_from)
        if bd_to:   inv_where.append('BOOKING_DATE <= %s'); inv_params.append(bd_to)
        inv_where_sql = ('WHERE ' + ' AND '.join(inv_where)) if inv_where else ''
        cur.execute(f"""
            SELECT TRANS_REF, BOOKING_DATE,
                   ISNULL(CUSTOMER_NAME,''), ISNULL(DEBIT_ACCOUNT_NO,''),
                   ISNULL(CURRENCY,'NGN'), ISNULL(AMOUNT,0),
                   ISNULL(CHARGE_DESCRIPTION,''), ISNULL(COMMISSION_TYPE,''),
                   CASE WHEN HTTP_STATUS IN (200,201) THEN 'Submitted' ELSE 'Failed' END,
                   ISNULL(IRN,''), ISNULL(ENVIRONMENT,''), ISNULL(ERROR_MESSAGE,'')
            FROM [dbo].[RESPONSES] {inv_where_sql} ORDER BY BOOKING_DATE DESC, TRANS_REF
        """, inv_params if inv_params else None)
        inv_rows = cur.fetchall()

        ft_data = {}
        if needs_ft and inv_rows:
            month_map = defaultdict(list)
            for r in inv_rows: month_map[str(r[1])[:6]].append(r[0])
            for month, refs in month_map.items():
                tbl = _ft_ref(f'FUNDS_TRANSFER_{month}')
                try:
                    ph = ','.join(['%s'] * len(refs))
                    cur.execute(f"SELECT TRAN_REFERENCE, ISNULL(TIN,''), ISNULL(SUPPL_ADDR,'') FROM {tbl} WHERE TRAN_REFERENCE IN ({ph})", refs)
                    for ref, tin, addr in cur.fetchall(): ft_data[ref] = {'tin': tin, 'addr': addr}
                except Exception: pass

        for ri, r in enumerate(inv_rows, HDR_ROW + 1):
            trans_ref = r[0]; bd = str(r[1])
            inv_date  = f"{bd[:4]}-{bd[4:6]}-{bd[6:8]}" if len(bd) == 8 else bd
            amount    = float(r[5]) if r[5] else 0.0
            vat       = round(amount * 0.075, 2)
            ft        = ft_data.get(trans_ref, {})
            row_vals  = {
                'invoice_no': trans_ref, 'invoice_date': inv_date, 'due_date': inv_date,
                'customer_name': r[2], 'customer_tin': ft.get('tin',''),
                'customer_addr': ft.get('addr',''), 'narration': r[6] or r[7],
                'amount': amount, 'subtotal': amount, 'vat': vat,
                'total': round(amount + vat, 2), 'currency': r[4], 'account_no': r[3],
                'commission_type': r[7], 'status': r[8], 'irn': r[9],
                'environment': r[10], 'error': r[11],
            }
            use_alt = (ri % 2 == 0)
            ws4.row_dimensions[ri].height = 15
            for ci, (key, _, is_money) in enumerate(active_cols, 1):
                val = row_vals.get(key, '')
                c = ws4.cell(ri, ci, val if val is not None else '')
                c.border = cell_bdr; c.alignment = Alignment(vertical='center')
                if use_alt: c.fill = ALT_FILL
                if is_money: c.number_format = MONEY; c.alignment = Alignment(horizontal='right', vertical='center')
                if key == 'status':
                    if val == 'Submitted': c.font = Font(color='065F46', bold=True)
                    elif val == 'Failed':  c.font = Font(color='991B1B', bold=True)

        # Totals footer
        if inv_rows:
            fr = HDR_ROW + 1 + len(inv_rows)
            ws4.row_dimensions[fr].height = 16
            tot_fill = PatternFill('solid', fgColor='EAF0FB')
            for ci, (key, _, is_money) in enumerate(active_cols, 1):
                c = ws4.cell(fr, ci); c.fill = tot_fill; c.border = cell_bdr
                if key in ('amount','subtotal','vat','total'):
                    c.value = f"=SUM({get_column_letter(ci)}{HDR_ROW+1}:{get_column_letter(ci)}{fr-1})"
                    c.number_format = MONEY; c.font = Font(bold=True, color=KB_BLUE)
                    c.alignment = Alignment(horizontal='right', vertical='center')
                elif ci == 1:
                    c.value = f'TOTAL  ({len(inv_rows)} records)'
                    c.font = Font(bold=True, color=KB_BLUE); c.alignment = Alignment(vertical='center')

        ws4.freeze_panes = f'A{HDR_ROW + 1}'
        _xl_autowidth(ws4)

    # Remove placeholder sheet if summary was skipped
    if '_remove' in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb['_remove']

    conn.close()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"FIRS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=fname)


@app.route('/invoices/export')
@role_required('System Admin', 'Administrator', 'Manager', 'Operator')
def invoices_export():
    """Download filtered invoice list as branded Excel with user-selected columns."""
    from collections import defaultdict
    from openpyxl.drawing.image import Image as XLImage

    search    = request.args.get('search', '').strip()
    status    = request.args.get('status', 'all')
    currency  = request.args.get('currency', 'all')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')
    cols_raw  = request.args.get('cols', '')

    # Ordered column catalogue: (key, display label, money?)
    ALL_COLS = [
        ('invoice_no',      'Invoice Number',        False),
        ('invoice_date',    'Invoice Date',          False),
        ('due_date',        'Due Date',              False),
        ('customer_name',   'Customer Name',         False),
        ('customer_tin',    'Customer TIN / Tax ID', False),
        ('customer_addr',   'Customer Address',      False),
        ('narration',       'Details / Narration',   False),
        ('amount',          'Amount',                True ),
        ('subtotal',        'Subtotal',              True ),
        ('vat',             'VAT @ 7.5%',            True ),
        ('total',           'Total (inc. VAT)',       True ),
        ('currency',        'Currency',              False),
        ('account_no',      'Account No.',           False),
        ('commission_type', 'Commission Type',       False),
        ('status',          'Status',                False),
        ('irn',             'IRN',                   False),
        ('environment',     'Environment',           False),
        ('error',           'Error Message',         False),
    ]
    DEFAULT_KEYS = {'invoice_no','invoice_date','due_date','customer_name',
                    'customer_tin','customer_addr','narration','amount','subtotal','vat'}

    selected_keys = [k.strip() for k in cols_raw.split(',') if k.strip()] if cols_raw else list(DEFAULT_KEYS)
    sel_set = set(selected_keys)
    active_cols = [(k, lbl, money) for k, lbl, money in ALL_COLS if k in sel_set]
    n_cols = len(active_cols) or 1

    needs_ft = 'customer_tin' in sel_set or 'customer_addr' in sel_set

    bd_from = date_from.replace('-', '') if date_from else ''
    bd_to   = date_to.replace('-', '')   if date_to   else ''

    where, params = [], []
    if currency and currency != 'all':
        where.append("ISNULL(CURRENCY,'NGN') = %s"); params.append(currency)
    if status == 'submitted':
        where.append('HTTP_STATUS IN (200,201)')
    elif status == 'failed':
        where.append('HTTP_STATUS NOT IN (200,201)')
    if search:
        where.append("(TRANS_REF LIKE %s OR ISNULL(CUSTOMER_NAME,'') LIKE %s OR ISNULL(IRN,'') LIKE %s)")
        params += [f'%{search}%', f'%{search}%', f'%{search}%']
    if bd_from:
        where.append('BOOKING_DATE >= %s'); params.append(bd_from)
    if bd_to:
        where.append('BOOKING_DATE <= %s'); params.append(bd_to)
    where_sql = ('WHERE ' + ' AND '.join(where)) if where else ''

    conn = get_conn()
    cur  = conn.cursor()
    q = f"""
        SELECT TRANS_REF, BOOKING_DATE,
               ISNULL(CUSTOMER_NAME,''), ISNULL(DEBIT_ACCOUNT_NO,''),
               ISNULL(CURRENCY,'NGN'), ISNULL(AMOUNT,0),
               ISNULL(CHARGE_DESCRIPTION,''), ISNULL(COMMISSION_TYPE,''),
               CASE WHEN HTTP_STATUS IN (200,201) THEN 'Submitted' ELSE 'Failed' END,
               ISNULL(IRN,''), ISNULL(ENVIRONMENT,''), ISNULL(ERROR_MESSAGE,'')
        FROM [dbo].[RESPONSES] {where_sql}
        ORDER BY BOOKING_DATE DESC, TRANS_REF
    """
    if params:
        cur.execute(q, params)
    else:
        cur.execute(q)
    rows = cur.fetchall()

    # Fetch TIN / address from FT tables if needed
    ft_data = {}
    if needs_ft and rows:
        month_map = defaultdict(list)
        for r in rows:
            month_map[str(r[1])[:6]].append(r[0])
        for month, refs in month_map.items():
            tbl = _ft_ref(f'FUNDS_TRANSFER_{month}')
            try:
                ph = ','.join(['%s'] * len(refs))
                cur.execute(f"""
                    SELECT TRAN_REFERENCE, ISNULL(TIN,''), ISNULL(SUPPL_ADDR,'')
                    FROM {tbl} WHERE TRAN_REFERENCE IN ({ph})
                """, refs)
                for ref, tin, addr in cur.fetchall():
                    ft_data[ref] = {'tin': tin, 'addr': addr}
            except Exception:
                pass
    conn.close()

    # ── Build workbook ───────────────────────────────────────────────────────
    KB_BLUE  = '003087'
    WHITE    = 'FFFFFF'
    ALT_FILL = PatternFill('solid', fgColor='EDF2FB')
    MONEY_FMT = '#,##0.00'
    thin_side = Side(style='thin', color='D1D5DB')
    cell_border = Border(left=thin_side, right=thin_side,
                         top=thin_side, bottom=thin_side)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'e-Invoices'

    # ── Branding header (rows 1-5) ───────────────────────────────────────────
    # Logo: A1:C5
    logo_path = Path(__file__).parent / 'assets' / 'logo.png'
    if logo_path.exists():
        logo_img = XLImage(str(logo_path))
        logo_img.width  = 90
        logo_img.height = 68
        ws.add_image(logo_img, 'A1')
    ws.merge_cells('A1:C5')

    for row_h, h in zip(range(1, 6), [20, 16, 14, 14, 10]):
        ws.row_dimensions[row_h].height = h

    def _hdr_cell(row, col, val, bold=False, size=10, color=KB_BLUE, align='left'):
        c = ws.cell(row, col, val)
        c.font      = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(horizontal=align, vertical='center')

    _hdr_cell(1, 4, 'KEYSTONE BANK LIMITED',                         bold=True, size=13)
    _hdr_cell(2, 4, 'RC Number: 969956',                             size=9,  color='374151')
    _hdr_cell(3, 4, '1 Keystone Crescent, Victoria Island, Lagos',   size=9,  color='374151')
    _hdr_cell(4, 4, 'Tel: 0700-KEYSTONE  |  www.keystonebankng.com', size=9,  color='374151')

    title_col = max(n_cols, 6)
    _hdr_cell(1, title_col, 'FIRS ELECTRONIC SALES INVOICE',
              bold=True, size=12, align='right')
    _hdr_cell(2, title_col, f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}",
              size=8, color='6B7280', align='right')
    _hdr_cell(3, title_col, f"Environment: {cfg.get('api',{}).get('base_url','').split('/')[2]}",
              size=8, color='6B7280', align='right')

    # ── Column header row (row 6) ────────────────────────────────────────────
    HDR_ROW = 6
    ws.row_dimensions[HDR_ROW].height = 26
    hdr_fill = PatternFill('solid', fgColor=KB_BLUE)
    hdr_font = Font(bold=True, color=WHITE, size=10)

    for ci, (_, lbl, _money) in enumerate(active_cols, 1):
        c = ws.cell(HDR_ROW, ci, lbl)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.border    = cell_border
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # ── Data rows (from row 7) ───────────────────────────────────────────────
    for ri, r in enumerate(rows, HDR_ROW + 1):
        trans_ref = r[0]
        bd        = str(r[1])
        inv_date  = f"{bd[:4]}-{bd[4:6]}-{bd[6:8]}" if len(bd) == 8 else bd
        amount    = float(r[5]) if r[5] else 0.0
        narration = r[6] or r[7]
        ft        = ft_data.get(trans_ref, {})
        vat       = round(amount * 0.075, 2)

        row_vals = {
            'invoice_no':      trans_ref,
            'invoice_date':    inv_date,
            'due_date':        inv_date,
            'customer_name':   r[2],
            'customer_tin':    ft.get('tin', ''),
            'customer_addr':   ft.get('addr', ''),
            'narration':       narration,
            'amount':          amount,
            'subtotal':        amount,
            'vat':             vat,
            'total':           round(amount + vat, 2),
            'currency':        r[4],
            'account_no':      r[3],
            'commission_type': r[7],
            'status':          r[8],
            'irn':             r[9],
            'environment':     r[10],
            'error':           r[11],
        }

        use_alt = (ri % 2 == 0)
        ws.row_dimensions[ri].height = 15

        for ci, (key, _, is_money) in enumerate(active_cols, 1):
            val = row_vals.get(key, '')
            c   = ws.cell(ri, ci, val if val is not None else '')
            c.border    = cell_border
            c.alignment = Alignment(vertical='center')
            if use_alt:
                c.fill = ALT_FILL
            if is_money:
                c.number_format = MONEY_FMT
                c.alignment     = Alignment(horizontal='right', vertical='center')
            if key == 'status':
                if val == 'Submitted':
                    c.font = Font(color='065F46', bold=True)
                elif val == 'Failed':
                    c.font = Font(color='991B1B', bold=True)

    # ── Totals footer ────────────────────────────────────────────────────────
    if rows:
        footer_row = HDR_ROW + 1 + len(rows)
        ws.row_dimensions[footer_row].height = 16
        total_fill = PatternFill('solid', fgColor='EAF0FB')
        for ci, (key, _, is_money) in enumerate(active_cols, 1):
            c = ws.cell(footer_row, ci)
            c.fill   = total_fill
            c.border = cell_border
            if key in ('amount', 'subtotal', 'vat', 'total'):
                c.value         = f"=SUM({get_column_letter(ci)}{HDR_ROW+1}:{get_column_letter(ci)}{footer_row-1})"
                c.number_format = MONEY_FMT
                c.font          = Font(bold=True, color=KB_BLUE)
                c.alignment     = Alignment(horizontal='right', vertical='center')
            elif ci == 1:
                c.value     = f'TOTAL  ({len(rows)} records)'
                c.font      = Font(bold=True, color=KB_BLUE)
                c.alignment = Alignment(vertical='center')

    ws.freeze_panes = f'A{HDR_ROW + 1}'
    _xl_autowidth(ws)
    # Keep logo columns narrow
    for col_letter in ('A', 'B', 'C'):
        ws.column_dimensions[col_letter].width = max(
            ws.column_dimensions[col_letter].width, 4)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"FIRS_eInvoice_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=fname)


@app.route('/invoices/<trans_ref>')
@login_required
def invoice_detail(trans_ref):
    invoice = fetch_invoice(trans_ref)
    if not invoice:
        abort(404)
    return render_template('invoice_detail.html', invoice=invoice)


def _make_qr_data_uri(text: str) -> str:
    """Generate a QR code from text and return a PNG data URI."""
    try:
        img = qrcode.make(text)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        import base64
        return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ''


@app.route('/invoices/<trans_ref>/print')
@login_required
def invoice_print(trans_ref):
    invoice = fetch_invoice(trans_ref)
    if not invoice:
        abort(404)
    qr_source = invoice.get('irn') or invoice.get('trans_ref', '')
    if qr_source:
        invoice['qr_data_uri'] = _make_qr_data_uri(qr_source)
    return render_template('invoice_print.html', invoice=invoice)


@app.route('/invoices/<trans_ref>/pdf')
@login_required
def invoice_pdf(trans_ref):
    invoice = fetch_invoice(trans_ref)
    if not invoice:
        abort(404)
    buf = generate_pdf(invoice)
    return send_file(buf, mimetype='application/pdf',
                     as_attachment=True,
                     download_name=f"invoice_{trans_ref}.pdf")


# ── PDF Generator ─────────────────────────────────────────────────────────────

def generate_pdf(inv):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=18*mm, leftMargin=18*mm,
                            topMargin=14*mm, bottomMargin=14*mm)
    styles  = getSampleStyleSheet()
    BLACK   = colors.HexColor('#0d0d0d')
    BLUE    = colors.HexColor('#003087')
    GREEN   = colors.HexColor('#16a34a')
    LGREY   = colors.HexColor('#f3f4f6')
    MGREY   = colors.HexColor('#e5e7eb')
    DGREY   = colors.HexColor('#6b7280')
    WHITE   = colors.white

    def sty(name, **kw):
        return ParagraphStyle(name, parent=styles['Normal'], **kw)

    lbl_s   = sty('lbl',  fontSize=7,  fontName='Helvetica-Bold', textColor=GREEN,  leading=9)
    val_s   = sty('val',  fontSize=10, textColor=BLACK, leading=13)
    val_sm  = sty('vsm',  fontSize=8.5,textColor=BLACK, leading=12)
    right_s = sty('ri',   fontSize=10, textColor=BLACK, alignment=TA_RIGHT, leading=13)
    bold_r  = sty('br',   fontSize=10, fontName='Helvetica-Bold', textColor=BLACK,
                  alignment=TA_RIGHT, leading=13)
    th_s    = sty('th',   fontSize=8.5,fontName='Helvetica-Bold', textColor=BLACK, leading=11)
    th_r    = sty('thr',  fontSize=8.5,fontName='Helvetica-Bold', textColor=BLACK,
                  alignment=TA_RIGHT, leading=11)
    foot_s  = sty('foot', fontSize=7,  textColor=DGREY, alignment=TA_CENTER, leading=10)
    irn_s   = sty('irn',  fontSize=8,  fontName='Courier', textColor=BLUE, leading=11)
    irn_lbl = sty('irnl', fontSize=7,  fontName='Helvetica-Bold', textColor=GREEN, leading=10)

    W         = 174*mm
    LOGO_PATH = 'assets/logo.png'
    elems     = []

    # ── 1. HEADER: "Invoice" left, logo + issuer right ───────────────────────
    invoice_word = Paragraph(
        '<font size="54" color="#0d0d0d"><b>Invoice</b></font>', sty('iw', leading=58))

    try:
        logo = RLImage(LOGO_PATH, width=36*mm, height=13*mm)
    except Exception:
        logo = Paragraph('<b>KEYSTONE BANK</b>', sty('kbl', fontSize=12, textColor=BLUE))

    issuer_lines = [
        [logo],
        [Paragraph('<b>Keystone Bank Limited</b>',
                   sty('bn', fontSize=12, fontName='Helvetica-Bold', textColor=BLACK, leading=16))],
        [Paragraph('1 Keystone Crescent, Victoria Island, Lagos',
                   sty('ba', fontSize=8, textColor=DGREY, leading=11))],
        [Paragraph('RC No: 969956  |  TIN: 00696744-0001  |  www.keystonebankng.com',
                   sty('bb', fontSize=8, textColor=DGREY, leading=11))],
    ]
    issuer_tbl = Table(issuer_lines, colWidths=[84*mm])
    issuer_tbl.setStyle(TableStyle([
        ('TOPPADDING',    (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
        ('ALIGN',         (0,0),(-1,-1), 'RIGHT'),
    ]))

    # TIN / IRN row under issuer block
    tin_cell = [Paragraph('TIN', lbl_s),
                Paragraph('00696744-0001', sty('tv', fontSize=9, textColor=BLACK, leading=12))]
    irn_cell = [Paragraph('IRN', lbl_s),
                Paragraph(inv.get('irn','—'), sty('iv', fontSize=8, fontName='Courier',
                           textColor=BLUE, leading=11))] if inv.get('irn') else None

    meta_right_rows = [tin_cell]
    if irn_cell:
        meta_right_rows.append(irn_cell)
    meta_right_tbl = Table([[item] for item in meta_right_rows],
                           colWidths=[84*mm])
    meta_right_tbl.setStyle(TableStyle([
        ('TOPPADDING',    (0,0),(-1,-1), 2),
        ('BOTTOMPADDING', (0,0),(-1,-1), 2),
        ('ALIGN',         (0,0),(-1,-1), 'RIGHT'),
    ]))

    right_block_rows = issuer_lines + [[Spacer(1, 3*mm)]] + [[item] for item in meta_right_rows]
    right_block = Table(right_block_rows, colWidths=[84*mm])
    right_block.setStyle(TableStyle([
        ('TOPPADDING',    (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
        ('ALIGN',         (0,0),(-1,-1), 'RIGHT'),
    ]))

    hdr = Table([[invoice_word, right_block]], colWidths=[90*mm, 84*mm])
    hdr.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('TOPPADDING',    (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ]))
    elems += [hdr]

    # Horizontal rule
    elems.append(Table([['']], colWidths=[W], rowHeights=[1],
                       style=[('LINEABOVE',(0,0),(-1,-1), 1, MGREY)]))
    elems.append(Spacer(1, 4*mm))

    # ── 2. BILLED TO ─────────────────────────────────────────────────────────
    customer = (inv.get('customer') or '').upper() or '—'
    elems.append(Paragraph(f'<b>Billed To: {customer}</b>',
                            sty('bt', fontSize=11, fontName='Helvetica-Bold',
                                textColor=BLACK, leading=15)))
    elems.append(Spacer(1, 4*mm))

    # Three-column meta grid: Address | Invoice No | Issue Date | Due Date | (TIN)
    addr_str = ' '.join(filter(None, [inv.get('street',''), inv.get('city',''), inv.get('country','')]))
    meta_cols = []
    if addr_str.strip():
        meta_cols.append([Paragraph('Address', lbl_s),
                          Paragraph(addr_str.strip(), val_sm)])
    meta_cols.append([Paragraph('Invoice No', lbl_s),
                      Paragraph(inv.get('trans_ref',''), val_sm)])
    meta_cols.append([Paragraph('Issue Date', lbl_s),
                      Paragraph(inv.get('date','—'), val_sm)])
    meta_cols.append([Paragraph('Due Date',   lbl_s),
                      Paragraph(inv.get('date','—'), val_sm)])
    if inv.get('tin'):
        meta_cols.append([Paragraph('Customer TIN', lbl_s),
                          Paragraph(inv['tin'], val_sm)])
    if inv.get('currency'):
        meta_cols.append([Paragraph('Currency', lbl_s),
                          Paragraph(inv['currency'], val_sm)])

    col_w   = W / max(len(meta_cols), 1)
    meta_tbl = Table(
        [col for col in meta_cols],
        colWidths=[col_w] * len(meta_cols)
    )
    # Rotate: each meta_col is [label, value] — make it a single-column table per field
    # Build as one row of N cells, each cell containing stacked label+value
    def _meta_cell(lbl_txt, val_txt):
        return [Paragraph(lbl_txt, lbl_s), Paragraph(val_txt, val_sm)]

    meta_cells = []
    for mc in meta_cols:
        meta_cells.append(Table(mc, colWidths=[col_w - 2*mm]))

    if meta_cells:
        meta_row_tbl = Table([meta_cells], colWidths=[col_w]*len(meta_cells))
        meta_row_tbl.setStyle(TableStyle([
            ('VALIGN',        (0,0),(-1,-1), 'TOP'),
            ('TOPPADDING',    (0,0),(-1,-1), 0),
            ('BOTTOMPADDING', (0,0),(-1,-1), 0),
            ('LEFTPADDING',   (0,0),(-1,-1), 0),
            ('RIGHTPADDING',  (0,0),(-1,-1), 4),
        ]))
        elems += [meta_row_tbl, Spacer(1, 5*mm)]

    # Horizontal rule
    elems.append(Table([['']], colWidths=[W], rowHeights=[1],
                       style=[('LINEABOVE',(0,0),(-1,-1), 0.8, MGREY)]))
    elems.append(Spacer(1, 4*mm))

    # ── 3. LINE ITEMS TABLE ───────────────────────────────────────────────────
    tax_rate = 7.5 if inv.get('tax_category') == 'STANDARD_VAT' else 0
    amount   = float(inv.get('amount') or 0)
    tax_amt  = round(amount * tax_rate / 100, 2)
    total    = amount + tax_amt
    ccy      = inv.get('currency') or 'NGN'
    sym      = '₦' if ccy == 'NGN' else f'{ccy} '
    desc     = inv.get('service_fee_name') or inv.get('description') or '—'

    item_rows = [
        # Header
        [Paragraph('Item', th_s),
         Paragraph('Description', th_s),
         Paragraph('Qty', th_r),
         Paragraph('Price', th_r)],
        # Data
        [Paragraph('1', val_s),
         Paragraph(desc, val_sm),
         Paragraph('1', right_s),
         Paragraph(f'{sym}{amount:,.2f}', right_s)],
    ]
    col_widths = [14*mm, 98*mm, 18*mm, 44*mm]
    items_tbl = Table(item_rows, colWidths=col_widths)
    items_tbl.setStyle(TableStyle([
        ('LINEBELOW',     (0,0),  (-1,0),  1.2, BLACK),
        ('LINEBELOW',     (0,1),  (-1,-1), 0.6, MGREY),
        ('TOPPADDING',    (0,0),  (-1,-1), 5),
        ('BOTTOMPADDING', (0,0),  (-1,-1), 5),
        ('LEFTPADDING',   (0,0),  (-1,-1), 2),
        ('RIGHTPADDING',  (0,0),  (-1,-1), 2),
        ('VALIGN',        (0,0),  (-1,-1), 'TOP'),
    ]))
    elems += [items_tbl, Spacer(1, 5*mm)]

    # ── 4. TOTALS (right-aligned) ─────────────────────────────────────────────
    totals_rows = [
        [Paragraph('Subtotal',        sty('tl', fontSize=10, textColor=DGREY, leading=13,alignment=TA_RIGHT)),
         Paragraph(f'{sym}{amount:,.2f}',  right_s)],
        [Paragraph(f'VAT ({tax_rate}%)',   sty('tl2',fontSize=10,textColor=DGREY,leading=13,alignment=TA_RIGHT)),
         Paragraph(f'{sym}{tax_amt:,.2f}', right_s)],
        [Paragraph('Total Due',       sty('tg', fontSize=11, fontName='Helvetica-Bold',
                                          textColor=BLACK, leading=15, alignment=TA_RIGHT)),
         Paragraph(f'{sym}{total:,.2f}',   bold_r)],
    ]
    tot_tbl = Table(totals_rows, colWidths=[90*mm, 50*mm],
                    hAlign='RIGHT')
    tot_tbl.setStyle(TableStyle([
        ('TOPPADDING',    (0,0),(-1,-1), 3),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LINEABOVE',     (0,2),(-1,2),  1.2, BLACK),
    ]))
    # Wrap in a right-aligned outer table
    outer = Table([[tot_tbl]], colWidths=[W], hAlign='LEFT')
    outer.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'RIGHT')]))
    elems += [outer, Spacer(1, 8*mm)]

    # ── 5. QR + IRN BLOCK ────────────────────────────────────────────────────
    irn_for_qr = inv.get('irn','') or inv.get('trans_ref','')
    qr_cell = Spacer(22*mm, 22*mm)
    if irn_for_qr:
        try:
            qr_img = qrcode.make(irn_for_qr)
            qr_buf = io.BytesIO()
            qr_img.save(qr_buf, format='PNG')
            qr_buf.seek(0)
            qr_cell = RLImage(qr_buf, width=28*mm, height=28*mm)
        except Exception:
            pass

    if inv.get('irn'):
        irn_block = [
            [qr_cell,
             Table([
                [Paragraph('Invoice Reference Number (IRN)', irn_lbl)],
                [Paragraph(inv['irn'], irn_s)],
                [Paragraph('Scan QR code to verify this invoice on the FIRS portal.', foot_s)],
             ], colWidths=[140*mm],
             style=[('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2)])
            ]
        ]
        qr_tbl = Table(irn_block, colWidths=[32*mm, 142*mm])
        qr_tbl.setStyle(TableStyle([
            ('VALIGN',        (0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',    (0,0),(-1,-1),6),
            ('BOTTOMPADDING', (0,0),(-1,-1),6),
            ('LINEABOVE',     (0,0),(-1,-1),0.7, MGREY),
        ]))
        elems += [qr_tbl, Spacer(1, 5*mm)]

    # ── 6. FOOTER ────────────────────────────────────────────────────────────
    elems += [
        Table([['']], colWidths=[W], rowHeights=[1],
              style=[('LINEABOVE',(0,0),(-1,-1),0.7,MGREY)]),
        Spacer(1, 2*mm),
        Paragraph(
            'FIRS-compliant electronic invoice issued by Keystone Bank Limited  |  '
            'Regulated by the Central Bank of Nigeria  |  '
            f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}",
            foot_s),
    ]

    doc.build(elems)
    buf.seek(0)
    return buf


# ── Auth routes ───────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('index'))
    error = None
    locked = False
    lockout_remaining = 0

    ip = request.remote_addr or '0.0.0.0'
    locked, lockout_remaining = _check_lockout(ip)

    if request.method == 'POST':
        email        = request.form.get('email', '').strip()
        password     = request.form.get('password', '')
        access_token = request.form.get('access_token', '').strip()

        if locked:
            m, s = divmod(lockout_remaining, 60)
            error  = f'Account locked — too many failed attempts. Try again in {m}m {s:02d}s.'
        else:
            try:
                conn = get_conn()
                cur  = conn.cursor()
                cur.execute("""
                    SELECT USER_ID, EMAIL, FULL_NAME, PASSWORD_HASH, ROLE, IS_ACTIVE,
                           ISNULL(ACCESS_TOKEN,'')
                    FROM dbo.FIRS_USERS WHERE EMAIL = %s
                """, (email,))
                row = cur.fetchone()
                if row and row[5] == 1 and check_password_hash(row[3], password):
                    user_role = row[4]
                    needs_2fa = _twofa_enabled() and email.lower() != 'administrator'

                    # ── Step 2: 2FA verification (non-admin only) ─────────────
                    if needs_2fa:
                        if not access_token:
                            conn.close()
                            error = 'Please enter your 2FA access token.'
                        else:
                            ok, msg = _call_2fa(email, password, access_token)
                            if not ok:
                                conn.close()
                                _record_failure(ip)
                                locked, lockout_remaining = _check_lockout(ip)
                                error = f'2FA verification failed: {msg}'
                            else:
                                _clear_attempts(ip)
                                session.permanent = False
                                session['user_id']        = row[0]
                                session['email']          = row[1]
                                session['full_name']      = row[2]
                                session['role']           = user_role
                                session['_last_activity'] = datetime.now().timestamp()
                                cur.execute(
                                    "UPDATE dbo.FIRS_USERS SET LAST_LOGIN=GETDATE() WHERE USER_ID=%s",
                                    (row[0],))
                                conn.commit()
                                conn.close()
                                next_url = request.args.get('next', url_for('index'))
                                return redirect(next_url)
                    else:
                        # Admin roles or 2FA disabled — local auth only
                        _clear_attempts(ip)
                        session.permanent = False
                        session['user_id']        = row[0]
                        session['email']          = row[1]
                        session['full_name']      = row[2]
                        session['role']           = user_role
                        session['_last_activity'] = datetime.now().timestamp()
                        cur.execute(
                            "UPDATE dbo.FIRS_USERS SET LAST_LOGIN=GETDATE() WHERE USER_ID=%s",
                            (row[0],))
                        conn.commit()
                        conn.close()
                        next_url = request.args.get('next', url_for('index'))
                        return redirect(next_url)
                else:
                    conn.close()
                    _record_failure(ip)
                    locked, lockout_remaining = _check_lockout(ip)
                    remaining_attempts = MAX_ATTEMPTS - (_login_attempts.get(ip, {}).get('count', 0))
                    if locked:
                        m, s = divmod(lockout_remaining, 60)
                        error = f'Account locked — too many failed attempts. Try again in {m}m {s:02d}s.'
                    else:
                        error = (f'Invalid credentials or account disabled. '
                                 f'{max(remaining_attempts, 0)} attempt(s) remaining before lockout.')
            except Exception as e:
                error = f'Login error: {e}'

    return render_template('login.html', error=error, locked=locked,
                           lockout_remaining=lockout_remaining,
                           twofa_enabled=_twofa_enabled())


@app.route('/api/ping')
@login_required
def api_ping():
    session['_last_activity'] = datetime.now().timestamp()
    return jsonify({'ok': True})


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── Admin — User Management (RBAC) ────────────────────────────────────────────

@app.route('/admin/users')
@role_required('System Admin', 'Administrator')
def admin_users():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT USER_ID, EMAIL, FULL_NAME, ROLE, IS_ACTIVE,
               CONVERT(varchar,CREATED_AT,120),
               CONVERT(varchar,LAST_LOGIN,120),
               CREATED_BY, ISNULL(ACCESS_TOKEN,'')
        FROM dbo.FIRS_USERS ORDER BY CREATED_AT
    """)
    users = [{'id': r[0], 'email': r[1], 'full_name': r[2], 'role': r[3],
               'is_active': bool(r[4]), 'created_at': r[5] or '',
               'last_login': r[6] or 'Never', 'created_by': r[7] or '',
               'access_token': r[8]}
             for r in cur.fetchall()]
    conn.close()
    return render_template('admin_users.html', users=users, roles=list(ROLES.keys()))


@app.route('/admin/users/add', methods=['POST'])
@role_required('System Admin', 'Administrator')
def admin_users_add():
    email        = request.form.get('email', '').strip()
    full_name    = request.form.get('full_name', '').strip()
    role         = request.form.get('role', 'Viewer')
    password     = request.form.get('password', '')
    access_token = request.form.get('access_token', '').strip()
    if not email or not password:
        flash('Email and password are required.', 'danger')
        return redirect(url_for('admin_users'))
    if not access_token:
        flash('2FA access token is required.', 'danger')
        return redirect(url_for('admin_users'))
    pw_error = _validate_password(password)
    if pw_error:
        flash(f'Password policy: {pw_error}', 'danger')
        return redirect(url_for('admin_users'))
    if role not in ROLES:
        flash('Invalid role.', 'danger')
        return redirect(url_for('admin_users'))
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM dbo.FIRS_USERS WHERE EMAIL = %s", (email,))
        if cur.fetchone()[0] > 0:
            conn.close()
            flash(f'User "{email}" already exists.', 'danger')
            return redirect(url_for('admin_users'))
        ph = generate_password_hash(password)
        cur.execute("""
            INSERT INTO dbo.FIRS_USERS
                (EMAIL, FULL_NAME, PASSWORD_HASH, ROLE, IS_ACTIVE, CREATED_BY, ACCESS_TOKEN)
            VALUES (%s, %s, %s, %s, 1, %s, %s)
        """, (email, full_name or email, ph, role, session['email'], access_token))
        conn.commit()
        conn.close()
        flash(f'User "{email}" created successfully.', 'success')
    except Exception as e:
        flash(f'Error creating user: {e}', 'danger')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/edit', methods=['POST'])
@role_required('System Admin', 'Administrator')
def admin_users_edit(uid):
    email        = request.form.get('email', '').strip()
    full_name    = request.form.get('full_name', '').strip()
    role         = request.form.get('role', 'Viewer')
    is_active    = 1 if request.form.get('is_active') == '1' else 0
    access_token = request.form.get('access_token', '').strip()
    if not email or '@' not in email:
        flash('A valid email is required.', 'danger')
        return redirect(url_for('admin_users'))
    if role not in ROLES:
        flash('Invalid role.', 'danger')
        return redirect(url_for('admin_users'))
    try:
        conn = get_conn()
        cur = conn.cursor()
        # Prevent locking out last admin
        if role != 'Administrator' or is_active == 0:
            cur.execute("""
                SELECT COUNT(*) FROM dbo.FIRS_USERS
                WHERE ROLE='Administrator' AND IS_ACTIVE=1 AND USER_ID != %s
            """, (uid,))
            if cur.fetchone()[0] == 0:
                conn.close()
                flash('Cannot demote or disable the only active Administrator.', 'danger')
                return redirect(url_for('admin_users'))
        # Check email not taken by another user
        cur.execute("SELECT COUNT(*) FROM dbo.FIRS_USERS WHERE EMAIL=%s AND USER_ID != %s", (email, uid))
        if cur.fetchone()[0] > 0:
            conn.close()
            flash(f'Email "{email}" is already in use by another account.', 'danger')
            return redirect(url_for('admin_users'))
        cur.execute("""
            UPDATE dbo.FIRS_USERS
            SET EMAIL=%s, FULL_NAME=%s, ROLE=%s, IS_ACTIVE=%s,
                ACCESS_TOKEN=CASE WHEN %s != '' THEN %s ELSE ACCESS_TOKEN END
            WHERE USER_ID=%s
        """, (email, full_name, role, is_active, access_token, access_token, uid))
        conn.commit()
        conn.close()
        flash('User updated.', 'success')
    except Exception as e:
        flash(f'Error updating user: {e}', 'danger')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/reset-password', methods=['POST'])
@role_required('System Admin', 'Administrator')
def admin_users_reset(uid):
    new_pw = request.form.get('new_password', '')
    pw_error = _validate_password(new_pw)
    if pw_error:
        flash(f'Password policy: {pw_error}', 'danger')
        return redirect(url_for('admin_users'))
    if len(new_pw) < 6:
        flash('Password must be at least 6 characters.', 'danger')
        return redirect(url_for('admin_users'))
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("UPDATE dbo.FIRS_USERS SET PASSWORD_HASH=%s WHERE USER_ID=%s",
                    (generate_password_hash(new_pw), uid))
        conn.commit()
        conn.close()
        flash('Password reset successfully.', 'success')
    except Exception as e:
        flash(f'Error resetting password: {e}', 'danger')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/delete', methods=['POST'])
@role_required('System Admin', 'Administrator')
def admin_users_delete(uid):
    if uid == session.get('user_id'):
        flash('You cannot delete your own account.', 'danger')
        return redirect(url_for('admin_users'))
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("""
            SELECT COUNT(*) FROM dbo.FIRS_USERS
            WHERE ROLE='Administrator' AND IS_ACTIVE=1 AND USER_ID != %s
        """, (uid,))
        if cur.fetchone()[0] == 0:
            conn.close()
            flash('Cannot delete the only active Administrator account.', 'danger')
            return redirect(url_for('admin_users'))
        cur.execute("DELETE FROM dbo.FIRS_USERS WHERE USER_ID=%s", (uid,))
        conn.commit()
        conn.close()
        flash('User deleted.', 'success')
    except Exception as e:
        flash(f'Error deleting user: {e}', 'danger')
    return redirect(url_for('admin_users'))


if __name__ == '__main__':
    import os, wsgiref.simple_server as _wss

    SSL_CERT = 'ssl/server.crt'
    SSL_KEY  = 'ssl/server.key'

    if os.path.exists(SSL_CERT) and os.path.exists(SSL_KEY):

        def _http_redirect(environ, start_response):
            host = (environ.get('HTTP_HOST') or '10.40.24.41').split(':')[0]
            path = environ.get('PATH_INFO', '/')
            qs   = environ.get('QUERY_STRING', '')
            url  = f'https://{host}:5443{path}' + (f'?{qs}' if qs else '')
            start_response('301 Moved Permanently',
                           [('Location', url), ('Content-Type', 'text/plain'),
                            ('Content-Length', '0')])
            return [b'']

        class _ReuseServer(_wss.WSGIServer):
            allow_reuse_address = True

        _redir = _ReuseServer(('0.0.0.0', 5000), _wss.WSGIRequestHandler)
        _redir.set_app(_http_redirect)
        threading.Thread(target=_redir.serve_forever, daemon=True).start()

        print(' * HTTP  :5000 -> redirects to https://:5443')
        print(f' * HTTPS :5443 using {SSL_CERT}')
        app.run(debug=False, host='0.0.0.0', port=5443,
                ssl_context=(SSL_CERT, SSL_KEY))

    else:
        print(' WARNING: SSL certificates not found — running on HTTP (insecure).')
        print('          Generate them first: python generate_ssl_cert.py')
        app.run(debug=True, host='0.0.0.0', port=5000)
