"""
FIRS Invoice Scheduler — FUNDS_TRANSFER source
Reads from dbo.FUNDS_TRANSFER_YYYYMM tables (via linked server on prod,
local Transactions DB on preprod), posts to FIRS API, and stores all API
responses in dbo.RESPONSES.
"""

import pymssql
import requests
import json
import logging
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from typing import Dict, List, Optional
import os
import sys
from pathlib import Path

logger = logging.getLogger(__name__)
# Logging is configured by setup_logging() after config is loaded.
# A fallback console handler is set here so early errors are visible.
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    )


def setup_logging(log_dir: str = 'logs', worker: str = 'invoice_scheduler') -> None:
    """Configure root logger to write to logs/<worker>.log and stdout."""
    log_path = Path(log_dir)
    log_path.mkdir(parents=True, exist_ok=True)
    log_file = log_path / f'{worker}.log'

    fmt = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    # Remove any handlers added by basicConfig
    root.handlers.clear()

    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setFormatter(fmt)
    root.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root.addHandler(sh)

    logger.info(f"Logging initialised → {log_file}")


class InvoiceScheduler:
    def __init__(self, config: Dict):
        self.config = config
        self.environment = 'PREPROD' if 'preprod' in config.get('base_url', '').lower() else 'PROD'
        self.ft_linked_server = config.get('ft_linked_server')
        self.ft_database = config.get('ft_database', 'Transactions')

    # ── Utility helpers ────────────────────────────────────────────────────────

    @staticmethod
    def _clean(val) -> str:
        """Return empty string for None or the literal string 'NULL'."""
        if val is None or str(val).strip().upper() == 'NULL':
            return ''
        return str(val).strip()

    @staticmethod
    def _normalise_country(raw: str) -> str:
        _dial = {"234": "NG"}
        raw = (raw or "").strip()
        if raw in _dial:
            return _dial[raw]
        return raw if len(raw) == 2 and raw.isalpha() else "NG"

    @staticmethod
    def _normalise_postal_zone(raw: str) -> str:
        raw = (raw or "").strip()
        return raw if len(raw) >= 5 else "100001"

    @staticmethod
    def _format_hsn_code(raw: str) -> str:
        raw = (raw or "").strip()
        if not raw:
            return "0000.00"
        return raw if '.' in raw else f"{raw}.00"

    @staticmethod
    def _format_phone(raw: str) -> str:
        raw = (raw or "").strip()
        if not raw:
            return "+2340000000000"
        if raw.startswith("+"):
            return raw
        if raw.startswith("234"):
            return "+" + raw
        if raw.startswith("0"):
            return "+234" + raw[1:]
        return "+234" + raw

    @staticmethod
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

    @staticmethod
    def _normalize_uom(commission_type: str, charge_code: str) -> str:
        """
        Map T24 COMMISSION_TYPE / CHARGE_CODE to a valid NRS Unit-of-Measure code.

        NRS codes used:
          E48  – service units (per-event service fee)         ← default
          1I   – rate for usage of a facility or service       ← rate / percentage charges
          LS   – lump sum                                      ← flat / lump charges
          EA   – each (generic count unit)                     ← quantity-based charges
        """
        ct  = (commission_type or '').upper()
        cc  = (charge_code     or '').upper()
        src = ct + ' ' + cc

        # Rate / percentage-based charges
        if any(k in src for k in ('RATE', 'PCT', 'PERCENT', 'RATIO', 'BPS')):
            return '1I'

        # Flat / lump-sum charges
        if any(k in src for k in ('FLAT', 'LUMP', 'FIXED', 'ANNUAL', 'MONTHLY', 'YEARLY')):
            return 'LS'

        # Quantity / per-item charges  (e.g. CHEQUE, LEAF, BOOK, STAMP)
        if any(k in src for k in ('CHEQUE', 'LEAF', 'BOOK', 'STAMP', 'STATEMENT', 'CARD')):
            return 'EA'

        # Default: per-event service fee — covers SWIFT, TRNF, COMM, VAT, FEE, CHARGE, etc.
        return 'E48'

    def _ft_ref(self, table_name: str) -> str:
        """Return full 3-part or 4-part table reference for a FUNDS_TRANSFER table."""
        if self.ft_linked_server:
            return f'[{self.ft_linked_server}].[{self.ft_database}].[dbo].[{table_name}]'
        return f'[{self.ft_database}].[dbo].[{table_name}]'

    @staticmethod
    def _ft_table_names(months_back: int = 2) -> List[str]:
        """Return FUNDS_TRANSFER_YYYYMM names for current + prior N-1 months."""
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

    # ── DB connection ──────────────────────────────────────────────────────────

    def _get_conn(self, timeout=60):
        return pymssql.connect(
            server=self.config['db_server'],
            database=self.config['db_name'],
            user=self.config['db_user'],
            password=self.config['db_password'],
            timeout=timeout
        )

    # ── Core scheduler methods ─────────────────────────────────────────────────

    def get_pending_invoices(self) -> List[Dict]:
        """
        Fetch up to 500 invoices per month from FUNDS_TRANSFER tables that have
        not yet been posted (no matching row in dbo.RESPONSES). Skips WAIVE and
        zero-amount records.

        Pre-fetches RESPONSES refs locally to avoid a cross-linked-server NOT IN
        subquery that times out on the 22M-row production FT table.
        """
        conn = self._get_conn(timeout=120)
        cursor = conn.cursor(as_dict=True)
        all_rows: List[Dict] = []

        # Pull already-posted refs from local DB (fast, no cross-server join)
        local_cur = conn.cursor()
        local_cur.execute("SELECT TRANS_REF FROM [dbo].[RESPONSES]")
        posted_refs = {r[0] for r in local_cur.fetchall()}
        local_cur.close()
        logger.info(f"Already posted: {len(posted_refs)} invoices in RESPONSES")

        # Default date filter to current month to avoid full-table scans
        now = datetime.now()
        month_start = f"{now.year}{now.month:02d}01"

        for table_name in self._ft_table_names(2):
            ft_ref = self._ft_ref(table_name)
            try:
                cursor.execute(f"""
                    SELECT TOP 500
                        ft.RECID,            ft.TRAN_REFERENCE,   ft.TRAN_TYPE,
                        ft.DEBIT_ACCOUNT_NO, ft.DEBIT_CURRENCY,   ft.DEBIT_AMOUNT,
                        ft.DEBIT_VALUE_DATE, ft.CREDIT_ACCOUNT_NO,ft.CREDIT_CURRENCY,
                        ft.CREDIT_AMOUNT,    ft.COMMISSION_TYPE,  ft.COMMISSION_CODE,
                        ft.COMMISSION_AMOUNT,ft.CHARGE_CODE,      ft.TIN,
                        ft.CONTR_NAME,       ft.SUPPL_ADDR,       ft.INV_NO,
                        ft.DATED,            ft.DEBIT_CUSTOMER,   ft.CREDIT_CUSTOMER,
                        ft.ORDERING_CUSTOMER,ft.CO_CODE,          ft.PROCESSING_DATE
                    FROM {ft_ref} ft
                    WHERE ft.COMMISSION_AMOUNT IS NOT NULL
                      AND ft.COMMISSION_AMOUNT != ''
                      AND ft.COMMISSION_AMOUNT NOT LIKE 'WAIVE%'
                      AND ft.TRAN_REFERENCE IS NOT NULL
                      AND ft.DEBIT_VALUE_DATE >= '{month_start}'
                """)
                rows = cursor.fetchall()
                positive = 0
                for row in rows:
                    if row.get('TRAN_REFERENCE') in posted_refs:
                        continue
                    amt = self._parse_commission_amount(row.get('COMMISSION_AMOUNT', ''))
                    if amt > 0:
                        row['_parsed_amount'] = amt
                        all_rows.append(row)
                        positive += 1
                logger.info(f"{table_name}: {len(rows)} candidates, {positive} new with positive amount")
            except Exception as e:
                logger.warning(f"Could not query {table_name}: {e}")

        cursor.close()
        conn.close()
        logger.info(f"Total pending invoices to process: {len(all_rows)}")
        return all_rows

    def map_ft_to_payload(self, row: Dict) -> Dict:
        """Map a single FUNDS_TRANSFER row to the Cryptware NRS API payload."""
        trans_ref = self._clean(row.get('TRAN_REFERENCE', ''))

        # Issue date: DEBIT_VALUE_DATE is YYYYMMDD char; fall back to DATED datetime
        raw_date = str(row.get('DEBIT_VALUE_DATE') or row.get('PROCESSING_DATE') or '').replace('-', '')[:8]
        if len(raw_date) == 8 and raw_date.isdigit():
            issue_date = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:8]}"
        elif row.get('DATED') and hasattr(row['DATED'], 'strftime'):
            issue_date = row['DATED'].strftime('%Y-%m-%d')
        else:
            issue_date = datetime.now().strftime('%Y-%m-%d')

        party_name = (self._clean(row.get('CONTR_NAME'))
                      or self._clean(row.get('ORDERING_CUSTOMER'))
                      or self._clean(row.get('DEBIT_CUSTOMER'))
                      or 'Unknown')
        tin       = self._clean(row.get('TIN')) or ''
        street    = self._clean(row.get('SUPPL_ADDR')) or 'Unknown'
        currency  = self._clean(row.get('DEBIT_CURRENCY')) or 'NGN'
        acct_no   = self._clean(row.get('DEBIT_ACCOUNT_NO')) or str(row.get('RECID', '1'))

        amount        = row.get('_parsed_amount') or self._parse_commission_amount(row.get('COMMISSION_AMOUNT', ''))
        comm_type     = self._clean(row.get('COMMISSION_TYPE') or '')
        charge_code   = (self._clean(row.get('CHARGE_CODE'))
                         or self._clean(row.get('COMMISSION_CODE'))
                         or 'Bank Charge')
        is_vat        = 'VAT' in comm_type.upper()
        tax_rate      = 7.5 if is_vat else 0
        tax_cat       = 'STANDARD_VAT' if is_vat else 'ZERO_VAT'
        tx_category   = 'B2B' if tin else 'B2C'
        price_unit    = self._normalize_uom(comm_type, charge_code)

        return {
            "document_identifier":    trans_ref,
            "invoice_type":           "STANDARD",
            "issue_date":             issue_date,
            "due_date":               issue_date,
            "invoice_type_code":      "381",
            "document_currency_code": currency,
            "transaction_category":   tx_category,
            "accounting_customer_party": {
                "party_name":          party_name,
                "email":               "noreply@na.ng",
                "telephone":           self._format_phone(''),
                "tin":                 tin or "000000000000",
                "business_description": party_name,
                "postal_address": {
                    "street_name": street,
                    "city_name":   "Unknown",
                    "postal_zone": "100001",
                    "country":     "NG"
                }
            },
            "invoice_lines": [{
                "internal_id":       acct_no[:50] or "1",
                "description":       charge_code[:100],
                "invoiced_quantity": 1,
                "price_amount":      round(float(amount), 2),
                "hsn_code":          "6499.00",
                "price_unit":        price_unit,
                "product_category":  "Financial Services",
                "tax_rate":          tax_rate,
                "tax_category_id":   tax_cat,
                "discount_rate":     0
            }]
        }

    def submit_invoice(self, payload: Dict) -> tuple:
        """Submit invoice to Cryptware API. Returns (status_code, response_dict)."""
        try:
            headers = {
                'Content-Type': 'application/json',
                'X-Api-Key': self.config['api_key']
            }
            url = f"{self.config['base_url']}/invoice/generate"
            logger.info(f"Submitting {payload['document_identifier']}")
            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            try:
                resp_data = resp.json()
            except Exception:
                resp_data = {"error": resp.text}
            logger.info(f"API → {payload['document_identifier']}: HTTP {resp.status_code}")
            return resp.status_code, resp_data
        except requests.exceptions.RequestException as e:
            logger.error(f"API call failed: {e}")
            return 500, {"error": str(e)}

    def write_response(self, trans_ref: str, booking_date_raw,
                       status_code: int, response_data: Dict,
                       ft_row: Dict = None):
        """
        Insert or update dbo.RESPONSES with the API response and optional
        display fields sourced from the originating FUNDS_TRANSFER row.
        """
        # Normalise booking_date → char(8) YYYYMMDD
        if booking_date_raw is None:
            booking_date = datetime.now().strftime('%Y%m%d')
        elif hasattr(booking_date_raw, 'strftime'):
            booking_date = booking_date_raw.strftime('%Y%m%d')
        else:
            bd = str(booking_date_raw).replace('-', '').replace(' ', '')[:8]
            booking_date = bd if (len(bd) == 8 and bd.isdigit()) else datetime.now().strftime('%Y%m%d')

        # Parse API response
        irn = qr_code = response_code = error_message = error_detail = ''
        if status_code in (200, 201):
            data = response_data.get('data', {})
            irn         = data.get('irn', '') or ''
            qr_code     = data.get('qr_code_url', '') or data.get('qr_code', '') or ''
            response_code = 'SUCCESS'
        else:
            error_message = (response_data.get('message', '') or str(response_data))[:2000]
            errors = response_data.get('errors') or response_data.get('data', {}).get('errors')
            if errors:
                error_detail = json.dumps(errors)[:4000]
            response_code = 'ERROR'

        response_json = json.dumps(response_data)
        if len(response_json) > 4000:
            response_json = response_json[:4000]

        # Extended display fields from the FT row
        customer_name = currency = charge_desc = debit_account = comm_type = None
        amount_val = None
        if ft_row:
            cn = (self._clean(ft_row.get('CONTR_NAME'))
                  or self._clean(ft_row.get('ORDERING_CUSTOMER'))
                  or self._clean(ft_row.get('DEBIT_CUSTOMER')))
            customer_name = cn[:200] if cn else None
            cur = self._clean(ft_row.get('DEBIT_CURRENCY')) or 'NGN'
            currency = cur[:10]
            amount_val = ft_row.get('_parsed_amount') or self._parse_commission_amount(ft_row.get('COMMISSION_AMOUNT', ''))
            cd = (self._clean(ft_row.get('CHARGE_CODE'))
                  or self._clean(ft_row.get('COMMISSION_CODE')))
            charge_desc   = cd[:500] if cd else None
            da = self._clean(ft_row.get('DEBIT_ACCOUNT_NO'))
            debit_account = da[:50] if da else None
            ct = self._clean(ft_row.get('COMMISSION_TYPE'))
            comm_type     = ct[:100] if ct else None

        try:
            conn = self._get_conn()
            cur = conn.cursor()

            cur.execute("SELECT 1 FROM [dbo].[RESPONSES] WHERE TRANS_REF = %s", (trans_ref,))
            exists = cur.fetchone() is not None

            if exists:
                cur.execute("""
                    UPDATE [dbo].[RESPONSES]
                    SET HTTP_STATUS=%s, RESPONSE_JSON=%s, IRN=%s, RESPONSE_CODE=%s,
                        ERROR_MESSAGE=%s, ERROR_DETAIL=%s, QR_CODE=%s,
                        ENVIRONMENT=%s, LAST_UPDATED=GETDATE(),
                        CUSTOMER_NAME=ISNULL(%s, CUSTOMER_NAME),
                        CURRENCY=ISNULL(%s, CURRENCY),
                        AMOUNT=ISNULL(%s, AMOUNT),
                        CHARGE_DESCRIPTION=ISNULL(%s, CHARGE_DESCRIPTION),
                        DEBIT_ACCOUNT_NO=ISNULL(%s, DEBIT_ACCOUNT_NO),
                        COMMISSION_TYPE=ISNULL(%s, COMMISSION_TYPE)
                    WHERE TRANS_REF=%s
                """, (status_code, response_json, irn, response_code,
                      error_message, error_detail, qr_code, self.environment,
                      customer_name, currency, amount_val, charge_desc, debit_account, comm_type,
                      trans_ref))
            else:
                cur.execute("""
                    INSERT INTO [dbo].[RESPONSES]
                    (BOOKING_DATE, TRANS_REF, ENVIRONMENT, HTTP_STATUS, RESPONSE_JSON,
                     IRN, RESPONSE_CODE, ERROR_MESSAGE, ERROR_DETAIL, QR_CODE, LAST_UPDATED,
                     CUSTOMER_NAME, CURRENCY, AMOUNT, CHARGE_DESCRIPTION, DEBIT_ACCOUNT_NO,
                     COMMISSION_TYPE)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, GETDATE(),
                            %s, %s, %s, %s, %s, %s)
                """, (booking_date, trans_ref, self.environment, status_code, response_json,
                      irn, response_code, error_message, error_detail, qr_code,
                      customer_name, currency, amount_val, charge_desc, debit_account, comm_type))

            conn.commit()
            cur.close()
            conn.close()

            if status_code in (200, 201):
                logger.info(f"Response written for {trans_ref}: IRN={irn}")
            else:
                logger.warning(f"Failed response for {trans_ref}: HTTP {status_code} — {error_message[:80]}")

        except Exception as e:
            logger.error(f"Error writing response for {trans_ref}: {e}")
            raise

    def process_invoices(self):
        """Main processing function — called by the APScheduler job."""
        logger.info("=" * 80)
        logger.info("Starting invoice processing job")
        try:
            pending = self.get_pending_invoices()
            if not pending:
                logger.info("No pending invoices to process")
                return
            logger.info(f"Processing {len(pending)} invoices")
            for row in pending:
                trans_ref = row.get('TRAN_REFERENCE', '')
                if not trans_ref:
                    continue
                try:
                    payload      = self.map_ft_to_payload(row)
                    status_code, response_data = self.submit_invoice(payload)
                    booking_date = (row.get('DEBIT_VALUE_DATE')
                                    or row.get('PROCESSING_DATE')
                                    or row.get('DATED'))
                    self.write_response(trans_ref, booking_date, status_code, response_data, row)
                except Exception as e:
                    logger.error(f"Error processing invoice {trans_ref}: {e}")
            logger.info("Invoice processing job completed")
        except Exception as e:
            logger.error(f"Error in invoice processing job: {e}")


def load_config(config_path: str = 'config.json') -> Dict:
    """Load configuration from JSON file, falling back to environment variables."""
    config = {
        'db_server': None, 'db_name': None, 'db_user': None, 'db_password': None,
        'base_url': None, 'api_key': None, 'interval_minutes': 5,
        'ft_linked_server': None, 'ft_database': 'Transactions',
        'log_dir': 'logs',
    }

    cfg_file = Path(config_path)
    if cfg_file.exists():
        try:
            with open(cfg_file, 'r') as f:
                fc = json.load(f)
            if 'database' in fc:
                config['db_server']   = fc['database'].get('server')
                config['db_name']     = fc['database'].get('name')
                config['db_user']     = fc['database'].get('user')
                config['db_password'] = fc['database'].get('password')
            if 'api' in fc:
                config['base_url'] = fc['api'].get('base_url')
                config['api_key']  = fc['api'].get('api_key')
            if 'scheduler' in fc:
                config['interval_minutes'] = fc['scheduler'].get('interval_minutes', 5)
            if 'fund_transfer' in fc:
                config['ft_linked_server'] = fc['fund_transfer'].get('linked_server')
                config['ft_database']      = fc['fund_transfer'].get('database', 'Transactions')
            if 'logging' in fc:
                config['log_dir'] = fc['logging'].get('log_dir', 'logs')
            logger.info(f"Loaded configuration from {config_path}")
        except Exception as e:
            logger.warning(f"Could not load config file {config_path}: {e}")
    else:
        logger.warning(f"Config file {config_path} not found")

    # Environment-variable overrides
    env_map = {
        'DB_SERVER': 'db_server', 'DB_NAME': 'db_name',
        'DB_USER': 'db_user', 'DB_PASSWORD': 'db_password',
        'BASE_URL': 'base_url', 'API_KEY': 'api_key',
        'INTERVAL_MINUTES': 'interval_minutes',
        'FT_LINKED_SERVER': 'ft_linked_server',
        'FT_DATABASE': 'ft_database',
        'LOG_DIR': 'log_dir',
    }
    for env_var, cfg_key in env_map.items():
        val = os.environ.get(env_var)
        if val:
            config[cfg_key] = int(val) if cfg_key == 'interval_minutes' else val

    required = ['db_server', 'db_name', 'db_user', 'db_password', 'base_url', 'api_key']
    missing  = [f for f in required if not config[f]]
    if missing:
        logger.error(f"Missing required configuration: {', '.join(missing)}")
        sys.exit(1)

    return config


def main():
    config = load_config('config.json')
    setup_logging(log_dir=config.get('log_dir', 'logs'), worker='invoice_scheduler')
    sched_inst = InvoiceScheduler(config)

    scheduler = BlockingScheduler()
    interval  = config.get('interval_minutes', 5)
    scheduler.add_job(
        sched_inst.process_invoices,
        'interval', minutes=interval,
        id='invoice_processing_job',
        name='Process pending FUNDS_TRANSFER invoices'
    )
    logger.info(f"Invoice scheduler started — running every {interval} minutes")
    logger.info("Press Ctrl+C to exit")
    try:
        sched_inst.process_invoices()   # run immediately on startup
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        logger.info("Scheduler stopped")


if __name__ == "__main__":
    main()
