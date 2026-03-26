import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import re
import numpy as np

# ────────────────────────────────────────────────
#  GLOBAL CONSTANTS
# ────────────────────────────────────────────────

BALANCE_ORDER = ["0-49999.99", "50000.00-99999.99", "100000.00 and up"]

EXCLUDED_SUBSTATUSES = [
    "BUSY TONE", "NIS/OOCA", "NO ANSWER", "NO ANSWER_SENT PAYMENT REMINDER",
    "ADC", "NA", "AB", "BUSY_OOCA", "NIS", "BP",
    "LETTER SENT - MANUAL AGENT SMS", "LETTER SENT - MANUAL AGENT EMAIL",
    "KEEPS ON RINGING", "BUSY", "NEGATIVE",
    "MANUAL AGENT EMAIL", "EMAIL", "CALL BARRED"
]

SUMMARY_COLUMNS = [
    'CYCLE', 'DATE', 'CLIENT', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE',
    'CONNECTED NU', 'CONNECTED UNIQUE',
    'TOTAL RPC', 'PTP',
    'CONNECTED % NU', 'CONNECTED % UNIQUE',
    'RPC %', 'PTP %', 'TOTAL BALANCE',
    'NEG DROP', 'SYSTEM DROP', 'CALL DROP RATE'
]

CYCLE_COLUMN_BY_BANK = {
    "BPI AUTO CURING SL":           "dpd",
    "BPI BANKO SL":                 "cycle",
    "BPI CARDS 30DPD SL":           "cycle",
    "BPI CARDS XDAYS SL":           "level",
    "BPI PL 30DPD SL":              "cycle",
    "BPI PL 60DPD SL":              "cycle",
    "BPI RBANK 30DPD CARDS SL":     "cycle",
    "BPI RBANK INSTABALE PL SL":    "cycle",
    "BPI RBANK PL 30DPD SL":        "dpd",
    "BPI RBANK PL 60DPD SL":        "dpd",
    "BPI PL XDAYS SL":              "cycle",
}

st.set_page_config(layout="wide", page_title="Debt Collection Summary(Texxen)",
                   page_icon="📊", initial_sidebar_state="expanded")

st.title('Debt Collection Summary')

# ────────────────────────────────────────────────
#  HELPERS
# ────────────────────────────────────────────────

def normalize_cycle(val):
    if pd.isna(val):
        return "Unknown"
    s = str(val).strip().upper()
    s = re.sub(r'^(LEVEL|LVL|CYCLE|BUCKET|AGE)\s*', '', s, flags=re.I)
    s = re.sub(r'[^0-9]', '', s)
    if not s:
        return "Unknown"
    try:
        return str(int(s))
    except ValueError:
        return "Unknown"


def get_cycle_value(row):
    bank_raw = row.get('bankname', None)
    if pd.isna(bank_raw) or not isinstance(bank_raw, str):
        bank_key = ""
    else:
        bank_key = bank_raw.strip().upper()

    if bank_key in CYCLE_COLUMN_BY_BANK:
        col = CYCLE_COLUMN_BY_BANK[bank_key]
        if col in row and pd.notna(row[col]) and str(row[col]).strip():
            return row[col]

    for col in ['cycle', 'level', 'dpd']:
        if col in row:
            val = row[col]
            if pd.notna(val) and str(val).strip():
                return val

    return None


# ────────────────────────────────────────────────
#  DATA LOADING
# ────────────────────────────────────────────────

@st.cache_data
def load_and_filter_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.lower()

    # Standardize 'chcode'
    possible = ['chcode', 'ch_code', 'account', 'accountno', 'acctno',
                'debtorid', 'debtor_id', 'contractno', 'cardno', 'customerid']
    for col in possible:
        if col in df.columns:
            df = df.rename(columns={col: 'chcode'})
            break

    if 'chcode' not in df.columns:
        st.error(f"No chcode/account column in {uploaded_file.name}")
        return pd.DataFrame()

    if 'contactsource' in df.columns:
        df = df[df['contactsource'].astype(str).str.contains("CALL", case=False, na=False)]

    exclude = ["PULLOUT", "LETTER RECEIVED", "LETTER SENT", "NEW ENDO"]
    if 'status' in df.columns:
        df = df[~df['status'].astype(str).str.upper().isin(exclude)]

    df['cycle'] = df.apply(get_cycle_value, axis=1).apply(normalize_cycle).fillna("Unknown")

    unknown_rate = (df['cycle'] == "Unknown").mean()
    if unknown_rate > 0.08:
        banks = df[df['cycle'] == "Unknown"]['bankname'].dropna().unique()
        msg = f"⚠️ {unknown_rate:.1%} Unknown cycle\n"
        if len(banks):
            msg += "Sample banks:\n" + "\n".join(banks[:10])
        st.warning(msg)

    return df


# ────────────────────────────────────────────────
#  EXCEL FORMATS
# ────────────────────────────────────────────────

def setup_excel_formats(workbook):
    return {
        'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                      'valign': 'vcenter', 'bg_color': '#A42A25', 'font_color': 'white'}),
        'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'header': workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                       'bg_color': '#D9B229', 'font_color': 'white', 'bold': True}),
        'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                      'border': 1, 'num_format': '#,##0'}),
        'percent': workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                        'border': 1, 'num_format': '0.00%'}),
        'date': workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                     'border': 1, 'num_format': 'yyyy-mm-dd'}),
    }


def write_excel_sheet(writer, sheet_name, df_dict, formats):
    if not df_dict:
        return
    worksheet = writer.book.add_worksheet(sheet_name)
    current_row = 0

    sorted_items = []
    if sheet_name == 'Overall Combined Summary':
        main = next((k for k in df_dict if 'Total' in k), None)
        if main:
            sorted_items.append((main, df_dict[main]))
        for k, v in df_dict.items():
            if k != main:
                sorted_items.append((k, v))
    elif sheet_name == 'Predictive Cycles (Total)':
        cycle_keys = [k for k in df_dict if k.startswith('Cycle ')]
        cycle_keys.sort(key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1)) if re.search(r'Cycle (\d+)', x) else 9999)
        for ck in cycle_keys:
            sorted_items.append((ck, df_dict[ck]))
    elif sheet_name.endswith('Cycles'):
        sorted_items = sorted(
            df_dict.items(),
            key=lambda x: int(re.search(r'Cycle (\d+)', x[0]).group(1)) if re.search(r'Cycle (\d+)', x[0]) else 9999
        )
    elif sheet_name.endswith('Balances'):
        sorted_items = sorted(
            df_dict.items(),
            key=lambda x: (
                int(re.search(r'Cycle (\d+)', x[0]).group(1)) if re.search(r'Cycle (\d+)', x[0]) else 9999,
                BALANCE_ORDER.index(re.search(r'Balance (.+)$', x[0]).group(1)) if re.search(r'Balance (.+)$', x[0]) else 999
            )
        )
    else:
        sorted_items = list(df_dict.items())

    for title, df in sorted_items:
        if df.empty:
            continue
        df_display = df.copy()

        worksheet.merge_range(current_row, 0, current_row, len(df_display.columns)-1, title, formats['title'])
        current_row += 1

        for col_num, col_name in enumerate(df_display.columns):
            worksheet.write(current_row, col_num, col_name, formats['header'])
            max_len = max(df_display[col_name].astype(str).str.len().max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, max_len)
        current_row += 1

        for r in range(len(df_display)):
            for c, col in enumerate(df_display.columns):
                val = df_display.iloc[r, c]
                if col == 'DATE':
                    try:
                        dt = pd.to_datetime(val).to_pydatetime()
                        worksheet.write_datetime(current_row + r, c, dt, formats['date'])
                    except:
                        worksheet.write_string(current_row + r, c, str(val), formats['center'])
                elif col in ['TOTAL BALANCE', 'CONNECTED NU', 'CONNECTED UNIQUE',
                             'ACCOUNTS', 'TOTAL DIALED', 'TOTAL RPC', 'PTP',
                             'NEG DROP', 'SYSTEM DROP']:
                    worksheet.write_number(current_row + r, c, float(val), formats['comma'])
                elif col in ['PENETRATION RATE', 'RPC %', 'PTP %', 'CALL DROP RATE',
                             'CONNECTED % NU', 'CONNECTED % UNIQUE']:
                    try:
                        worksheet.write_number(current_row + r, c, float(val.strip('%'))/100, formats['percent'])
                    except:
                        worksheet.write_string(current_row + r, c, str(val), formats['center'])
                else:
                    worksheet.write(current_row + r, c, str(val), formats['center'])
        current_row += len(df_display) + 2


def to_excel(summary_groups):
    output = BytesIO()
    order = ['Combined', 'Overall Combined Summary', 'Predictive', 'Manual',
             'Combined Cycles', 'Predictive Cycles', 'Manual Cycles', 'Predictive Cycles (Total)',
             'Combined Balances', 'Predictive Balances', 'Manual Balances']
    ordered = {k: summary_groups[k] for k in order if k in summary_groups}

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        fmts = setup_excel_formats(writer.book)
        for sheet_name, d in ordered.items():
            write_excel_sheet(writer, sheet_name, d, fmts)
    return output.getvalue()


# ────────────────────────────────────────────────
#  RAW CYCLE BREAKDOWN – NOW BASED ON OVERALL COMBINED (Predictive + Manual)
# ────────────────────────────────────────────────

def create_raw_cycle_breakdown_excel(raw_combined_df):
    if raw_combined_df.empty:
        return BytesIO().getvalue()

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book

        bold   = wb.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
        head   = wb.add_format({'bold': True, 'bg_color': '#E6F0FA', 'border': 1, 'align': 'center'})
        numf   = wb.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
        cent   = wb.add_format({'align': 'center', 'border': 1})
        sect   = wb.add_format({'bold': True, 'font_size': 12, 'align': 'left'})

        def list_codes(ws, title, series, start_row, col_idx, unique=True):
            if len(series) == 0:
                ws.write(start_row, col_idx, f"{title}: (none)", sect)
                return start_row + 1
            ws.write(start_row, col_idx, title, sect)
            start_row += 1
            vals = pd.Series(series).dropna().astype(str)
            if unique:
                vals = vals.unique()
            for i, code in enumerate(sorted(vals)):
                ws.write(start_row + i, col_idx, code, cent)
            return start_row + len(vals) + 1

        df = raw_combined_df.copy()   # ← Now using Combined (Predictive + Manual)

        # ── SHEET 1: Combined Cycles Raw ───────────────────────────────────
        ws_cycles = wb.add_worksheet("Combined Cycles Raw")
        row = 0

        sorted_cycles = sorted(
            [c for c in df['cycle'].unique() if c.isdigit()],
            key=int
        )

        for cycle_str in sorted_cycles:
            cycle_raw = df[df['cycle'] == cycle_str]
            if cycle_raw.empty:
                continue

            excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
            has_valid_substatus = (
                cycle_raw['substatus'].notna() &
                (cycle_raw['substatus'].astype(str).str.strip() != '')
            )
            not_excluded = ~cycle_raw['substatus'].astype(str).str.upper().isin(excluded_upper)
            connected_mask = has_valid_substatus & not_excluded

            summary = compute_group_metrics(cycle_raw)
            summary['CYCLE'] = cycle_str
            summary['DATE']  = 'Cycle Total'
            summary['CLIENT'] = 'All'

            ws_cycles.write(row, 0, f"Cycle {cycle_str}", bold)
            row += 2

            headers = [
                'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                'TOTAL RPC', 'PTP', 'NEG DROP', 'SYSTEM DROP'
            ]
            for col_idx, h in enumerate(headers):
                ws_cycles.write(row, col_idx, h, head)
            row += 1

            values = [
                summary.get('ACCOUNTS', 0),
                summary.get('TOTAL DIALED', 0),
                summary.get('CONNECTED NU', 0),
                summary.get('CONNECTED UNIQUE', 0),
                summary.get('TOTAL RPC', 0),
                summary.get('PTP', 0),
                summary.get('NEG DROP', 0),
                summary.get('SYSTEM DROP', 0),
            ]
            for col_idx, val in enumerate(values):
                ws_cycles.write_number(row, col_idx, val, numf)

            row += 2

            list_start_row = row

            list_codes(ws_cycles, "Raw Accounts (unique CHCODE)", cycle_raw['chcode'], list_start_row, 0)
            list_codes(ws_cycles, "Raw Total Dialed (all attempts CHCODE)", cycle_raw['chcode'], list_start_row, 1, unique=False)
            list_codes(ws_cycles, "Raw Connected NU (all connected attempts CHCODE)", cycle_raw[connected_mask]['chcode'], list_start_row, 2, unique=False)
            list_codes(ws_cycles, "Raw Connected Unique (unique CHCODE connected)", cycle_raw[connected_mask]['chcode'].unique(), list_start_row, 3)
            rpc_ptp_codes = cycle_raw[cycle_raw['groupstatus'].str.contains('rpc|ptp', case=False, na=False)]['chcode'].unique()
            list_codes(ws_cycles, "Raw Total RPC (unique CHCODE rpc/ptp)", rpc_ptp_codes, list_start_row, 4)
            ptp_codes = cycle_raw[cycle_raw['groupstatus'].str.contains('ptp', case=False, na=False)]['chcode'].unique()
            list_codes(ws_cycles, "Raw PTP (unique CHCODE ptp)", ptp_codes, list_start_row, 5)
            neg_codes = cycle_raw[cycle_raw['substatus'].str.contains('call drop', case=False, na=False)]['chcode'].unique()
            list_codes(ws_cycles, "Raw Negative Drop (unique CHCODE)", neg_codes, list_start_row, 6)
            drop_codes = cycle_raw[cycle_raw['substatus'].str.contains('pdrop', case=False, na=False)]['chcode'].unique()
            list_codes(ws_cycles, "Raw System Drop (unique CHCODE)", drop_codes, list_start_row, 7)

            max_list_len = max(
                len(cycle_raw['chcode'].unique()),
                len(cycle_raw['chcode']),
                len(cycle_raw[connected_mask]['chcode']),
                len(cycle_raw[connected_mask]['chcode'].unique()),
                len(rpc_ptp_codes),
                len(ptp_codes),
                len(neg_codes),
                len(drop_codes)
            )
            row = list_start_row + max_list_len + 6

        # ── SHEET 2: Overall Total Summary ───────────────────────────────────
        ws_total = wb.add_worksheet("Overall Total Summary")
        row = 0

        excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
        has_valid_substatus = (
            df['substatus'].notna() &
            (df['substatus'].astype(str).str.strip() != '')
        )
        not_excluded = ~df['substatus'].astype(str).str.upper().isin(excluded_upper)
        connected_mask = has_valid_substatus & not_excluded

        all_summary = compute_group_metrics(df)

        ws_total.write(row, 0, "Overall Combined Summary (Total Period)", bold)
        row += 2

        headers = [
            'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
            'TOTAL RPC', 'PTP', 'NEG DROP', 'SYSTEM DROP'
        ]
        for col_idx, h in enumerate(headers):
            ws_total.write(row, col_idx, h, head)
        row += 1

        values = [
            all_summary.get('ACCOUNTS', 0),
            all_summary.get('TOTAL DIALED', 0),
            all_summary.get('CONNECTED NU', 0),
            all_summary.get('CONNECTED UNIQUE', 0),
            all_summary.get('TOTAL RPC', 0),
            all_summary.get('PTP', 0),
            all_summary.get('NEG DROP', 0),
            all_summary.get('SYSTEM DROP', 0),
        ]
        for col_idx, val in enumerate(values):
            ws_total.write_number(row, col_idx, val, numf)

        row += 2
        list_start_row = row

        list_codes(ws_total, "Raw Accounts (unique CHCODE)", df['chcode'], list_start_row, 0)
        list_codes(ws_total, "Raw Total Dialed (all attempts CHCODE)", df['chcode'], list_start_row, 1, unique=False)
        list_codes(ws_total, "Raw Connected NU (all connected attempts CHCODE)", df[connected_mask]['chcode'], list_start_row, 2, unique=False)
        list_codes(ws_total, "Raw Connected Unique (unique CHCODE connected)", df[connected_mask]['chcode'].unique(), list_start_row, 3)
        rpc_ptp_codes = df[df['groupstatus'].str.contains('rpc|ptp', case=False, na=False)]['chcode'].unique()
        list_codes(ws_total, "Raw Total RPC (unique CHCODE rpc/ptp)", rpc_ptp_codes, list_start_row, 4)
        ptp_codes = df[df['groupstatus'].str.contains('ptp', case=False, na=False)]['chcode'].unique()
        list_codes(ws_total, "Raw PTP (unique CHCODE ptp)", ptp_codes, list_start_row, 5)
        neg_codes = df[df['substatus'].str.contains('call drop', case=False, na=False)]['chcode'].unique()
        list_codes(ws_total, "Raw Negative Drop (unique CHCODE)", neg_codes, list_start_row, 6)
        drop_codes = df[df['substatus'].str.contains('pdrop', case=False, na=False)]['chcode'].unique()
        list_codes(ws_total, "Raw System Drop (unique CHCODE)", drop_codes, list_start_row, 7)

    return output.getvalue()


# ────────────────────────────────────────────────
#  METRIC CALCULATION (unchanged)
# ────────────────────────────────────────────────

def compute_group_metrics(g):
    if g.empty:
        return pd.Series({c: 0 for c in SUMMARY_COLUMNS if c not in ['CYCLE','DATE','CLIENT']})

    accounts      = g['chcode'].nunique()
    total_dialed  = len(g)

    excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
    has_valid_substatus = (
        g['substatus'].notna() &
        (g['substatus'].astype(str).str.strip() != '')
    )
    not_excluded = ~g['substatus'].astype(str).str.upper().isin(excluded_upper)
    connected_mask = has_valid_substatus & not_excluded

    connected_nu     = len(g[connected_mask])
    connected_unique = g[connected_mask]['chcode'].nunique()

    rpc_ptp_mask  = g['groupstatus'].str.contains('rpc|ptp', case=False, na=False)
    total_rpc     = g[rpc_ptp_mask]['chcode'].nunique()

    ptp_mask      = g['groupstatus'].str.contains('ptp', case=False, na=False)
    ptp           = g[ptp_mask]['chcode'].nunique()

    ptp_accounts_ob = g[ptp_mask].groupby('chcode')['ob'].first()
    total_balance   = ptp_accounts_ob.sum()

    neg_drop_mask = g['substatus'].str.contains('call drop', case=False, na=False)
    neg_drop      = g[neg_drop_mask]['chcode'].nunique()

    sys_drop_mask = g['substatus'].str.contains('pdrop', case=False, na=False)
    system_drop   = g[sys_drop_mask]['chcode'].nunique()

    pen_rate         = f"{(total_dialed / accounts * 100):.2f}%" if accounts > 0 else "0.00%"
    connected_pct_nu = f"{(connected_nu / accounts * 100):.2f}%" if accounts > 0 else "0.00%"
    connected_pct_u  = f"{(connected_unique / accounts * 100):.2f}%" if accounts > 0 else "0.00%"

    rpc_pct          = f"{(total_rpc / connected_unique * 100):.2f}%" if connected_unique > 0 else "0.00%"

    ptp_pct          = f"{(ptp / total_rpc * 100):.2f}%"          if total_rpc > 0 else "0.00%"
    drop_rate        = f"{((neg_drop + system_drop) / accounts * 100):.2f}%" if accounts > 0 else "0.00%"

    return pd.Series({
        'ACCOUNTS':          accounts,
        'TOTAL DIALED':      total_dialed,
        'PENETRATION RATE':  pen_rate,
        'CONNECTED NU':      connected_nu,
        'CONNECTED UNIQUE':  connected_unique,
        'TOTAL RPC':         total_rpc,
        'PTP':               ptp,
        'CONNECTED % NU':    connected_pct_nu,
        'CONNECTED % UNIQUE': connected_pct_u,
        'RPC %':             rpc_pct,
        'PTP %':             ptp_pct,
        'TOTAL BALANCE':     total_balance,
        'NEG DROP':          neg_drop,
        'SYSTEM DROP':       system_drop,
        'CALL DROP RATE':    drop_rate,
    })


def build_daily_summary(df):
    if df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    df = df.copy()
    df['date']   = pd.to_datetime(df['barcodedate'], errors='coerce').dt.date
    df['client'] = df['bankname'].astype(str).str.strip()

    agg = df.groupby(['date', 'client'], dropna=False).apply(compute_group_metrics).reset_index()

    agg = agg.rename(columns={'date':'DATE', 'client':'CLIENT'})
    agg['CYCLE'] = ''

    agg = agg.reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    agg = agg.sort_values(['DATE', 'CLIENT'])

    return agg


def build_total_row(df, title_suffix=""):
    if df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    m = compute_group_metrics(df)
    m['CYCLE'] = 'All'
    m['DATE']  = 'Overall' + (f" {title_suffix}" if title_suffix else "")
    m['CLIENT'] = 'All Clients'
    return pd.DataFrame([m])[SUMMARY_COLUMNS]


def get_cycle_summaries(df):
    result = {}
    df = df.copy()

    for cyc in sorted(df['cycle'].unique(), key=lambda x: int(x) if x.isdigit() else 9999):
        if not cyc or cyc == "Unknown":
            continue
        sub = df[df['cycle'] == cyc]
        if sub.empty:
            continue
        s = build_daily_summary(sub)
        s['CYCLE'] = cyc
        result[f"Cycle {cyc}"] = s
    return result


def get_balance_summaries(df):
    result = {}
    df = df.copy()

    ranges = [(0, 49999.99, "0-49999.99"), (50000, 99999.99, "50000.00-99999.99"), (100000, float('inf'), "100000.00 and up")]
    for min_b, max_b, name in ranges:
        b_df = df[(df['ob'] >= min_b) & (df['ob'] <= max_b)]
        if b_df.empty:
            continue
        for cyc in sorted(b_df['cycle'].unique(), key=lambda x: int(x) if x.isdigit() else 9999):
            if not cyc or cyc == "Unknown":
                continue
            sub = b_df[b_df['cycle'] == cyc]
            if sub.empty:
                continue
            s = build_daily_summary(sub)
            s['CYCLE'] = cyc
            result[f"Cycle {cyc} Balance {name}"] = s
    return result


def combine_summaries(pred_dict, man_dict):
    combined = {}
    all_keys = set(pred_dict.keys()) | set(man_dict.keys())

    for key in sorted(all_keys, key=lambda k: int(re.search(r'Cycle (\d+)', k).group(1)) if re.search(r'Cycle (\d+)', k) else 9999):
        pred = pred_dict.get(key, pd.DataFrame(columns=SUMMARY_COLUMNS))
        man  = man_dict.get(key, pd.DataFrame(columns=SUMMARY_COLUMNS))

        if pred.empty and man.empty:
            continue

        df = pd.concat([pred, man], ignore_index=True)
        if df.empty:
            continue

        if 'contactsource' in df.columns:
            df = df[df['contactsource'].astype(str).str.contains("CALL", case=False, na=False)]

        df['DATE'] = pd.to_datetime(df['DATE']).dt.date

        num_cols = [
            'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
            'TOTAL RPC', 'PTP', 'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP'
        ]
        df[num_cols] = df[num_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

        agg = df.groupby(['DATE', 'CLIENT'], as_index=False)[num_cols].sum()

        agg['PENETRATION RATE'] = agg.apply(
            lambda r: f"{(r['TOTAL DIALED'] / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)
        agg['CONNECTED % NU'] = agg.apply(
            lambda r: f"{(r['CONNECTED NU'] / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)
        agg['CONNECTED % UNIQUE'] = agg.apply(
            lambda r: f"{(r['CONNECTED UNIQUE'] / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)

        agg['RPC %'] = agg.apply(
            lambda r: f"{(r['TOTAL RPC'] / r['CONNECTED UNIQUE'] * 100):.2f}%" if r['CONNECTED UNIQUE'] > 0 else "0.00%", axis=1)

        agg['PTP %'] = agg.apply(
            lambda r: f"{(r['PTP'] / r['TOTAL RPC'] * 100):.2f}%" if r['TOTAL RPC'] > 0 else "0.00%", axis=1)
        agg['CALL DROP RATE'] = agg.apply(
            lambda r: f"{((r['NEG DROP'] + r['SYSTEM DROP']) / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)

        agg['CYCLE'] = key.replace("Cycle ", "") if "Cycle " in key else ""

        combined[key] = agg.reindex(columns=SUMMARY_COLUMNS, fill_value=0).sort_values('DATE')

    return combined


# ────────────────────────────────────────────────
#  MAIN APP
# ────────────────────────────────────────────────

uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_combined, all_predictive, all_manual = [], [], []
    predictive_cycles, manual_cycles = {}, {}
    predictive_balances, manual_balances = {}, {}
    raw_combined_all = pd.DataFrame()          # ← Changed to combined

    prog = st.progress(0)
    for idx, f in enumerate(uploaded_files):
        with st.spinner(f"Processing {f.name}..."):
            df_clean = load_and_filter_data(f)

            disp = df_clean.get('dispositiontype', pd.Series()).astype(str).str.upper()

            df_pred = df_clean[disp == 'PREDICTIVE'].copy()
            df_man  = df_clean[disp == 'MANUAL'].copy()
            df_comb = pd.concat([df_pred, df_man], ignore_index=True)

            # Use combined for raw breakdown
            raw_combined_all = pd.concat([raw_combined_all, df_comb], ignore_index=True)

            all_combined.append(build_daily_summary(df_comb))
            all_predictive.append(build_daily_summary(df_pred))
            all_manual.append(build_daily_summary(df_man))

            for k, v in get_cycle_summaries(df_pred).items():
                predictive_cycles[k] = pd.concat([predictive_cycles.get(k, pd.DataFrame()), v], ignore_index=True)
            for k, v in get_cycle_summaries(df_man).items():
                manual_cycles[k] = pd.concat([manual_cycles.get(k, pd.DataFrame()), v], ignore_index=True)

            for k, v in get_balance_summaries(df_pred).items():
                predictive_balances[k] = pd.concat([predictive_balances.get(k, pd.DataFrame()), v], ignore_index=True)
            for k, v in get_balance_summaries(df_man).items():
                manual_balances[k] = pd.concat([manual_balances.get(k, pd.DataFrame()), v], ignore_index=True)

        prog.progress((idx + 1) / len(uploaded_files))

    prog.empty()
    st.success(f"Processed {len(uploaded_files)} file(s)")

    combined_summary = pd.concat(all_combined, ignore_index=True).sort_values('DATE') if all_combined else pd.DataFrame(columns=SUMMARY_COLUMNS)
    predictive_summary = pd.concat(all_predictive, ignore_index=True).sort_values('DATE') if all_predictive else pd.DataFrame(columns=SUMMARY_COLUMNS)
    manual_summary = pd.concat(all_manual, ignore_index=True).sort_values('DATE') if all_manual else pd.DataFrame(columns=SUMMARY_COLUMNS)

    combined_cycle   = combine_summaries(predictive_cycles, manual_cycles)
    combined_balance = combine_summaries(predictive_balances, manual_balances)

    overall_combined_total = build_total_row(raw_combined_all)

    predictive_cycles_total = {}
    unique_cycles = set()
    for key in predictive_cycles:
        match = re.search(r'Cycle (\d+)', key)
        if match:
            unique_cycles.add(match.group(1))
    for cycle in sorted(list(unique_cycles), key=int):
        cycle_key = f"Cycle {cycle}"
        raw = raw_combined_all[raw_combined_all['cycle'] == cycle]   # ← Also updated here
        if raw.empty:
            continue
        row = compute_group_metrics(raw)
        row['CYCLE'] = cycle
        row['DATE'] = 'Cycle Total'
        row['CLIENT'] = 'All'
        predictive_cycles_total[cycle_key] = pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)

    # Display sections (unchanged)
    st.write("## Overall Combined Summary (Daily)")
    st.dataframe(combined_summary, use_container_width=True)

    if not overall_combined_total.empty:
        st.write("## Overall Combined Summary (Total Period)")
        st.dataframe(overall_combined_total, use_container_width=True)

    if predictive_cycles_total:
        st.write("## Predictive Cycles (Total)")
        for k in sorted(predictive_cycles_total, key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1))):
            st.subheader(k)
            st.dataframe(predictive_cycles_total[k], use_container_width=True)

    for title, df in [("Overall Predictive", predictive_summary), ("Overall Manual", manual_summary)]:
        if not df.empty:
            st.write(f"## {title} Summary")
            st.dataframe(df, use_container_width=True)

    # Excel Export
    summary_groups = {
        'Combined': {'Combined Summary (Daily)': combined_summary},
        'Overall Combined Summary': {'Overall Combined Summary (Total)': overall_combined_total},
        'Predictive': {'Predictive Summary': predictive_summary},
        'Manual': {'Manual Summary': manual_summary},
        'Combined Cycles': combined_cycle,
        'Predictive Cycles': predictive_cycles,
        'Manual Cycles': manual_cycles,
        'Predictive Cycles (Total)': predictive_cycles_total,
        'Combined Balances': combined_balance,
        'Predictive Balances': predictive_balances,
        'Manual Balances': manual_balances,
    }

    excel_bytes = to_excel(summary_groups)

    st.sidebar.download_button(
        label="Download All Summaries as Excel",
        data=excel_bytes,
        file_name=f"Debt_Collection_Summary_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Raw Cycle Breakdown - Now matches Overall Combined Summary
    if not raw_combined_all.empty:
        raw_cycle_bytes = create_raw_cycle_breakdown_excel(raw_combined_all)

        st.sidebar.download_button(
            label="Download Raw Cycle Breakdown (Overall Combined)",
            data=raw_cycle_bytes,
            file_name=f"Raw_Combined_Cycle_Breakdown_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="raw_cycle_download"
        )

else:
    st.info("Upload one or more XLSX files to start.")