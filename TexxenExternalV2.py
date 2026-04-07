import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter
import re

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

# Removed: COLLECTORS, RPC, BANK ESCALATION, CONNECTED AVE
SUMMARY_COLUMNS = [
    'CYCLE', 'DATE', 'CLIENT', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE',
    'CONNECTED NU', 'CONNECTED UNIQUE', 'TOTAL RPC', 'PTP',
    'CONNECTED % NU', 'CONNECTED % UNIQUE', 'RPC %', 'PTP %',
    'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP', 'CALL DROP RATE'
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

def apply_date_filter(df, start, end):
    if df.empty:
        return df

    date_col = next((col for col in ['barcodedate', 'date'] if col in df.columns), None)
    if not date_col:
        return df

    dt = pd.to_datetime(df[date_col], errors='coerce')
    lo = pd.Timestamp(start)
    hi = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return df.loc[(dt >= lo) & (dt <= hi)].copy()

# ────────────────────────────────────────────────
#  METRIC CALCULATION (Original Logic)
# ────────────────────────────────────────────────

def compute_group_metrics(g):
    if g.empty:
        return pd.Series({c: 0 for c in SUMMARY_COLUMNS if c not in ['CYCLE','DATE','CLIENT']})

    accounts      = g['chcode'].nunique()
    total_dialed  = len(g)

    excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
    has_valid_substatus = (g['substatus'].notna() & (g['substatus'].astype(str).str.strip() != ''))
    not_excluded = ~g['substatus'].astype(str).str.upper().isin(excluded_upper)
    connected_mask = has_valid_substatus & not_excluded

    connected_nu     = len(g[connected_mask])
    connected_unique = g[connected_mask]['chcode'].nunique()

    rpc_ptp_mask  = g['groupstatus'].str.contains('rpc|ptp', case=False, na=False)
    total_rpc     = g[rpc_ptp_mask]['chcode'].nunique()

    ptp_mask      = g['groupstatus'].str.contains('ptp', case=False, na=False)
    ptp           = g[ptp_mask]['chcode'].nunique()

    ptp_accounts_ob = g[ptp_mask].groupby('chcode')['ob'].first()
    total_balance   = ptp_accounts_ob.sum() if not ptp_accounts_ob.empty else 0

    neg_drop_mask = g['substatus'].str.contains('call drop', case=False, na=False)
    neg_drop      = len(g[neg_drop_mask])

    sys_drop_mask = g['substatus'].str.contains('pdrop', case=False, na=False)
    system_drop   = len(g[sys_drop_mask])

    pen_rate         = f"{(total_dialed / accounts * 100):.2f}%" if accounts > 0 else "0.00%"
    connected_pct_nu = f"{(connected_nu / total_dialed * 100):.2f}%" if total_dialed > 0 else "0.00%"
    connected_pct_u  = f"{(connected_unique / accounts * 100):.2f}%" if accounts > 0 else "0.00%"

    rpc_pct          = f"{(total_rpc / connected_unique * 100):.2f}%" if connected_unique > 0 else "0.00%"
    ptp_pct          = f"{(ptp / total_rpc * 100):.2f}%" if total_rpc > 0 else "0.00%"
    drop_rate        = f"{((neg_drop + system_drop) / total_dialed * 100):.2f}%" if total_dialed > 0 else "0.00%"

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

# ────────────────────────────────────────────────
#  BUILD FUNCTIONS
# ────────────────────────────────────────────────

def build_daily_summary(df):
    if df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    df = df.copy()
    df['date']   = pd.to_datetime(df.get('barcodedate', pd.NaT), errors='coerce').dt.date
    df['client'] = df.get('bankname', pd.Series()).astype(str).str.strip()

    agg = df.groupby(['date', 'client'], dropna=False).apply(compute_group_metrics).reset_index()

    agg = agg.rename(columns={'date': 'DATE', 'client': 'CLIENT'})
    agg['CYCLE'] = ''

    agg = agg.reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    agg = agg.sort_values(['DATE', 'CLIENT'])

    return agg


def build_total_row(df):
    if df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    m = compute_group_metrics(df)
    m['CYCLE'] = 'All'
    m['DATE']  = 'Overall'
    m['CLIENT'] = 'All Clients'
    return pd.DataFrame([m])[SUMMARY_COLUMNS]


def get_cycle_summaries(df):
    result = {}
    for cyc in sorted(df['cycle'].unique(), key=lambda x: int(x) if str(x).isdigit() else 9999):
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
    ranges = [(0, 49999.99, "0-49999.99"), (50000, 99999.99, "50000.00-99999.99"), (100000, float('inf'), "100000.00 and up")]
    for min_b, max_b, name in ranges:
        b_df = df[(df['ob'] >= min_b) & (df['ob'] <= max_b)]
        if b_df.empty:
            continue
        for cyc in sorted(b_df['cycle'].unique(), key=lambda x: int(x) if str(x).isdigit() else 9999):
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

        df['DATE'] = pd.to_datetime(df['DATE']).dt.date

        num_cols = ['ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                    'TOTAL RPC', 'PTP', 'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP']
        df[num_cols] = df[num_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

        agg = df.groupby(['DATE', 'CLIENT'], as_index=False)[num_cols].sum()

        agg['PENETRATION RATE'] = agg.apply(
            lambda r: f"{(r['TOTAL DIALED'] / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)
        agg['CONNECTED % NU'] = agg.apply(
            lambda r: f"{(r['CONNECTED NU'] / r['TOTAL DIALED'] * 100):.2f}%" if r['TOTAL DIALED'] > 0 else "0.00%", axis=1)
        agg['CONNECTED % UNIQUE'] = agg.apply(
            lambda r: f"{(r['CONNECTED UNIQUE'] / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)
        agg['RPC %'] = agg.apply(
            lambda r: f"{(r['TOTAL RPC'] / r['CONNECTED UNIQUE'] * 100):.2f}%" if r['CONNECTED UNIQUE'] > 0 else "0.00%", axis=1)
        agg['PTP %'] = agg.apply(
            lambda r: f"{(r['PTP'] / r['TOTAL RPC'] * 100):.2f}%" if r['TOTAL RPC'] > 0 else "0.00%", axis=1)
        agg['CALL DROP RATE'] = agg.apply(
            lambda r: f"{((r['SYSTEM DROP']) / r['ACCOUNTS'] * 100):.2f}%" if r['ACCOUNTS'] > 0 else "0.00%", axis=1)

        agg['CYCLE'] = key.replace("Cycle ", "") if "Cycle " in key else ""

        combined[key] = agg.reindex(columns=SUMMARY_COLUMNS, fill_value=0).sort_values('DATE')

    return combined

# ────────────────────────────────────────────────
#  EXCEL WRITERS
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
    elif sheet_name in ['Overall Combined Cycles', 'Predictive Cycles (Total)']:
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
    order = ['Combined', 'Overall Combined Summary', 'Overall Combined Cycles', 'Predictive', 'Manual',
             'Combined Cycles', 'Predictive Cycles', 'Manual Cycles', 'Predictive Cycles (Total)',
             'Combined Balances', 'Predictive Balances', 'Manual Balances']
    ordered = {k: summary_groups[k] for k in order if k in summary_groups}

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        fmts = setup_excel_formats(writer.book)
        for sheet_name, d in ordered.items():
            write_excel_sheet(writer, sheet_name, d, fmts)
    return output.getvalue()

# ────────────────────────────────────────────────
#  RAW CYCLE BREAKDOWN
# ────────────────────────────────────────────────

def create_raw_cycle_breakdown_excel(raw_combined_df, overall_total_df, overall_cycles_dict):
    if raw_combined_df.empty:
        return BytesIO().getvalue()

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book

        bold = wb.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
        head = wb.add_format({'bold': True, 'bg_color': '#E6F0FA', 'border': 1, 'align': 'center'})
        numf = wb.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
        cent = wb.add_format({'align': 'center', 'border': 1})
        sect = wb.add_format({'bold': True, 'font_size': 12, 'align': 'left'})

        def write_cycle_raw(ws, cycle_key, summary, raw_df, start_row, is_combined=False):
            title = f"{cycle_key} (Combined)" if is_combined else cycle_key
            ws.write(start_row, 0, title, bold)
            current_row = start_row + 2

            headers = ['ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                       'TOTAL RPC', 'PTP', 'NEG DROP', 'SYSTEM DROP']
            for col_idx, h in enumerate(headers):
                ws.write(current_row, col_idx, h, head)
            current_row += 1

            values = [summary.get(h, 0) for h in headers]
            for col_idx, val in enumerate(values):
                ws.write_number(current_row, col_idx, val, numf)
            current_row += 3

            list_headers = ["Raw Accounts", "Raw Connected NU", "Raw Connected Unique",
                            "Raw Total RPC", "Raw PTP", "Raw Negative Drop", "Raw System Drop"]
            for col_idx, header in enumerate(list_headers):
                ws.write(current_row, col_idx, header, head)
            current_row += 1

            if raw_df.empty:
                ws.write(current_row, 0, "(No raw data)", sect)
                return current_row + 5

            excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
            has_valid = (raw_df['substatus'].notna() & (raw_df['substatus'].astype(str).str.strip() != ''))
            not_excluded = ~raw_df['substatus'].astype(str).str.upper().isin(excluded_upper)
            connected_mask = has_valid & not_excluded

            raw_lists = [
                sorted(raw_df['chcode'].dropna().unique()),
                list(raw_df['chcode']),
                sorted(raw_df[connected_mask]['chcode'].dropna().unique()),
                sorted(raw_df[raw_df['groupstatus'].str.contains('rpc|ptp', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_df[raw_df['groupstatus'].str.contains('ptp', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_df[raw_df['substatus'].str.contains('call drop', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_df[raw_df['substatus'].str.contains('pdrop', case=False, na=False)]['chcode'].dropna().unique())
            ]

            max_rows = max(len(lst) for lst in raw_lists) if raw_lists else 0

            for r in range(max_rows):
                for c, lst in enumerate(raw_lists):
                    if r < len(lst):
                        ws.write(current_row + r, c, str(lst[r]), cent)

            for c in range(len(list_headers)):
                ws.set_column(c, c, 28)

            return current_row + max_rows + 5

        # Combined Cycles Raw
        ws_cycles = wb.add_worksheet("Combined Cycles Raw")
        row = 0
        sorted_cycles = sorted([c for c in raw_combined_df['cycle'].unique() if str(c).isdigit()], key=int)

        for cycle_str in sorted_cycles:
            cycle_raw = raw_combined_df[raw_combined_df['cycle'] == cycle_str]
            if cycle_raw.empty:
                continue
            summary = compute_group_metrics(cycle_raw)
            row = write_cycle_raw(ws_cycles, f"Cycle {cycle_str}", summary, cycle_raw, row, is_combined=True)

        # Overall Combined Raw
        ws_total = wb.add_worksheet("Overall Combined Raw")
        if not overall_total_df.empty:
            ws_total.write(0, 0, "Overall Combined Summary (Total Period)", bold)

            headers = ['ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                       'TOTAL RPC', 'PTP', 'NEG DROP', 'SYSTEM DROP']
            for col_idx, h in enumerate(headers):
                ws_total.write(2, col_idx, h, head)
            for col_idx, h in enumerate(headers):
                ws_total.write_number(3, col_idx, float(overall_total_df.iloc[0].get(h, 0)), numf)

            row = 6
            list_headers = ["Raw Accounts", "Raw Connected NU", "Raw Connected Unique",
                            "Raw Total RPC", "Raw PTP", "Raw Negative Drop", "Raw System Drop"]
            for col_idx, header in enumerate(list_headers):
                ws_total.write(row, col_idx, header, head)
            row += 1

            excluded_upper = [s.upper() for s in EXCLUDED_SUBSTATUSES]
            has_valid = (raw_combined_df['substatus'].notna() & (raw_combined_df['substatus'].astype(str).str.strip() != ''))
            not_excluded = ~raw_combined_df['substatus'].astype(str).str.upper().isin(excluded_upper)
            connected_mask = has_valid & not_excluded

            raw_lists = [
                sorted(raw_combined_df['chcode'].dropna().unique()),
                list(raw_combined_df['chcode']),
                sorted(raw_combined_df[connected_mask]['chcode'].dropna().unique()),
                sorted(raw_combined_df[raw_combined_df['groupstatus'].str.contains('rpc|ptp', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_combined_df[raw_combined_df['groupstatus'].str.contains('ptp', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_combined_df[raw_combined_df['substatus'].str.contains('call drop', case=False, na=False)]['chcode'].dropna().unique()),
                sorted(raw_combined_df[raw_combined_df['substatus'].str.contains('pdrop', case=False, na=False)]['chcode'].dropna().unique())
            ]
            max_rows = max((len(lst) for lst in raw_lists), default=0)

            for r in range(max_rows):
                for c, lst in enumerate(raw_lists):
                    if r < len(lst):
                        ws_total.write(row + r, c, str(lst[r]), cent)
            for c in range(len(list_headers)):
                ws_total.set_column(c, c, 28)

        # Overall Combined Cycles Raw
        ws_total_cycles = wb.add_worksheet("Overall Combined Cycles Raw")
        row = 0
        for cycle_key, summary_df in overall_cycles_dict.items():
            if summary_df.empty:
                continue
            cycle_match = re.search(r'Cycle (\d+)', cycle_key)
            if not cycle_match:
                continue
            cycle_str = cycle_match.group(1)
            cycle_raw = raw_combined_df[raw_combined_df['cycle'] == cycle_str]
            row = write_cycle_raw(ws_total_cycles, cycle_key, summary_df.iloc[0], cycle_raw, row, is_combined=True)

    return output.getvalue()

# ────────────────────────────────────────────────
#  MAIN APP
# ────────────────────────────────────────────────

uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### ðŸ“… Date Range Filter")
st.sidebar.caption("Only data within this range will be processed.")

_today = datetime.date.today()
_30ago = _today - datetime.timedelta(days=30)

date_filter_on = st.sidebar.toggle("Enable date range filter", value=False)

if date_filter_on:
    c1, c2 = st.sidebar.columns(2)
    start_date = c1.date_input("Start", value=_30ago, key="texxen_sd")
    end_date = c2.date_input("End", value=_today, key="texxen_ed")
    if start_date > end_date:
        st.sidebar.error("âš ï¸ Start must be â‰¤ End date.")
        st.stop()
    st.sidebar.info(f"ðŸ“† {start_date:%b %d, %Y} â†’ {end_date:%b %d, %Y}")
else:
    start_date = end_date = None
    st.sidebar.caption("_Filter off â€” all dates processed._")

st.sidebar.markdown("---")

if uploaded_files:
    if date_filter_on:
        st.info(f"ðŸ“… **Date filter active:** {start_date:%B %d, %Y} â†’ {end_date:%B %d, %Y}")

    all_combined, all_predictive, all_manual = [], [], []
    predictive_cycles, manual_cycles = {}, {}
    predictive_balances, manual_balances = {}, {}
    raw_combined_all = pd.DataFrame()

    prog = st.progress(0)
    for idx, f in enumerate(uploaded_files):
        with st.spinner(f"Processing {f.name}..."):
            df_clean = load_and_filter_data(f)
            if date_filter_on:
                df_clean = apply_date_filter(df_clean, start_date, end_date)
                if df_clean.empty:
                    prog.progress((idx + 1) / len(uploaded_files))
                    continue

            disp = df_clean.get('dispositiontype', pd.Series()).astype(str).str.upper()

            df_pred = df_clean[disp == 'PREDICTIVE'].copy()
            df_man  = df_clean[disp == 'MANUAL'].copy()
            df_comb = pd.concat([df_pred, df_man], ignore_index=True)

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
    date_tag = (f" â€” filtered {start_date:%b %d} â€“ {end_date:%b %d, %Y}" if date_filter_on else "")
    st.success(f"Processed {len(uploaded_files)} file(s){date_tag}")

    combined_summary = pd.concat(all_combined, ignore_index=True).sort_values('DATE') if all_combined else pd.DataFrame(columns=SUMMARY_COLUMNS)
    predictive_summary = pd.concat(all_predictive, ignore_index=True).sort_values('DATE') if all_predictive else pd.DataFrame(columns=SUMMARY_COLUMNS)
    manual_summary = pd.concat(all_manual, ignore_index=True).sort_values('DATE') if all_manual else pd.DataFrame(columns=SUMMARY_COLUMNS)

    if raw_combined_all.empty:
        st.warning("âš ï¸ No records found in the selected date range. Adjust the filter or disable it.")
        st.stop()

    combined_cycle   = combine_summaries(predictive_cycles, manual_cycles)
    combined_balance = combine_summaries(predictive_balances, manual_balances)

    overall_combined_total = build_total_row(raw_combined_all)

    # Overall Combined Cycles
    overall_combined_cycles = {}
    for cycle_str in sorted(raw_combined_all['cycle'].unique(), key=lambda x: int(x) if str(x).isdigit() else 9999):
        if not cycle_str or cycle_str == "Unknown":
            continue
        sub = raw_combined_all[raw_combined_all['cycle'] == cycle_str]
        if sub.empty:
            continue
        m = compute_group_metrics(sub)
        m['CYCLE'] = cycle_str
        m['DATE'] = 'Cycle Total'
        m['CLIENT'] = 'All'
        overall_combined_cycles[f"Cycle {cycle_str}"] = pd.DataFrame([m])[SUMMARY_COLUMNS]

    # Display
    st.write("## Overall Combined Summary (Daily)")
    st.dataframe(combined_summary, use_container_width=True)

    if not overall_combined_total.empty:
        st.write("## Overall Combined Summary (Total Period)")
        st.dataframe(overall_combined_total, use_container_width=True)

    if overall_combined_cycles:
        st.write("## Overall Combined Cycles (Total per Cycle)")
        for k in sorted(overall_combined_cycles.keys(), key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1))):
            st.subheader(k)
            st.dataframe(overall_combined_cycles[k], use_container_width=True)

    for title, df in [("Overall Predictive", predictive_summary), ("Overall Manual", manual_summary)]:
        if not df.empty:
            st.write(f"## {title} Summary")
            st.dataframe(df, use_container_width=True)

    # Export
    summary_groups = {
        'Combined': {'Combined Summary (Daily)': combined_summary},
        'Overall Combined Summary': {'Overall Combined Summary (Total)': overall_combined_total},
        'Overall Combined Cycles': overall_combined_cycles,
        'Predictive': {'Predictive Summary': predictive_summary},
        'Manual': {'Manual Summary': manual_summary},
        'Combined Cycles': combined_cycle,
        'Predictive Cycles': predictive_cycles,
        'Manual Cycles': manual_cycles,
        'Combined Balances': combined_balance,
        'Predictive Balances': predictive_balances,
        'Manual Balances': manual_balances,
    }

    date_suffix = (f"_{start_date:%Y%m%d}_to_{end_date:%Y%m%d}"
                   if date_filter_on else f"_{datetime.datetime.now():%Y%m%d_%H%M%S}")

    excel_bytes = to_excel(summary_groups)

    st.sidebar.download_button(
        label="Download All Summaries as Excel",
        data=excel_bytes,
        file_name=f"Debt_Collection_Summary{date_suffix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    raw_cycle_bytes = create_raw_cycle_breakdown_excel(
        raw_combined_all,
        overall_combined_total,
        overall_combined_cycles,
    )

    st.sidebar.download_button(
        label="Download Raw Cycle Breakdown",
        data=raw_cycle_bytes,
        file_name=f"Raw_Combined_Cycle_Breakdown{date_suffix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload one or more XLSX files to start.")
