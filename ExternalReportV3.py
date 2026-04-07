import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter
import numpy as np
import re

# ----------------------------------------------------------------------
# GLOBAL CONSTANTS
# ----------------------------------------------------------------------
BALANCE_ORDER = ["0-49999.99", "50000.00-99999.99", "100000.00 and up"]
SUMMARY_COLUMNS = [
    'CYCLE', 'DATE', 'CLIENT', 'COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE',
    'CONNECTED NU', 'CONNECTED UNIQUE', 'TOTAL RPC', 'RPC', 'PTP', 'BANK ESCALATION',
    'CONNECTED % NU', 'CONNECTED % UNIQUE', 'RPC %', 'PTP %', 'TOTAL TALK TIME',
    'TALK TIME AVE', 'CONNECTED AVE', 'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP', 'CALL DROP RATE'
]
PERCENTAGE_COLS = ['PENETRATION RATE', 'CONNECTED % NU', 'CONNECTED % UNIQUE',
                   'RPC %', 'PTP %', 'CALL DROP RATE']
NUMERICAL_COLS = ['COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                  'TOTAL RPC', 'RPC', 'PTP', 'BANK ESCALATION', 'TOTAL BALANCE',
                  'NEG DROP', 'SYSTEM DROP', 'CONNECTED AVE']
CONNECTED_STATUS_TITLE = 'Connected Status'
CONNECTED_STATUS_CYCLE_TITLE = 'Connected Status Breakdown'

st.set_page_config(layout="wide", page_title="Debt Collection Summary (Volare)",
                   page_icon="📊", initial_sidebar_state="expanded")
st.title('Debt Collection Summary')

# ----------------------------------------------------------------------
# DATA LOADING
# ----------------------------------------------------------------------
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    return df

# ----------------------------------------------------------------------
# EXCEL FORMATS
# ----------------------------------------------------------------------
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
        'time': workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                     'border': 1, 'num_format': 'hh:mm:ss'})
    }

def parse_percentage_value(value):
    if pd.isna(value):
        return 0.0

    if isinstance(value, str):
        cleaned = value.strip().replace(',', '')
        if not cleaned:
            return 0.0

        has_percent = '%' in cleaned
        cleaned = cleaned.rstrip('%').strip()
        numeric = pd.to_numeric(cleaned, errors='coerce')
        if pd.isna(numeric):
            return np.nan
        return float(numeric) / 100 if has_percent else float(numeric)

    numeric = pd.to_numeric(value, errors='coerce')
    return float(numeric) if not pd.isna(numeric) else np.nan

def normalize_percentage_columns(df):
    for col in PERCENTAGE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(parse_percentage_value)
    return df

# ----------------------------------------------------------------------
# EXCEL WRITER (unchanged)
# ----------------------------------------------------------------------
def write_excel_sheet(writer, sheet_name, df_dict, formats):
    if not df_dict:
        return

    worksheet = writer.book.add_worksheet(sheet_name)
    current_row = 0

    sorted_items = []
    if sheet_name == 'Overall Combined Summary':
        main_summary_key = next((k for k in df_dict.keys() if 'Total' in k), None)
        connected_status_key = next((k for k in df_dict.keys() if k == CONNECTED_STATUS_TITLE), None)
        
        if main_summary_key: sorted_items.append((main_summary_key, df_dict[main_summary_key]))
        if connected_status_key: sorted_items.append((connected_status_key, df_dict[connected_status_key]))
        
        for k, v in df_dict.items():
             if k not in [main_summary_key, connected_status_key]: sorted_items.append((k, v))
             
    elif sheet_name == 'Predictive Cycles (Total)':
        cycle_keys = [k for k in df_dict.keys() if k.startswith('Cycle ') and not k.endswith('Breakdown')]
        cycle_keys.sort(key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1)))
        
        for cycle_key in cycle_keys:
            sorted_items.append((cycle_key, df_dict[cycle_key]))
            breakdown_key = f"{cycle_key} {CONNECTED_STATUS_CYCLE_TITLE}"
            if breakdown_key in df_dict:
                sorted_items.append((breakdown_key, df_dict[breakdown_key]))
        
    elif sheet_name.endswith('Cycles'):
        sorted_items = sorted(
            df_dict.items(),
            key=lambda x: int(re.search(r'Cycle (\d+)', x[0]).group(1))
            if re.search(r'Cycle (\d+)', x[0]) else float('inf')
        )
    elif sheet_name.endswith('Balances'):
        sorted_items = sorted(
            df_dict.items(),
            key=lambda x: (
                int(re.search(r'Cycle (\d+)', x[0]).group(1))
                if re.search(r'Cycle (\d+)', x[0]) else float('inf'),
                BALANCE_ORDER.index(re.search(r'Balance (.+)$', x[0]).group(1))
                if re.search(r'Balance (.+)$', x[0]) else float('inf')
            )
        )
    else:
        sorted_items = list(df_dict.items())

    for title, df in sorted_items:
        if df.empty:
            continue

        df_excel = df.copy()
        
        if title == CONNECTED_STATUS_TITLE or title.endswith(CONNECTED_STATUS_CYCLE_TITLE):
             df_display = df_excel.T.reset_index()
             df_display.columns = ['Status/Metric', 'Count']
        else:
            df_display = df_excel.copy()
            if 'DATE' in df_display.columns:
                df_display = df_display.sort_values(by=['DATE'], na_position='first')

            for col in NUMERICAL_COLS:
                if col in df_display.columns:
                    df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)

            df_display = normalize_percentage_columns(df_display)
        
        worksheet.merge_range(current_row, 0, current_row, len(df_display.columns) - 1, title, formats['title'])
        current_row += 1

        for col_num, col_name in enumerate(df_display.columns):
            worksheet.write(current_row, col_num, col_name, formats['header'])
            max_len = max(df_display[col_name].astype(str).str.len().max(), len(col_name)) + 2
            worksheet.set_column(col_num, col_num, max_len)
        current_row += 1

        for row_num in range(len(df_display)):
            for col_num, col_name in enumerate(df_display.columns):
                value = df_display.iloc[row_num, col_num]

                if col_name == 'DATE':
                    if pd.isna(value):
                        worksheet.write_blank(current_row + row_num, col_num, None, formats['center'])
                    elif isinstance(value, str) and ' to ' in value:
                        worksheet.write_string(current_row + row_num, col_num, value, formats['center'])
                    else:
                        try:
                            dt_val = pd.to_datetime(value).to_pydatetime()
                            worksheet.write_datetime(current_row + row_num, col_num, dt_val, formats['date'])
                        except:
                            worksheet.write_string(current_row + row_num, col_num, str(value), formats['center'])
                elif col_name == 'TOTAL BALANCE':
                    worksheet.write_number(current_row + row_num, col_num, float(value), formats['comma'])
                elif col_name in PERCENTAGE_COLS:
                    percent_value = parse_percentage_value(value)
                    if pd.isna(percent_value):
                        worksheet.write(current_row + row_num, col_num, value, formats['center'])
                    else:
                        worksheet.write_number(current_row + row_num, col_num, percent_value, formats['percent'])
                elif col_name in ['TOTAL TALK TIME', 'TALK TIME AVE']:
                    worksheet.write_string(current_row + row_num, col_num, str(value), formats['time'])
                elif col_name == 'Count' or col_name in NUMERICAL_COLS:
                    try:
                        worksheet.write_number(current_row + row_num, col_num, float(value), formats['comma'])
                    except:
                        worksheet.write(current_row + row_num, col_num, value, formats['center'])
                else:
                    worksheet.write(current_row + row_num, col_num, value, formats['center'])
        current_row += len(df_display) + 2

def to_excel(summary_groups):
    output = BytesIO()
    order = ['Combined', 'Overall Combined Summary', 'Predictive', 'Manual',
             'Combined Cycles', 'Predictive Cycles', 'Manual Cycles', 'Predictive Cycles (Total)',
             'Combined Balances', 'Predictive Balances', 'Manual Balances']
    ordered_groups = {k: summary_groups[k] for k in order if k in summary_groups}

    with ExcelWriter(output, engine='xlsxwriter') as writer:
        formats = setup_excel_formats(writer.book)
        for sheet_name, df_dict in ordered_groups.items():
            write_excel_sheet(writer, sheet_name, df_dict, formats)
    return output.getvalue()

# ----------------------------------------------------------------------
# Raw Cycle Breakdown (two sheets – mirrored structure)
# ----------------------------------------------------------------------
def create_raw_cycle_breakdown_excel(raw_predictive_df, raw_combined_df, predictive_cycles_dict, overall_combined_total_df):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        bold_title    = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
        header_fmt    = workbook.add_format({'bold': True, 'bg_color': '#E6F0FA', 'border': 1, 'align': 'center'})
        number_fmt    = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
        center        = workbook.add_format({'align': 'center', 'border': 1})
        section_title = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left'})

        def write_card_list(ws, title, card_series, start_row):
            if card_series.empty or len(card_series.dropna()) == 0:
                ws.write(start_row, 0, f"{title}: (none)", section_title)
                return start_row + 2

            ws.write(start_row, 0, title, section_title)
            start_row += 1

            cards = card_series.astype(str).replace('', np.nan).dropna().unique()
            cards_sorted = sorted(cards)

            for i, card in enumerate(cards_sorted):
                ws.write(start_row + i, 0, card, center)

            return start_row + len(cards_sorted) + 2

        # ── Predictive Cycles Raw ───────────────────────────────
        ws_pred = workbook.add_worksheet("Predictive Cycles Raw")
        current_row = 0

        cycle_totals = aggregate_cycles_by_cycle(predictive_cycles_dict, raw_predictive_df)

        sorted_cycles = sorted(
            [k for k in cycle_totals if not k.endswith(CONNECTED_STATUS_CYCLE_TITLE)],
            key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1)) if re.search(r'Cycle (\d+)', x) else 9999
        )

        for cycle_key in sorted_cycles:
            summary_df = cycle_totals.get(cycle_key)
            if summary_df.empty:
                continue

            summary = summary_df.iloc[0]

            ws_pred.write(current_row, 0, cycle_key, bold_title)
            current_row += 2

            headers = [
                'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                'TOTAL RPC', 'RPC', 'PTP', 'BANK ESCALATION', 'NEG DROP', 'SYSTEM DROP'
            ]
            for col_idx, header in enumerate(headers):
                ws_pred.write(current_row, col_idx, header, header_fmt)
            current_row += 1

            values = [summary.get(h, 0) for h in headers]
            for col_idx, val in enumerate(values):
                ws_pred.write_number(current_row, col_idx, val, number_fmt)
            current_row += 2

            cycle_num = re.search(r'Cycle (\d+)', cycle_key).group(1)
            cycle_raw = raw_predictive_df[raw_predictive_df['CYCLE'] == cycle_num].copy()

            if cycle_raw.empty:
                ws_pred.write(current_row, 0, "(No raw data)", section_title)
                current_row += 3
                continue

            connected_mask = cycle_raw['TALK TIME DURATION'] > 0
            ptp_mask = cycle_raw['PTP AMOUNT'] > 0
            rpc_mask = cycle_raw['STATUS'].str.contains('rpc', case=False, na=False)
            esc_mask = cycle_raw['STATUS'].str.contains('bank escalation', case=False, na=False)
            neg_mask = cycle_raw['STATUS'].str.contains('NEGATIVE CALLOUTS - CALL DROP', case=False, na=False)
            drop_mask = cycle_raw['STATUS'].str.contains('DROPPED', case=False, na=False)

            current_row = write_card_list(ws_pred, "Raw Accounts (unique CARD NO.)", cycle_raw['CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw Connected", cycle_raw.loc[connected_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw Total RPC", cycle_raw.loc[ptp_mask | rpc_mask | esc_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw RPC", cycle_raw.loc[rpc_mask & connected_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw PTP", cycle_raw.loc[ptp_mask & connected_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw Bank Escalation", cycle_raw.loc[esc_mask & connected_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw Negative Drop", cycle_raw.loc[neg_mask, 'CARD NO.'], current_row)
            current_row = write_card_list(ws_pred, "Raw System Drop", cycle_raw.loc[drop_mask, 'CARD NO.'], current_row)

            current_row += 4

        # ── Overall Combined Raw ────────────────────────────────
        ws_overall = workbook.add_worksheet("Overall Combined Raw")
        row = 0

        if not overall_combined_total_df.empty:
            overall_summary = overall_combined_total_df.iloc[0]

            ws_overall.write(row, 0, "Overall Combined Summary (Total)", bold_title)
            row += 2

            headers = [
                'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                'TOTAL RPC', 'RPC', 'PTP', 'BANK ESCALATION', 'NEG DROP', 'SYSTEM DROP'
            ]
            for col_idx, header in enumerate(headers):
                ws_overall.write(row, col_idx, header, header_fmt)
            row += 1

            values = [overall_summary.get(h, 0) for h in headers]
            for col_idx, val in enumerate(values):
                ws_overall.write_number(row, col_idx, val, number_fmt)
            row += 2

            if not raw_combined_df.empty:
                connected_mask = raw_combined_df['TALK TIME DURATION'] > 0
                ptp_mask = raw_combined_df['PTP AMOUNT'] > 0
                rpc_mask = raw_combined_df['STATUS'].str.contains('rpc', case=False, na=False)
                esc_mask = raw_combined_df['STATUS'].str.contains('bank escalation', case=False, na=False)
                neg_mask = raw_combined_df['STATUS'].str.contains('NEGATIVE CALLOUTS - CALL DROP', case=False, na=False)
                drop_mask = raw_combined_df['STATUS'].str.contains('DROPPED', case=False, na=False)

                row = write_card_list(ws_overall, "Raw Accounts (unique CARD NO.)", raw_combined_df['CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw Connected", raw_combined_df.loc[connected_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw Total RPC", raw_combined_df.loc[ptp_mask | rpc_mask | esc_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw RPC", raw_combined_df.loc[rpc_mask & connected_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw PTP", raw_combined_df.loc[ptp_mask & connected_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw Bank Escalation", raw_combined_df.loc[esc_mask & connected_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw Negative Drop", raw_combined_df.loc[neg_mask, 'CARD NO.'], row)
                row = write_card_list(ws_overall, "Raw System Drop", raw_combined_df.loc[drop_mask, 'CARD NO.'], row)

    return output.getvalue()

# ----------------------------------------------------------------------
# FILE PROCESSING
# ----------------------------------------------------------------------
def process_file(df):
    string_cols = ['DEBTOR', 'STATUS', 'REMARK', 'CALL STATUS', 'CARD NO.']
    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str)

    filter_conditions = []
    if 'REMARK BY' in df.columns:
        filter_conditions.append(df['REMARK BY'] != 'SPMADRID')
    if 'DEBTOR' in df.columns:
        filter_conditions.append(~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False))
    if 'STATUS' in df.columns:
        filter_conditions.append(~df['STATUS'].str.contains('ABORT', na=False))
    if 'REMARK' in df.columns:
        filter_conditions.append(~df['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True))
        excluded = ["Broken Promise", "New files imported", "Updates when case reassign to another collector",
                    "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD",
                    "New Assignment -", "broadcast", "File Unhold"]
        filter_conditions.append(~df['REMARK'].str.contains('|'.join(excluded), case=False, na=False))

    if filter_conditions:
        df = df[pd.concat(filter_conditions, axis=1).all(axis=1)]

    if 'CARD NO.' in df.columns:
        df['CYCLE'] = df['CARD NO.'].str[:2].fillna('Unknown')

    numeric_cols = {'CALL DURATION': 'coerce', 'TALK TIME DURATION': 'coerce',
                    'PTP AMOUNT': 'coerce', 'BALANCE': 'coerce'}
    for col, err in numeric_cols.items():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors=err).fillna(0)

    return df

# ----------------------------------------------------------------------
# TIME HELPERS
# ----------------------------------------------------------------------
def format_seconds_to_hms(seconds):
    if pd.isna(seconds) or seconds <= 0:
        return "00:00:00"
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def hms_to_seconds(hms):
    if pd.isna(hms) or str(hms).strip() == "00:00:00":
        return 0
    try:
        if isinstance(hms, (int, float)):
            return int(hms)
        h, m, s = map(int, str(hms).split(':'))
        return h * 3600 + m * 60 + s
    except:
        return 0

# ----------------------------------------------------------------------
# METRIC CALCULATION
# ----------------------------------------------------------------------
def calculate_metrics(group):
    if 'CALL DURATION' not in group.columns or 'REMARK BY' not in group.columns:
        return None

    collectors = group[group['CALL DURATION'] > 0]['REMARK BY'].nunique()
    if collectors == 0:
        return None

    if 'DEBTOR ID' not in group.columns:
        return None

    accounts = group['DEBTOR ID'].nunique()
    total_dialed = len(group)
    connected_nu = (group['TALK TIME DURATION'] > 0).sum() if 'TALK TIME DURATION' in group.columns else 0
    connected_unique = group[group['TALK TIME DURATION'] > 0]['DEBTOR ID'].nunique() if 'TALK TIME DURATION' in group.columns else 0

    total_rpc = rpc = ptp = bank_escalation = 0
    if all(col in group.columns for col in ['PTP AMOUNT', 'STATUS', 'DEBTOR ID']):
        ptp_mask = group['PTP AMOUNT'] > 0
        rpc_escalation_mask = group['STATUS'].str.contains('bank escalation|rpc', case=False, na=False)
        ptp_ids = group.loc[ptp_mask, 'DEBTOR ID']
        rpc_escalation_ids = group.loc[rpc_escalation_mask, 'DEBTOR ID']
        total_rpc = len(pd.concat([ptp_ids, rpc_escalation_ids]).unique())

        rpc = group.loc[(group['STATUS'].str.contains('rpc', case=False, na=False)) &
                         (group['TALK TIME DURATION'] > 0), 'DEBTOR ID'].nunique()
        ptp = group.loc[ptp_mask & (group['TALK TIME DURATION'] > 0), 'DEBTOR ID'].nunique()
        bank_escalation = group.loc[(group['STATUS'].str.contains('bank escalation', case=False, na=False)) &
                                     (group['TALK TIME DURATION'] > 0), 'DEBTOR ID'].nunique()

    total_talk_time_secs = group['TALK TIME DURATION'].sum() if 'TALK TIME DURATION' in group.columns else 0
    total_talk_time = format_seconds_to_hms(total_talk_time_secs)
    connected_nu_float = float(connected_nu)
    talk_time_ave = format_seconds_to_hms(total_talk_time_secs / connected_nu_float) if connected_nu_float > 0 else "00:00:00"
    connected_ave = round(connected_nu_float / collectors, 2) if collectors > 0 else 0

    total_balance = group.loc[group['PTP AMOUNT'] > 0, 'BALANCE'].sum() if 'BALANCE' in group.columns else 0
    neg_drop = group['STATUS'].str.contains('NEGATIVE CALLOUTS - CALL DROP', case=False, na=False).sum() if 'STATUS' in group.columns else 0
    system_drop = group['STATUS'].str.contains('DROPPED', case=False, na=False).sum() if 'STATUS' in group.columns else 0

    def pct(n, d): return n / d if d > 0 else 0.0
    penetration_rate = f"{pct(total_dialed, accounts)*100:.2f}%"
    connected_nu_rate = f"{pct(connected_nu_float, total_dialed)*100:.2f}%"
    connected_unique_rate = f"{pct(connected_unique, accounts)*100:.2f}%"
    rpc_rate = f"{pct(total_rpc, connected_unique)*100:.2f}%" if connected_unique > 0 else "0.00%"
    ptp_rate = f"{pct(ptp, total_rpc)*100:.2f}%" if total_rpc > 0 else "0.00%"
    call_drop_rate = f"{pct(system_drop, connected_nu_float)*100:.2f}%" if connected_nu_float > 0 else "0.00%"

    return {
        'COLLECTORS': collectors, 'ACCOUNTS': accounts, 'TOTAL DIALED': total_dialed,
        'PENETRATION RATE': penetration_rate, 'CONNECTED NU': connected_nu,
        'CONNECTED UNIQUE': connected_unique, 'TOTAL RPC': total_rpc, 'RPC': rpc,
        'PTP': ptp, 'BANK ESCALATION': bank_escalation,
        'CONNECTED % NU': connected_nu_rate, 'CONNECTED % UNIQUE': connected_unique_rate,
        'RPC %': rpc_rate, 'PTP %': ptp_rate,
        'TOTAL TALK TIME': total_talk_time, 'TALK TIME AVE': talk_time_ave,
        'CONNECTED AVE': connected_ave, 'TOTAL BALANCE': total_balance,
        'NEG DROP': neg_drop, 'SYSTEM DROP': system_drop, 'CALL DROP RATE': call_drop_rate
    }

def calculate_summary(df, remark_types):
    if 'REMARK TYPE' not in df.columns:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    df_f = df[df['REMARK TYPE'].isin(remark_types)].copy()
    if df_f.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    df_f['DATE'] = df_f['DATE'].dt.date
    rows = []
    for (date, client), g in df_f.groupby(['DATE', 'CLIENT']):
        m = calculate_metrics(g)
        if m:
            m.update({'DATE': date, 'CLIENT': client})
            rows.append(m)
    if rows:
        return pd.DataFrame(rows).reindex(columns=SUMMARY_COLUMNS, fill_value=0).sort_values('DATE')
    return pd.DataFrame(columns=SUMMARY_COLUMNS)

def get_cycle_summary(df, remark_types):
    result = {}
    if 'CYCLE' not in df.columns:
        return result
    for cycle in [c for c in df['CYCLE'].unique() if c and str(c).lower() not in ['unknown', 'na']]:
        c_df = df[df['CYCLE'] == cycle]
        if c_df.empty:
            continue
        s = calculate_summary(c_df, remark_types)
        if not s.empty:
            s['CYCLE'] = cycle
            result[f"Cycle {cycle}"] = s
    return result

def get_balance_summary(df, remark_types):
    ranges = [(0, 49999.99, "0-49999.99"), (50000.00, 99999.99, "50000.00-99999.99"), (100000.00, float('inf'), "100000.00 and up")]
    result = {}
    if 'CYCLE' not in df.columns or 'BALANCE' not in df.columns:
        return result
    for cycle in [c for c in df['CYCLE'].unique() if c and str(c).lower() not in ['unknown', 'na']]:
        c_df = df[df['CYCLE'] == cycle]
        for min_b, max_b, name in ranges:
            b_df = c_df[(c_df['BALANCE'] >= min_b) & (c_df['BALANCE'] <= max_b)]
            if b_df.empty:
                continue
            s = calculate_summary(b_df, remark_types)
            if not s.empty:
                s['CYCLE'] = cycle
                result[f"Cycle {cycle} Balance {name}"] = s
    return result

def combine_summaries(pred, man):
    combined = {}
    for key in set(pred) | set(man):
        if "na" in key.lower():
            continue
        p = pred.get(key, pd.DataFrame(columns=SUMMARY_COLUMNS))
        m = man.get(key, pd.DataFrame(columns=SUMMARY_COLUMNS))
        df = pd.concat([p, m], ignore_index=True)
        if df.empty:
            continue

        df['DATE'] = pd.to_datetime(df['DATE']).dt.date
        df['TOTAL_TALK_TIME_SECS'] = df['TOTAL TALK TIME'].apply(hms_to_seconds)

        df = normalize_percentage_columns(df)

        agg = {c: 'sum' for c in NUMERICAL_COLS if c in df.columns}
        agg['TOTAL_TALK_TIME_SECS'] = 'sum'
        g = df.groupby(['DATE', 'CLIENT'], as_index=False).agg(agg)

        g['TOTAL TALK TIME'] = g['TOTAL_TALK_TIME_SECS'].apply(format_seconds_to_hms)
        g['CONNECTED NU'] = g['CONNECTED NU'].apply(float)
        g['TALK TIME AVE'] = g.apply(lambda r: format_seconds_to_hms(r['TOTAL_TALK_TIME_SECS'] / r['CONNECTED NU']) if r['CONNECTED NU'] > 0 else "00:00:00", axis=1)
        g['CONNECTED AVE'] = g.apply(lambda r: round(r['CONNECTED NU'] / r['COLLECTORS'], 2) if r['COLLECTORS'] > 0 else 0, axis=1)

        def pct(n, d): return n / d if d > 0 else 0.0
        g['PENETRATION RATE'] = g.apply(lambda r: f"{pct(r['TOTAL DIALED'], r['ACCOUNTS'])*100:.2f}%", axis=1)
        g['CONNECTED % NU'] = g.apply(lambda r: f"{pct(r['CONNECTED NU'], r['TOTAL DIALED'])*100:.2f}%", axis=1)
        g['CONNECTED % UNIQUE'] = g.apply(lambda r: f"{pct(r['CONNECTED UNIQUE'], r['ACCOUNTS'])*100:.2f}%", axis=1)
        g['RPC %'] = g.apply(lambda r: f"{pct(r['TOTAL RPC'], r['CONNECTED UNIQUE'])*100:.2f}%" if r['CONNECTED UNIQUE'] > 0 else "0.00%", axis=1)
        g['PTP %'] = g.apply(lambda r: f"{pct(r['PTP'], r['TOTAL RPC'])*100:.2f}%" if r['TOTAL RPC'] > 0 else "0.00%", axis=1)
        g['CALL DROP RATE'] = g.apply(lambda r: f"{pct(r['SYSTEM DROP'], r['CONNECTED NU'])*100:.2f}%" if r['CONNECTED NU'] > 0 else "0.00%", axis=1)

        g = g.drop(columns=['TOTAL_TALK_TIME_SECS'], errors='ignore')
        combined[key] = g.reindex(columns=SUMMARY_COLUMNS, fill_value=0).sort_values('DATE')
    return combined

def calculate_connected_status_by_cycle(raw_df):
    if raw_df.empty or 'DEBTOR ID' not in raw_df.columns or 'TALK TIME DURATION' not in raw_df.columns or 'STATUS' not in raw_df.columns:
        return pd.DataFrame()

    connected_calls = raw_df[raw_df['TALK TIME DURATION'] > 0].copy()
    last_connected_status = connected_calls.drop_duplicates(subset=['DEBTOR ID'], keep='last')
    total_connected_unique = last_connected_status['DEBTOR ID'].nunique()
    status_counts = last_connected_status['STATUS'].value_counts().to_dict()

    output_data = {'CONNECTED UNIQUE': total_connected_unique}
    output_data.update(status_counts)
    return pd.DataFrame([output_data])

def aggregate_cycles_by_cycle(cycle_dict, raw_predictive_df):
    if not cycle_dict or raw_predictive_df.empty:
        return {}
    out = {}
    
    unique_cycles = set()
    for key in cycle_dict.keys():
        match = re.search(r'Cycle (\d+)', key)
        if match:
            unique_cycles.add(match.group(1))

    for cycle in sorted(list(unique_cycles)):
        cycle_key = f"Cycle {cycle}"
        raw = raw_predictive_df[raw_predictive_df['CYCLE'] == cycle].copy()
        daily_summary_df = cycle_dict.get(cycle_key, pd.DataFrame(columns=SUMMARY_COLUMNS))
        
        if raw.empty or daily_summary_df.empty:
            continue

        accounts = raw['DEBTOR ID'].nunique()
        connected_mask = raw['TALK TIME DURATION'] > 0
        connected_unique = raw[connected_mask]['DEBTOR ID'].nunique()
        ptp_mask = raw['PTP AMOUNT'] > 0
        rpc_escalation_mask = raw['STATUS'].str.contains('bank escalation|rpc', case=False, na=False)
        total_rpc = len(pd.concat([raw[ptp_mask]['DEBTOR ID'], raw[rpc_escalation_mask]['DEBTOR ID']]).unique())
        rpc = raw[(raw['STATUS'].str.contains('rpc', case=False, na=False)) & connected_mask]['DEBTOR ID'].nunique()
        ptp = raw[ptp_mask & connected_mask]['DEBTOR ID'].nunique()
        bank_escalation = raw[(raw['STATUS'].str.contains('bank escalation', case=False, na=False)) & connected_mask]['DEBTOR ID'].nunique()

        daily_summary_df['TOTAL_TALK_TIME_SECS'] = daily_summary_df['TOTAL TALK TIME'].apply(hms_to_seconds)
        total_dialed = daily_summary_df['TOTAL DIALED'].sum()
        connected_nu = daily_summary_df['CONNECTED NU'].sum()
        total_talk_time_secs = daily_summary_df['TOTAL_TALK_TIME_SECS'].sum()
        neg_drop = daily_summary_df['NEG DROP'].sum()
        system_drop = daily_summary_df['SYSTEM DROP'].sum()
        total_balance = daily_summary_df['TOTAL BALANCE'].sum()
        collectors = daily_summary_df['COLLECTORS'].nunique()

        row = {
            'CYCLE': cycle, 'COLLECTORS': collectors, 'ACCOUNTS': accounts,
            'TOTAL DIALED': total_dialed, 'CONNECTED NU': connected_nu,
            'CONNECTED UNIQUE': connected_unique, 'TOTAL RPC': total_rpc,
            'RPC': rpc, 'PTP': ptp, 'BANK ESCALATION': bank_escalation,
            'TOTAL TALK TIME': format_seconds_to_hms(total_talk_time_secs),
            'TOTAL BALANCE': total_balance, 'NEG DROP': neg_drop, 'SYSTEM DROP': system_drop,
        }
        connected_nu_float = float(connected_nu)
        row['TALK TIME AVE'] = format_seconds_to_hms(total_talk_time_secs / connected_nu_float) if connected_nu_float > 0 else "00:00:00"
        row['CONNECTED AVE'] = round(connected_nu_float / collectors, 2) if collectors > 0 else 0

        def pct(n, d): return n / d if d > 0 else 0.0
        row['PENETRATION RATE'] = f"{pct(total_dialed, accounts)*100:.2f}%"
        row['CONNECTED % NU'] = f"{pct(connected_nu_float, total_dialed)*100:.2f}%"
        row['CONNECTED % UNIQUE'] = f"{pct(connected_unique, accounts)*100:.2f}%"
        row['RPC %'] = f"{pct(total_rpc, connected_unique)*100:.2f}%" if connected_unique > 0 else "0.00%"
        row['PTP %'] = f"{pct(ptp, total_rpc)*100:.2f}%" if total_rpc > 0 else "0.00%"
        row['CALL DROP RATE'] = f"{pct(system_drop, connected_nu_float)*100:.2f}%" if connected_nu_float > 0 else "0.00%"

        out[cycle_key] = pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)
        
        status_breakdown_df = calculate_connected_status_by_cycle(raw)
        if not status_breakdown_df.empty:
            out[f"{cycle_key} {CONNECTED_STATUS_CYCLE_TITLE}"] = status_breakdown_df
            
    return out

def aggregate_overall_summary(raw_df, daily_summary_df):
    if raw_df.empty or daily_summary_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    raw = raw_df.copy()
    accounts = raw['DEBTOR ID'].nunique()
    connected_mask = raw['TALK TIME DURATION'] > 0
    connected_unique = raw[connected_mask]['DEBTOR ID'].nunique()
    ptp_mask = raw['PTP AMOUNT'] > 0
    rpc_escalation_mask = raw['STATUS'].str.contains('bank escalation|rpc', case=False, na=False)
    total_rpc = len(pd.concat([raw[ptp_mask]['DEBTOR ID'], raw[rpc_escalation_mask]['DEBTOR ID']]).unique())
    rpc = raw[(raw['STATUS'].str.contains('rpc', case=False, na=False)) & connected_mask]['DEBTOR ID'].nunique()
    ptp = raw[ptp_mask & connected_mask]['DEBTOR ID'].nunique()
    bank_escalation = raw[(raw['STATUS'].str.contains('bank escalation', case=False, na=False)) & connected_mask]['DEBTOR ID'].nunique()

    daily = daily_summary_df.copy()
    daily['TOTAL_TALK_TIME_SECS'] = daily['TOTAL TALK TIME'].apply(hms_to_seconds)
    total_dialed = daily['TOTAL DIALED'].sum()
    connected_nu = daily['CONNECTED NU'].sum()
    total_talk_time_secs = daily['TOTAL_TALK_TIME_SECS'].sum()
    neg_drop = daily['NEG DROP'].sum()
    system_drop = daily['SYSTEM DROP'].sum()
    total_balance = daily['TOTAL BALANCE'].sum()
    collectors = daily['COLLECTORS'].nunique()

    date_min = pd.to_datetime(daily['DATE']).min()
    date_max = pd.to_datetime(daily['DATE']).max()
    date_range = f"{date_min.strftime('%Y-%m-%d')} to {date_max.strftime('%Y-%m-%d')}" if not pd.isna(date_min) else "N/A"

    row = {
        'CYCLE': 'All', 'DATE': date_range, 'CLIENT': 'All',
        'COLLECTORS': collectors, 'ACCOUNTS': accounts,
        'TOTAL DIALED': total_dialed, 'CONNECTED NU': connected_nu,
        'CONNECTED UNIQUE': connected_unique, 'TOTAL RPC': total_rpc,
        'RPC': rpc, 'PTP': ptp, 'BANK ESCALATION': bank_escalation,
        'TOTAL TALK TIME': format_seconds_to_hms(total_talk_time_secs),
        'TOTAL BALANCE': total_balance, 'NEG DROP': neg_drop, 'SYSTEM DROP': system_drop,
    }
    connected_nu_float = float(connected_nu)
    row['TALK TIME AVE'] = format_seconds_to_hms(total_talk_time_secs / connected_nu_float) if connected_nu_float > 0 else "00:00:00"
    row['CONNECTED AVE'] = round(connected_nu_float / collectors, 2) if collectors > 0 else 0

    def pct(n, d): return n / d if d > 0 else 0.0
    row['PENETRATION RATE'] = f"{pct(total_dialed, accounts)*100:.2f}%"
    row['CONNECTED % NU'] = f"{pct(connected_nu_float, total_dialed)*100:.2f}%"
    row['CONNECTED % UNIQUE'] = f"{pct(connected_unique, accounts)*100:.2f}%"
    row['RPC %'] = f"{pct(total_rpc, connected_unique)*100:.2f}%" if connected_unique > 0 else "0.00%"
    row['PTP %'] = f"{pct(ptp, total_rpc)*100:.2f}%" if total_rpc > 0 else "0.00%"
    row['CALL DROP RATE'] = f"{pct(system_drop, connected_nu_float)*100:.2f}%" if connected_nu_float > 0 else "0.00%"

    return pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)

def calculate_connected_status(raw_df):
    if raw_df.empty or 'DEBTOR ID' not in raw_df.columns or 'TALK TIME DURATION' not in raw_df.columns or 'STATUS' not in raw_df.columns:
        return pd.DataFrame({'CONNECTED UNIQUE': [0]})

    connected_calls = raw_df[raw_df['TALK TIME DURATION'] > 0].copy()
    last_connected_status = connected_calls.drop_duplicates(subset=['DEBTOR ID'], keep='last')
    total_connected_unique = last_connected_status['DEBTOR ID'].nunique()
    status_counts = last_connected_status.groupby('STATUS')['DEBTOR ID'].nunique().to_dict()
    
    output_data = {'CONNECTED UNIQUE': total_connected_unique}
    output_data.update(status_counts)
    return pd.DataFrame([output_data])

# ----------------------------------------------------------------------
# MAIN APP
# ----------------------------------------------------------------------
uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_combined, all_predictive, all_manual = [], [], []
    predictive_cycles, manual_cycles = {}, {}
    predictive_balances, manual_balances = {}, {}
    raw_predictive_all = pd.DataFrame()
    raw_combined_all = pd.DataFrame()

    prog = st.progress(0)
    for idx, f in enumerate(uploaded_files):
        with st.spinner(f"Processing {f.name}..."):
            df = load_data(f)
            df = process_file(df)

            follow_up = df[(df['REMARK TYPE'] == 'Follow Up') & (df['REMARK'].str.contains('Predictive', case=False, na=False))] if 'REMARK TYPE' in df.columns and 'REMARK' in df.columns else pd.DataFrame()
            predictive = df[df['REMARK TYPE'] == 'Predictive'] if 'REMARK TYPE' in df.columns else pd.DataFrame()
            outgoing = df[df['REMARK TYPE'] == 'Outgoing'] if 'REMARK TYPE' in df.columns else pd.DataFrame()

            predictive_combined_df = pd.concat([follow_up, predictive], ignore_index=True)
            combined_df = pd.concat([predictive_combined_df, outgoing], ignore_index=True)

            raw_predictive_all = pd.concat([raw_predictive_all, predictive_combined_df], ignore_index=True)
            raw_combined_all = pd.concat([raw_combined_all, combined_df], ignore_index=True)

            all_combined.append(calculate_summary(combined_df, ['Predictive', 'Follow Up', 'Outgoing']))
            all_predictive.append(calculate_summary(predictive_combined_df, ['Predictive', 'Follow Up']))
            all_manual.append(calculate_summary(outgoing, ['Outgoing']))

            for k, v in get_cycle_summary(predictive_combined_df, ['Predictive', 'Follow Up']).items():
                predictive_cycles[k] = pd.concat([predictive_cycles.get(k, pd.DataFrame()), v], ignore_index=True)
            for k, v in get_cycle_summary(outgoing, ['Outgoing']).items():
                manual_cycles[k] = pd.concat([manual_cycles.get(k, pd.DataFrame()), v], ignore_index=True)

            for k, v in get_balance_summary(predictive_combined_df, ['Predictive', 'Follow Up']).items():
                predictive_balances[k] = pd.concat([predictive_balances.get(k, pd.DataFrame()), v], ignore_index=True)
            for k, v in get_balance_summary(outgoing, ['Outgoing']).items():
                manual_balances[k] = pd.concat([manual_balances.get(k, pd.DataFrame()), v], ignore_index=True)

        prog.progress((idx + 1) / len(uploaded_files))

    prog.empty()
    st.success(f"Processed {len(uploaded_files)} file(s)")

    combined_summary = pd.concat(all_combined, ignore_index=True).sort_values('DATE') if all_combined else pd.DataFrame(columns=SUMMARY_COLUMNS)
    predictive_summary = pd.concat(all_predictive, ignore_index=True).sort_values('DATE') if all_predictive else pd.DataFrame(columns=SUMMARY_COLUMNS)
    manual_summary = pd.concat(all_manual, ignore_index=True).sort_values('DATE') if all_manual else pd.DataFrame(columns=SUMMARY_COLUMNS)

    combined_cycle = combine_summaries(predictive_cycles, manual_cycles)
    combined_balance = combine_summaries(predictive_balances, manual_balances)
    
    predictive_cycles_total = aggregate_cycles_by_cycle(predictive_cycles, raw_predictive_all)
    
    overall_combined_total = aggregate_overall_summary(raw_combined_all, combined_summary)
    connected_status_summary = calculate_connected_status(raw_combined_all)

    # Display
    st.write("## Overall Combined Summary (Daily)")
    st.dataframe(combined_summary, use_container_width=True)

    if not overall_combined_total.empty:
        st.write("## Overall Combined Summary (Total Period)")
        st.dataframe(overall_combined_total, use_container_width=True)

    if not connected_status_summary.empty:
        st.write("## Connected Status (Total Period Breakdown)")
        st.dataframe(connected_status_summary.T.rename(columns={0: 'Count'}), use_container_width=True)

    if predictive_cycles_total:
        st.write("## Predictive Cycles (Total) with Status Breakdown")
        cycle_keys = sorted([k for k in predictive_cycles_total.keys() if not k.endswith(CONNECTED_STATUS_CYCLE_TITLE)], 
                            key=lambda x: int(re.search(r'Cycle (\d+)', x).group(1)) if re.search(r'Cycle (\d+)', x) else float('inf'))
        for cycle_key in cycle_keys:
            st.subheader(f"{cycle_key} Summary")
            st.dataframe(predictive_cycles_total[cycle_key], use_container_width=True)
            
            breakdown_key = f"{cycle_key} {CONNECTED_STATUS_CYCLE_TITLE}"
            if breakdown_key in predictive_cycles_total and not predictive_cycles_total[breakdown_key].empty:
                st.write(f"### {cycle_key} Connected Status Breakdown")
                st.dataframe(predictive_cycles_total[breakdown_key].T.rename(columns={0: 'Count'}), use_container_width=True)

    for title, df in [("Overall Predictive", predictive_summary), ("Overall Manual", manual_summary)]:
        if not df.empty:
            st.write(f"## {title} Summary")
            st.dataframe(df, use_container_width=True)

    # Summary Export
    summary_groups = {
        'Combined': {'Combined Summary (Daily)': combined_summary},
        'Overall Combined Summary': {
            'Overall Combined Summary (Total)': overall_combined_total,
            CONNECTED_STATUS_TITLE: connected_status_summary
        },
        'Predictive': {'Predictive Summary': predictive_summary},
        'Manual': {'Manual Summary': manual_summary},
        'Combined Cycles': {k: v for k, v in combined_cycle.items() if "na" not in k.lower()},
        'Predictive Cycles': {k: v for k, v in predictive_cycles.items() if "na" not in k.lower()},
        'Manual Cycles': {k: v for k, v in manual_cycles.items() if "na" not in k.lower()},
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

    # Raw Breakdown
    raw_cycle_bytes = create_raw_cycle_breakdown_excel(
        raw_predictive_all,
        raw_combined_all,
        predictive_cycles,
        overall_combined_total
    )

    st.sidebar.download_button(
        label="Download Raw Cycle Breakdown",
        data=raw_cycle_bytes,
        file_name=f"Raw_Cycle_Breakdown_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="raw_cycle_btn"
    )

else:
    st.info("Upload one or more XLSX files to start.")
