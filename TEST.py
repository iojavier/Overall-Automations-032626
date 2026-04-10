"""
Debt Collection Summary (Volare) — Streamlit App
Optimized: deduplicated aggregation logic, vectorized percentage calc,
           list-then-concat accumulation, pre-compiled regex, column-type
           dispatch in Excel writer.
"""
import re
import datetime
from io import BytesIO

import pandas as pd
import polars as pl
from pandas import ExcelWriter
import streamlit as st

# ── Constants ────────────────────────────────────────────────────────────────

SUMMARY_COLUMNS = [
    'CYCLE', 'DATE', 'CLIENT',
    'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE',
    'CONNECTED NU', 'CONNECTED UNIQUE', 'TOTAL RPC', 'PTP',
    'CONNECTED % NU', 'CONNECTED % UNIQUE', 'RPC %', 'PTP %',
    'TOTAL TALK TIME', 'TALK TIME AVE',
    'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP', 'CALL DROP RATE',
]

PERCENTAGE_COLS = [
    'PENETRATION RATE', 'CONNECTED % NU', 'CONNECTED % UNIQUE',
    'RPC %', 'PTP %', 'CALL DROP RATE',
]

NUMERICAL_COLS = [
    'ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
    'TOTAL RPC', 'PTP', 'TOTAL BALANCE', 'NEG DROP', 'SYSTEM DROP',
]

BALANCE_RANGES = [
    (0,        49_999.99,    "0-49999.99"),
    (50_000,   99_999.99,    "50000.00-99999.99"),
    (100_000,  float('inf'), "100000.00 and up"),
]

SHEET_ORDER = [
    'Combined', 'Overall Combined Summary', 'Overall Combined Cycles',
    'Predictive', 'Manual',
    'Combined Cycles', 'Predictive Cycles', 'Manual Cycles',
    'Predictive Cycles (Total)',
    'Combined Balances', 'Predictive Balances', 'Manual Balances',
]

# Pre-compiled regex — avoids recompiling on every row/call
_RE_PTP_NEW   = re.compile(r'1_\d{11} - PTP NEW', re.IGNORECASE)
_RE_EXCLUDED  = re.compile(
    r'Broken Promise|New files imported'
    r'|Updates when case reassign to another collector'
    r'|NDF IN ICS|FOR PULL OUT|END OF HANDLING PERIOD'
    r'|New Assignment -|broadcast|File Unhold',
    re.IGNORECASE,
)
_RE_RPC_ESC   = re.compile(r'bank escalation|rpc', re.IGNORECASE)
_RE_NEG_DROP  = re.compile(r'NEGATIVE CALLOUTS - CALL DROP', re.IGNORECASE)
_RE_DROPPED   = re.compile(r'DROPPED', re.IGNORECASE)
_RE_CYCLE_NUM = re.compile(r'Cycle (\d+)')
_RE_PTP_NEW_STR  = r'1_\d{11} - PTP NEW'
_RE_EXCLUDED_STR = (
    r'Broken Promise|New files imported'
    r'|Updates when case reassign to another collector'
    r'|NDF IN ICS|FOR PULL OUT|END OF HANDLING PERIOD'
    r'|New Assignment -|broadcast|File Unhold'
)
_RE_RPC_ESC_STR  = r'bank escalation|rpc'
_RE_NEG_DROP_STR = r'NEGATIVE CALLOUTS - CALL DROP'
_RE_DROPPED_STR  = r'DROPPED'

_PCT_SET  = set(PERCENTAGE_COLS)
_NUM_SET  = set(NUMERICAL_COLS)
_TIME_SET = {'TOTAL TALK TIME', 'TALK TIME AVE'}

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    layout="wide",
    page_title="Debt Collection Summary (Volare)",
    page_icon="📊",
    initial_sidebar_state="expanded",
)
st.title('Debt Collection Summary')

# ── Utility helpers ───────────────────────────────────────────────────────────

def _pct(n, d) -> float:
    return n / d if d else 0.0

def _empty_summary() -> pd.DataFrame:
    return pd.DataFrame(columns=SUMMARY_COLUMNS)

def _empty_pl_df() -> pl.DataFrame:
    return pl.DataFrame()

def _cycle_sort_key(k: str) -> int:
    m = _RE_CYCLE_NUM.search(k)
    return int(m.group(1)) if m else 9999

def _valid_cycles(series: pd.Series) -> list:
    bad = {'', 'unknown', 'na'}
    return [c for c in series.unique() if str(c).lower() not in bad]

def _valid_cycles_pl(series: pl.Series) -> list:
    bad = {'', 'unknown', 'na'}
    return [c for c in series.unique().to_list() if str(c).lower() not in bad]

def _pct_str_to_float(series: pd.Series) -> pd.Series:
    """'12.34%' → 0.1234"""
    return (
        series.astype(str)
        .str.rstrip('%')
        .str.replace(',', '', regex=False)
        .pipe(pd.to_numeric, errors='coerce')
        .fillna(0)
        / 100
    )

def format_seconds_to_hms(seconds) -> str:
    if pd.isna(seconds) or seconds <= 0:
        return "00:00:00"
    s = int(seconds)
    h, s = divmod(s, 3600)
    m, s = divmod(s, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def hms_to_seconds(hms) -> int:
    v = str(hms).strip()
    if not v or v == "00:00:00":
        return 0
    try:
        h, m, s = v.split(':')
        return int(h) * 3600 + int(m) * 60 + int(s)
    except Exception:
        return 0

def _format_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return _empty_summary()

    df = df.copy()
    talk_secs = pd.to_numeric(df.pop('_TALK_SECS'), errors='coerce').fillna(0) if '_TALK_SECS' in df.columns else pd.Series(0, index=df.index)

    accounts = pd.to_numeric(df['ACCOUNTS'], errors='coerce').fillna(0)
    total_dialed = pd.to_numeric(df['TOTAL DIALED'], errors='coerce').fillna(0)
    connected_nu = pd.to_numeric(df['CONNECTED NU'], errors='coerce').fillna(0)
    connected_unique = pd.to_numeric(df['CONNECTED UNIQUE'], errors='coerce').fillna(0)
    total_rpc = pd.to_numeric(df['TOTAL RPC'], errors='coerce').fillna(0)
    ptp = pd.to_numeric(df['PTP'], errors='coerce').fillna(0)
    system_drop = pd.to_numeric(df['SYSTEM DROP'], errors='coerce').fillna(0)

    df['PENETRATION RATE'] = (total_dialed / accounts.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['CONNECTED % NU'] = (connected_nu / total_dialed.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['CONNECTED % UNIQUE'] = (connected_unique / accounts.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['RPC %'] = (total_rpc / connected_unique.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['PTP %'] = (ptp / total_rpc.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['CALL DROP RATE'] = (system_drop / connected_nu.replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
    df['TOTAL TALK TIME'] = talk_secs.map(format_seconds_to_hms)
    talk_avg = (talk_secs / connected_nu.replace(0, pd.NA)).fillna(0)
    df['TALK TIME AVE'] = talk_avg.map(format_seconds_to_hms)

    return df.reindex(columns=SUMMARY_COLUMNS, fill_value=0).sort_values('DATE')

# ── Data loading / cleaning ───────────────────────────────────────────────────

@st.cache_data
def load_data(file_name: str, file_bytes: bytes) -> pl.DataFrame:
    try:
        df = pl.read_excel(BytesIO(file_bytes), engine="calamine")
    except Exception:
        pdf = pd.read_excel(BytesIO(file_bytes))
        df = pl.from_pandas(pdf)

    renamed = {col: col.strip().upper() for col in df.columns}
    df = df.rename(renamed)
    if 'DATE' in df.columns:
        df = df.with_columns(pl.col('DATE').cast(pl.Datetime, strict=False).alias('DATE'))
    return df


def process_file(df: pl.DataFrame) -> pl.DataFrame:
    for col in ('DEBTOR', 'STATUS', 'REMARK', 'CALL STATUS', 'CARD NO.', 'REMARK BY', 'REMARK TYPE', 'CLIENT', 'DEBTOR ID'):
        if col in df.columns:
            df = df.with_columns(pl.col(col).fill_null('').cast(pl.Utf8, strict=False).alias(col))

    mask = pl.lit(True)
    if 'REMARK BY' in df.columns:
        mask = mask & (pl.col('REMARK BY') != 'SPMADRID')
    if 'DEBTOR' in df.columns:
        mask = mask & (~pl.col('DEBTOR').str.contains("(?i)DEFAULT_LEAD_"))
    if 'STATUS' in df.columns:
        mask = mask & (~pl.col('STATUS').str.contains('ABORT'))
    if 'REMARK' in df.columns:
        mask = mask & (~pl.col('REMARK').str.contains(f"(?i){_RE_PTP_NEW_STR}"))
        mask = mask & (~pl.col('REMARK').str.contains(f"(?i){_RE_EXCLUDED_STR}"))
    df = df.filter(mask)

    if 'CARD NO.' in df.columns:
        df = df.with_columns(
            pl.when(pl.col('CARD NO.').str.slice(0, 2) == '')
            .then(pl.lit('Unknown'))
            .otherwise(pl.col('CARD NO.').str.slice(0, 2))
            .alias('CYCLE')
        )

    for col in ('CALL DURATION', 'TALK TIME DURATION', 'PTP AMOUNT', 'BALANCE'):
        expr = pl.col(col).cast(pl.Float64, strict=False).fill_null(0) if col in df.columns else pl.lit(0.0)
        df = df.with_columns(expr.alias(col))

    derived_exprs = []
    derived_exprs.append((pl.col('TALK TIME DURATION') > 0).alias('_IS_CONNECTED'))
    derived_exprs.append((pl.col('PTP AMOUNT') > 0).alias('_IS_PTP'))
    if 'STATUS' in df.columns:
        derived_exprs.extend([
            pl.col('STATUS').str.contains(f"(?i){_RE_RPC_ESC_STR}").alias('_IS_RPC'),
            pl.col('STATUS').str.contains(f"(?i){_RE_NEG_DROP_STR}").alias('_IS_NEG_DROP'),
            pl.col('STATUS').str.contains(f"(?i){_RE_DROPPED_STR}").alias('_IS_DROPPED'),
        ])
    else:
        derived_exprs.extend([
            pl.lit(False).alias('_IS_RPC'),
            pl.lit(False).alias('_IS_NEG_DROP'),
            pl.lit(False).alias('_IS_DROPPED'),
        ])
    if derived_exprs:
        df = df.with_columns(derived_exprs)
    return df


def apply_date_filter(df: pl.DataFrame, start, end) -> pl.DataFrame:
    if 'DATE' not in df.columns or df.is_empty():
        return df
    lo = pd.Timestamp(start)
    hi = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return df.filter((pl.col('DATE') >= lo) & (pl.col('DATE') <= hi))

# ── Metric / summary calculation ──────────────────────────────────────────────

def _summary_agg_exprs() -> list[pl.Expr]:
    return [
        pl.col('DEBTOR ID').n_unique().alias('ACCOUNTS'),
        pl.len().alias('TOTAL DIALED'),
        pl.col('_IS_CONNECTED').sum().cast(pl.Int64).alias('CONNECTED NU'),
        pl.col('DEBTOR ID').filter(pl.col('_IS_CONNECTED')).n_unique().alias('CONNECTED UNIQUE'),
        pl.col('DEBTOR ID').filter(pl.col('_IS_PTP') | pl.col('_IS_RPC')).n_unique().alias('TOTAL RPC'),
        pl.col('DEBTOR ID').filter(pl.col('_IS_PTP') & pl.col('_IS_CONNECTED')).n_unique().alias('PTP'),
        pl.col('TALK TIME DURATION').sum().alias('_TALK_SECS'),
        pl.col('BALANCE').filter(pl.col('_IS_PTP')).sum().alias('TOTAL BALANCE'),
        pl.col('_IS_NEG_DROP').sum().cast(pl.Int64).alias('NEG DROP'),
        pl.col('_IS_DROPPED').sum().cast(pl.Int64).alias('SYSTEM DROP'),
    ]


def _aggregate_summary(sub: pl.DataFrame) -> pd.DataFrame:
    required = {'DATE', 'CLIENT', 'DEBTOR ID'}
    if sub.is_empty() or not required.issubset(set(sub.columns)):
        return _empty_summary()

    grouped = (
        sub
        .with_columns(pl.col('DATE').cast(pl.Date, strict=False).alias('DATE'))
        .group_by(['DATE', 'CLIENT'])
        .agg(_summary_agg_exprs())
        .sort('DATE')
    )
    if grouped.is_empty():
        return _empty_summary()

    return _format_summary_df(grouped.to_pandas())


def calculate_summary(df: pl.DataFrame, remark_types: list) -> pd.DataFrame:
    if 'REMARK TYPE' not in df.columns:
        return _empty_summary()
    sub = df.filter(pl.col('REMARK TYPE').is_in(remark_types))
    return _aggregate_summary(sub)


def get_cycle_summary(df: pl.DataFrame, remark_types: list) -> dict:
    if 'CYCLE' not in df.columns:
        return {}
    result = {}
    for cycle in _valid_cycles_pl(df.get_column('CYCLE')):
        s = calculate_summary(df.filter(pl.col('CYCLE') == cycle), remark_types)
        if not s.empty:
            s['CYCLE'] = cycle
            result[f"Cycle {cycle}"] = s
    return result


def get_balance_summary(df: pl.DataFrame, remark_types: list) -> dict:
    if 'CYCLE' not in df.columns or 'BALANCE' not in df.columns:
        return {}
    df = df.with_columns(
        pl.when((pl.col('BALANCE') >= BALANCE_RANGES[0][0]) & (pl.col('BALANCE') <= BALANCE_RANGES[0][1]))
        .then(pl.lit(BALANCE_RANGES[0][2]))
        .when((pl.col('BALANCE') >= BALANCE_RANGES[1][0]) & (pl.col('BALANCE') <= BALANCE_RANGES[1][1]))
        .then(pl.lit(BALANCE_RANGES[1][2]))
        .when(pl.col('BALANCE') >= BALANCE_RANGES[2][0])
        .then(pl.lit(BALANCE_RANGES[2][2]))
        .otherwise(None)
        .alias('_BALANCE_BUCKET')
    )
    result = {}
    cycles = _valid_cycles_pl(df.get_column('CYCLE'))
    for cycle in cycles:
        c_df = df.filter(pl.col('CYCLE') == cycle)
        for _, _, name in BALANCE_RANGES:
            s = calculate_summary(c_df.filter(pl.col('_BALANCE_BUCKET') == name), remark_types)
            if not s.empty:
                s['CYCLE'] = cycle
                result[f"Cycle {cycle} Balance {name}"] = s
    return result

# ── Aggregation — shared core ────────────────────────────────────────────────

def _build_aggregate_row(raw: pd.DataFrame, daily: pd.DataFrame) -> dict:
    """
    Single source of truth for rolling up raw + daily data into one metric row.
    Used by all three aggregate_* functions below.
    """
    conn_m = raw['_IS_CONNECTED'] if '_IS_CONNECTED' in raw.columns else (raw['TALK TIME DURATION'] > 0)
    ptp_m  = raw['_IS_PTP'] if '_IS_PTP' in raw.columns else (raw['PTP AMOUNT'] > 0)
    rpc_m  = raw['_IS_RPC'] if '_IS_RPC' in raw.columns else raw['STATUS'].str.contains(_RE_RPC_ESC, na=False)

    accounts         = raw['DEBTOR ID'].nunique()
    connected_unique = raw.loc[conn_m, 'DEBTOR ID'].nunique()
    total_rpc        = raw.loc[ptp_m | rpc_m, 'DEBTOR ID'].nunique()
    ptp              = raw.loc[ptp_m & conn_m, 'DEBTOR ID'].nunique()

    daily = daily.copy()
    daily['_secs'] = daily['TOTAL TALK TIME'].map(hms_to_seconds)
    total_dialed  = int(daily['TOTAL DIALED'].sum())
    connected_nu  = int(daily['CONNECTED NU'].sum())
    talk_secs     = int(daily['_secs'].sum())
    neg_drop      = int(daily['NEG DROP'].sum())
    system_drop   = int(daily['SYSTEM DROP'].sum())
    total_balance = float(daily['TOTAL BALANCE'].sum())

    cnuf = float(connected_nu)
    return {
        'ACCOUNTS':           accounts,
        'TOTAL DIALED':       total_dialed,
        'CONNECTED NU':       connected_nu,
        'CONNECTED UNIQUE':   connected_unique,
        'TOTAL RPC':          total_rpc,
        'PTP':                ptp,
        'TOTAL TALK TIME':    format_seconds_to_hms(talk_secs),
        'TALK TIME AVE':      format_seconds_to_hms(talk_secs / cnuf) if cnuf else "00:00:00",
        'TOTAL BALANCE':      total_balance,
        'NEG DROP':           neg_drop,
        'SYSTEM DROP':        system_drop,
        'PENETRATION RATE':   f"{_pct(total_dialed, accounts)*100:.2f}%",
        'CONNECTED % NU':     f"{_pct(cnuf, total_dialed)*100:.2f}%",
        'CONNECTED % UNIQUE': f"{_pct(connected_unique, accounts)*100:.2f}%",
        'RPC %':              f"{_pct(total_rpc, connected_unique)*100:.2f}%" if connected_unique else "0.00%",
        'PTP %':              f"{_pct(ptp, total_rpc)*100:.2f}%"              if total_rpc        else "0.00%",
        'CALL DROP RATE':     f"{_pct(system_drop, cnuf)*100:.2f}%"           if cnuf             else "0.00%",
    }


def aggregate_cycles_by_cycle(cycle_dict: dict, raw_df: pd.DataFrame) -> dict:
    if not cycle_dict or raw_df.empty:
        return {}
    out = {}
    for cycle in sorted({_RE_CYCLE_NUM.search(k).group(1)
                         for k in cycle_dict if _RE_CYCLE_NUM.search(k)}):
        key   = f"Cycle {cycle}"
        raw_c = raw_df[raw_df['CYCLE'] == cycle]
        daily = cycle_dict.get(key, _empty_summary())
        if raw_c.empty or daily.empty:
            continue
        row = {**_build_aggregate_row(raw_c, daily), 'CYCLE': cycle}
        out[key] = pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    return out


def aggregate_overall_summary(raw_df: pd.DataFrame, daily_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty or daily_df.empty:
        return _empty_summary()
    row   = _build_aggregate_row(raw_df, daily_df)
    dates = pd.to_datetime(daily_df['DATE'])
    d_lo, d_hi = dates.min(), dates.max()
    date_range = (f"{d_lo:%Y-%m-%d} to {d_hi:%Y-%m-%d}" if not pd.isna(d_lo) else "N/A")
    row.update({'CYCLE': 'All', 'DATE': date_range, 'CLIENT': 'All'})
    return pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)


def aggregate_overall_combined_cycles(combined_cycle_dict: dict, raw_combined: pd.DataFrame) -> dict:
    if not combined_cycle_dict or raw_combined.empty:
        return {}
    result = {}
    for key, daily_df in combined_cycle_dict.items():
        m = _RE_CYCLE_NUM.search(key)
        if not m:
            continue
        cycle = m.group(1)
        raw_c = raw_combined[raw_combined['CYCLE'] == cycle]
        if raw_c.empty or daily_df.empty:
            continue
        row = {**_build_aggregate_row(raw_c, daily_df), 'CYCLE': cycle}
        result[f"Cycle {cycle}"] = pd.DataFrame([row]).reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    return result


def combine_summaries(pred: dict, man: dict) -> dict:
    combined = {}
    for key in set(pred) | set(man):
        if 'na' in key.lower():
            continue
        df = pd.concat(
            [pred.get(key, _empty_summary()), man.get(key, _empty_summary())],
            ignore_index=True,
        )
        if df.empty:
            continue

        df['DATE']  = pd.to_datetime(df['DATE']).dt.date
        df['_secs'] = df['TOTAL TALK TIME'].map(hms_to_seconds)

        for col in PERCENTAGE_COLS:
            if col in df.columns:
                df[col] = _pct_str_to_float(df[col])

        agg = {c: 'sum' for c in NUMERICAL_COLS if c in df.columns}
        agg['_secs'] = 'sum'
        g = df.groupby(['DATE', 'CLIENT'], as_index=False).agg(agg)

        # Vectorised derived columns
        g['TOTAL TALK TIME']    = g['_secs'].map(format_seconds_to_hms)
        g['TALK TIME AVE']      = (g['_secs'] / g['CONNECTED NU'].replace(0, pd.NA)).fillna(0).map(format_seconds_to_hms)
        g['PENETRATION RATE']   = (g['TOTAL DIALED']     / g['ACCOUNTS'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
        g['CONNECTED % NU']     = (g['CONNECTED NU']     / g['TOTAL DIALED'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
        g['CONNECTED % UNIQUE'] = (g['CONNECTED UNIQUE'] / g['ACCOUNTS'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
        g['RPC %']              = (g['TOTAL RPC']        / g['CONNECTED UNIQUE'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
        g['PTP %']              = (g['PTP']              / g['TOTAL RPC'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)
        g['CALL DROP RATE']     = (g['SYSTEM DROP']      / g['CONNECTED NU'].replace(0, pd.NA)).fillna(0).mul(100).map("{:.2f}%".format)

        combined[key] = (
            g.drop(columns=['_secs'], errors='ignore')
             .reindex(columns=SUMMARY_COLUMNS, fill_value=0)
             .sort_values('DATE')
        )
    return combined

# ── Excel writing ─────────────────────────────────────────────────────────────

def _excel_formats(wb) -> dict:
    base = {'align': 'center', 'valign': 'vcenter', 'border': 1}
    return {
        'title':   wb.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                  'valign': 'vcenter', 'bg_color': '#A42A25', 'font_color': 'white'}),
        'header':  wb.add_format({'bold': True, 'bg_color': '#D9B229', 'font_color': 'white',
                                  'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'comma':   wb.add_format({**base, 'num_format': '#,##0'}),
        'percent': wb.add_format({**base, 'num_format': '0.00%'}),
        'date':    wb.add_format({**base, 'num_format': 'yyyy-mm-dd'}),
        'time':    wb.add_format({**base, 'num_format': 'hh:mm:ss'}),
        'center':  wb.add_format(base),
    }


def _write_cell(ws, row: int, col: int, col_name: str, value, fmts: dict):
    """Single-cell write dispatched by column type — no branching overhead per sheet."""
    if col_name == 'DATE':
        try:
            ws.write_datetime(row, col, pd.to_datetime(value).to_pydatetime(), fmts['date'])
        except Exception:
            ws.write(row, col, str(value), fmts['center'])
    elif col_name == 'TOTAL BALANCE':
        ws.write_number(row, col, float(value), fmts['comma'])
    elif col_name in _PCT_SET:
        try:
            ws.write_number(row, col, float(value), fmts['percent'])
        except Exception:
            try:
                ws.write_number(row, col, float(str(value).rstrip('%').replace(',', '')) / 100, fmts['percent'])
            except Exception:
                ws.write(row, col, str(value), fmts['center'])
    elif col_name in _TIME_SET:
        ws.write_string(row, col, str(value), fmts['time'])
    elif col_name in _NUM_SET:
        try:
            ws.write_number(row, col, float(value), fmts['comma'])
        except Exception:
            ws.write(row, col, value, fmts['center'])
    else:
        ws.write(row, col, value, fmts['center'])


def _sorted_sheet_items(sheet_name: str, df_dict: dict) -> list:
    if sheet_name == 'Overall Combined Summary':
        main  = next((k for k in df_dict if 'Total' in k), None)
        items = [(main, df_dict[main])] if main else []
        items += [(k, v) for k, v in df_dict.items() if k != main]
        return items
    if sheet_name in ('Overall Combined Cycles', 'Predictive Cycles (Total)'):
        keys = sorted([k for k in df_dict if k.startswith('Cycle ')], key=_cycle_sort_key)
        return [(k, df_dict[k]) for k in keys]
    return sorted(df_dict.items(), key=lambda x: _cycle_sort_key(x[0]))


def write_excel_sheet(writer, sheet_name: str, df_dict: dict, fmts: dict):
    if not df_dict:
        return
    ws  = writer.book.add_worksheet(sheet_name)
    cur = 0
    for title, df in _sorted_sheet_items(sheet_name, df_dict):
        if df.empty:
            continue
        display = df.copy()
        if 'DATE' in display.columns:
            display = display.sort_values('DATE', na_position='first')
        for col in NUMERICAL_COLS:
            if col in display.columns:
                display[col] = pd.to_numeric(display[col], errors='coerce').fillna(0)
        for col in PERCENTAGE_COLS:
            if col in display.columns:
                display[col] = _pct_str_to_float(display[col])

        ws.merge_range(cur, 0, cur, len(display.columns) - 1, title, fmts['title'])
        cur += 1
        for ci, cn in enumerate(display.columns):
            ws.write(cur, ci, cn, fmts['header'])
            ws.set_column(ci, ci, 20)
        cur += 1
        for ri in range(len(display)):
            for ci, cn in enumerate(display.columns):
                _write_cell(ws, cur + ri, ci, cn, display.iat[ri, ci], fmts)
        cur += len(display) + 2


@st.cache_data(show_spinner=False)
def to_excel(summary_groups: dict) -> bytes:
    output  = BytesIO()
    ordered = {k: summary_groups[k] for k in SHEET_ORDER if k in summary_groups}
    with ExcelWriter(output, engine='xlsxwriter') as writer:
        fmts = _excel_formats(writer.book)
        for sheet_name, df_dict in ordered.items():
            write_excel_sheet(writer, sheet_name, df_dict, fmts)
    return output.getvalue()


@st.cache_data(show_spinner=False)
def create_raw_cycle_breakdown_excel(
    raw_pred, raw_combined, pred_cycles_dict, overall_total_df, overall_cycles_dict
) -> bytes:
    RAW_HDRS  = ['ACCOUNTS', 'TOTAL DIALED', 'CONNECTED NU', 'CONNECTED UNIQUE',
                 'TOTAL RPC', 'PTP', 'NEG DROP', 'SYSTEM DROP']
    LIST_HDRS = ['Raw Accounts', 'Raw Connected NU', 'Raw Connected Unique',
                 'Raw Total RPC', 'Raw PTP', 'Raw Negative Drop', 'Raw System Drop']

    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter') as writer:
        wb      = writer.book
        bold    = wb.add_format({'bold': True, 'font_size': 14})
        hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#E6F0FA', 'border': 1, 'align': 'center'})
        num_fmt = wb.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
        ctr_fmt = wb.add_format({'align': 'center', 'border': 1})
        no_data = wb.add_format({'bold': True, 'font_size': 12})

        def write_cycle_block(ws, cycle_key, summary_row, raw_df, start, is_combined=False):
            ws.write(start, 0, f"{cycle_key} (Combined)" if is_combined else cycle_key, bold)
            r = start + 2

            for ci, h in enumerate(RAW_HDRS):
                ws.write(r, ci, h, hdr_fmt)
            r += 1
            for ci, h in enumerate(RAW_HDRS):
                ws.write_number(r, ci, float(summary_row.get(h, 0)), num_fmt)
            r += 3

            for ci, h in enumerate(LIST_HDRS):
                ws.write(r, ci, h, hdr_fmt)
            r += 1

            if raw_df.empty:
                ws.write(r, 0, "(No raw data)", no_data)
                return r + 5

            conn_m = raw_df['_IS_CONNECTED'] if '_IS_CONNECTED' in raw_df.columns else (raw_df['TALK TIME DURATION'] > 0)
            ptp_m  = raw_df['_IS_PTP'] if '_IS_PTP' in raw_df.columns else (raw_df['PTP AMOUNT'] > 0)
            rpc_m  = raw_df['_IS_RPC'] if '_IS_RPC' in raw_df.columns else raw_df['STATUS'].str.contains(_RE_RPC_ESC, na=False)
            neg_m  = raw_df['_IS_NEG_DROP'] if '_IS_NEG_DROP' in raw_df.columns else raw_df['STATUS'].str.contains(_RE_NEG_DROP, na=False)
            drp_m  = raw_df['_IS_DROPPED'] if '_IS_DROPPED' in raw_df.columns else raw_df['STATUS'].str.contains(_RE_DROPPED,  na=False)

            lists = [
                sorted(raw_df['CARD NO.'].dropna().unique()),
                list(raw_df.loc[conn_m, 'CARD NO.']),
                sorted(raw_df.loc[conn_m, 'CARD NO.'].dropna().unique()),
                sorted(raw_df.loc[ptp_m | rpc_m, 'CARD NO.'].dropna().unique()),
                sorted(raw_df.loc[ptp_m & conn_m, 'CARD NO.'].dropna().unique()),
                sorted(raw_df.loc[neg_m, 'CARD NO.'].dropna().unique()),
                sorted(raw_df.loc[drp_m, 'CARD NO.'].dropna().unique()),
            ]
            max_r = max((len(l) for l in lists), default=0)
            for ri in range(max_r):
                for ci, lst in enumerate(lists):
                    if ri < len(lst):
                        ws.write(r + ri, ci, str(lst[ri]), ctr_fmt)
            for ci in range(len(LIST_HDRS)):
                ws.set_column(ci, ci, 28)
            return r + max_r + 5

        # Sheet 1 — Predictive Cycles Raw
        ws1 = wb.add_worksheet("Predictive Cycles Raw")
        row = 0
        totals = aggregate_cycles_by_cycle(pred_cycles_dict, raw_pred)
        for ck in sorted((k for k in totals if k.startswith('Cycle ')), key=_cycle_sort_key):
            sdf = totals[ck]
            if sdf.empty:
                continue
            cyc = _RE_CYCLE_NUM.search(ck).group(1)
            row = write_cycle_block(ws1, ck, sdf.iloc[0], raw_pred[raw_pred['CYCLE'] == cyc], row)

        # Sheet 2 — Overall Combined Raw
        ws2 = wb.add_worksheet("Overall Combined Raw")
        if not overall_total_df.empty:
            ws2.write(0, 0, "Overall Combined Summary (Total Period)", bold)
            for ci, h in enumerate(RAW_HDRS):
                ws2.write(2, ci, h, hdr_fmt)
            for ci, h in enumerate(RAW_HDRS):
                ws2.write_number(3, ci, float(overall_total_df.iloc[0].get(h, 0)), num_fmt)
            row = 6
            for ci, h in enumerate(LIST_HDRS):
                ws2.write(row, ci, h, hdr_fmt)
            row += 1

            if raw_combined.empty:
                ws2.write(row, 0, "(No raw data)", no_data)
            else:
                conn_m = raw_combined['_IS_CONNECTED'] if '_IS_CONNECTED' in raw_combined.columns else (raw_combined['TALK TIME DURATION'] > 0)
                ptp_m  = raw_combined['_IS_PTP'] if '_IS_PTP' in raw_combined.columns else (raw_combined['PTP AMOUNT'] > 0)
                rpc_m  = raw_combined['_IS_RPC'] if '_IS_RPC' in raw_combined.columns else raw_combined['STATUS'].str.contains(_RE_RPC_ESC, na=False)
                neg_m  = raw_combined['_IS_NEG_DROP'] if '_IS_NEG_DROP' in raw_combined.columns else raw_combined['STATUS'].str.contains(_RE_NEG_DROP, na=False)
                drp_m  = raw_combined['_IS_DROPPED'] if '_IS_DROPPED' in raw_combined.columns else raw_combined['STATUS'].str.contains(_RE_DROPPED,  na=False)

                lists = [
                    sorted(raw_combined['CARD NO.'].dropna().unique()),
                    list(raw_combined.loc[conn_m, 'CARD NO.']),
                    sorted(raw_combined.loc[conn_m, 'CARD NO.'].dropna().unique()),
                    sorted(raw_combined.loc[ptp_m | rpc_m, 'CARD NO.'].dropna().unique()),
                    sorted(raw_combined.loc[ptp_m & conn_m, 'CARD NO.'].dropna().unique()),
                    sorted(raw_combined.loc[neg_m, 'CARD NO.'].dropna().unique()),
                    sorted(raw_combined.loc[drp_m, 'CARD NO.'].dropna().unique()),
                ]
                max_r = max((len(l) for l in lists), default=0)
                for ri in range(max_r):
                    for ci, lst in enumerate(lists):
                        if ri < len(lst):
                            ws2.write(row + ri, ci, str(lst[ri]), ctr_fmt)
                for ci in range(len(LIST_HDRS)):
                    ws2.set_column(ci, ci, 28)

        # Sheet 3 — Overall Combined Cycles Raw
        ws3 = wb.add_worksheet("Overall Combined Cycles Raw")
        row = 0
        for ck, sdf in overall_cycles_dict.items():
            if sdf.empty:
                continue
            cyc = _RE_CYCLE_NUM.search(ck).group(1)
            row = write_cycle_block(ws3, ck, sdf.iloc[0],
                                    raw_combined[raw_combined['CYCLE'] == cyc], row, is_combined=True)

    return output.getvalue()

# ── Sidebar ───────────────────────────────────────────────────────────────────

uploaded_files = st.sidebar.file_uploader(
    "Upload Daily Remark Files", type="xlsx", accept_multiple_files=True
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Date Range Filter")
st.sidebar.caption("Only data within this range will be processed.")

_today = datetime.date.today()
_30ago = _today - datetime.timedelta(days=30)

date_filter_on = st.sidebar.toggle("Enable date range filter", value=False)

if date_filter_on:
    c1, c2 = st.sidebar.columns(2)
    start_date = c1.date_input("Start", value=_30ago, key="sd")
    end_date   = c2.date_input("End",   value=_today, key="ed")
    if start_date > end_date:
        st.sidebar.error("⚠️ Start must be ≤ End date.")
        st.stop()
    st.sidebar.info(f"📆 {start_date:%b %d, %Y} → {end_date:%b %d, %Y}")
else:
    start_date = end_date = None
    st.sidebar.caption("_Filter off — all dates processed._")

st.sidebar.markdown("---")

# ── Main processing ───────────────────────────────────────────────────────────

if not uploaded_files:
    st.info("Upload one or more XLSX files to start.")
    st.stop()

if date_filter_on:
    st.info(f"📅 **Date filter active:** {start_date:%B %d, %Y} → {end_date:%B %d, %Y}")

# Collect into lists; single pd.concat per accumulator (avoids O(n²) frame copies)
raw_pred_frames: list[pl.DataFrame] = []
raw_comb_frames: list[pl.DataFrame] = []
all_combined, all_predictive, all_manual = [], [], []
pred_cycles_acc:  dict[str, list] = {}
man_cycles_acc:   dict[str, list] = {}
pred_bals_acc:    dict[str, list] = {}
man_bals_acc:     dict[str, list] = {}

prog = st.progress(0)
n    = len(uploaded_files)

for idx, f in enumerate(uploaded_files):
    with st.spinner(f"Processing {f.name}…"):
        file_bytes = f.getvalue()
        df = load_data(f.name, file_bytes)
        df = process_file(df)

        if date_filter_on:
            df = apply_date_filter(df, start_date, end_date)
            if df.is_empty():
                prog.progress((idx + 1) / n)
                continue

        if 'REMARK TYPE' not in df.columns:
            prog.progress((idx + 1) / n)
            continue

        remark_expr = pl.lit(False)
        if 'REMARK' in df.columns:
            remark_expr = pl.col('REMARK').str.contains('(?i)Predictive')

        follow_up = df.filter(
            (pl.col('REMARK TYPE') == 'Follow Up') & remark_expr
        )
        predictive = df.filter(pl.col('REMARK TYPE') == 'Predictive')
        outgoing = df.filter(pl.col('REMARK TYPE') == 'Outgoing')

        pred_combined = pl.concat([follow_up, predictive], how='diagonal_relaxed')
        combined = pl.concat([pred_combined, outgoing], how='diagonal_relaxed')

        raw_pred_frames.append(pred_combined)
        raw_comb_frames.append(combined)

        all_combined.append(calculate_summary(combined,      ['Predictive', 'Follow Up', 'Outgoing']))
        all_predictive.append(calculate_summary(pred_combined, ['Predictive', 'Follow Up']))
        all_manual.append(calculate_summary(outgoing,        ['Outgoing']))

        for k, v in get_cycle_summary(pred_combined, ['Predictive', 'Follow Up']).items():
            pred_cycles_acc.setdefault(k, []).append(v)
        for k, v in get_cycle_summary(outgoing, ['Outgoing']).items():
            man_cycles_acc.setdefault(k, []).append(v)
        for k, v in get_balance_summary(pred_combined, ['Predictive', 'Follow Up']).items():
            pred_bals_acc.setdefault(k, []).append(v)
        for k, v in get_balance_summary(outgoing, ['Outgoing']).items():
            man_bals_acc.setdefault(k, []).append(v)

    prog.progress((idx + 1) / n)

prog.empty()

# Final single-concat per accumulator
def _merge(acc: dict) -> dict:
    return {k: pd.concat(v, ignore_index=True) for k, v in acc.items()}

def _concat_summaries(frames: list) -> pd.DataFrame:
    valid = [f for f in frames if not f.empty]
    return pd.concat(valid, ignore_index=True).sort_values('DATE') if valid else _empty_summary()

raw_predictive_all_pl = pl.concat(raw_pred_frames, how='diagonal_relaxed') if raw_pred_frames else _empty_pl_df()
raw_combined_all_pl = pl.concat(raw_comb_frames, how='diagonal_relaxed') if raw_comb_frames else _empty_pl_df()

raw_predictive_all = raw_predictive_all_pl.to_pandas() if not raw_predictive_all_pl.is_empty() else pd.DataFrame()
raw_combined_all = raw_combined_all_pl.to_pandas() if not raw_combined_all_pl.is_empty() else pd.DataFrame()

if raw_combined_all.empty:
    st.warning("⚠️ No records found in the selected date range. Adjust the filter or disable it.")
    st.stop()

pred_cycles_merged = _merge(pred_cycles_acc)
man_cycles_merged  = _merge(man_cycles_acc)
pred_bals_merged   = _merge(pred_bals_acc)
man_bals_merged    = _merge(man_bals_acc)

combined_summary   = _concat_summaries(all_combined)
predictive_summary = _concat_summaries(all_predictive)
manual_summary     = _concat_summaries(all_manual)

combined_cycle         = combine_summaries(pred_cycles_merged, man_cycles_merged)
combined_balance       = combine_summaries(pred_bals_merged,   man_bals_merged)
pred_cycles_total      = aggregate_cycles_by_cycle(pred_cycles_merged, raw_predictive_all)
overall_combined_total = aggregate_overall_summary(raw_combined_all, combined_summary)
overall_combined_cycs  = aggregate_overall_combined_cycles(combined_cycle, raw_combined_all)

date_tag = (f" — filtered {start_date:%b %d} – {end_date:%b %d, %Y}" if date_filter_on else "")
st.success(f"✅ Processed {n} file(s){date_tag}")

# ── Display ───────────────────────────────────────────────────────────────────

st.write("## Overall Combined Summary (Daily)")
st.dataframe(combined_summary, use_container_width=True)

if not overall_combined_total.empty:
    st.write("## Overall Combined Summary (Total Period)")
    st.dataframe(overall_combined_total, use_container_width=True)

if overall_combined_cycs:
    st.write("## Overall Combined Cycles (Total per Cycle)")
    for ck in sorted(overall_combined_cycs, key=_cycle_sort_key):
        st.subheader(f"{ck} Summary")
        st.dataframe(overall_combined_cycs[ck], use_container_width=True)

for title, df in [("Overall Predictive", predictive_summary), ("Overall Manual", manual_summary)]:
    if not df.empty:
        st.write(f"## {title} Summary")
        st.dataframe(df, use_container_width=True)

# ── Exports ───────────────────────────────────────────────────────────────────

date_suffix = (f"_{start_date:%Y%m%d}_to_{end_date:%Y%m%d}"
               if date_filter_on else f"_{datetime.datetime.now():%Y%m%d_%H%M%S}")

summary_groups = {
    'Combined':                  {'Combined Summary (Daily)': combined_summary},
    'Overall Combined Summary':  {'Overall Combined Summary (Total)': overall_combined_total},
    'Overall Combined Cycles':   overall_combined_cycs,
    'Predictive':                {'Predictive Summary': predictive_summary},
    'Manual':                    {'Manual Summary': manual_summary},
    'Combined Cycles':           {k: v for k, v in combined_cycle.items()       if 'na' not in k.lower()},
    'Predictive Cycles':         {k: v for k, v in pred_cycles_merged.items()   if 'na' not in k.lower()},
    'Manual Cycles':             {k: v for k, v in man_cycles_merged.items()    if 'na' not in k.lower()},
    'Predictive Cycles (Total)': pred_cycles_total,
    'Combined Balances':         combined_balance,
    'Predictive Balances':       pred_bals_merged,
    'Manual Balances':           man_bals_merged,
}

st.sidebar.download_button(
    label="📥 Download All Summaries",
    data=to_excel(summary_groups),
    file_name=f"Debt_Collection_Summary{date_suffix}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.sidebar.download_button(
    label="📥 Download Raw Cycle Breakdown",
    data=create_raw_cycle_breakdown_excel(
        raw_predictive_all, raw_combined_all,
        pred_cycles_merged, overall_combined_total, overall_combined_cycs,
    ),
    file_name=f"Raw_Cycle_Breakdown{date_suffix}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
