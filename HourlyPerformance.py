import pandas as pd
import polars as pl
import streamlit as st
from io import BytesIO, StringIO


st.set_page_config(page_title="Hourly Performance Report", layout="wide")
st.title("Call Summary & DRR Predictive Performance")

time_groups = [
    "7am", "8am", "9am", "10am", "11am", "12pm",
    "1pm", "2pm", "3pm", "4pm", "5pm", "6pm", "7pm", "8pm",
]

TIME_GROUP_MAP = {
    7: "7am", 8: "8am", 9: "9am", 10: "10am", 11: "11am", 12: "12pm",
    13: "1pm", 14: "2pm", 15: "3pm", 16: "4pm", 17: "5pm", 18: "6pm",
    19: "7pm", 20: "8pm",
}

DRR_FORMATS = [
    "%d/%m/%Y %I:%M:%S %p",
    "%d/%m/%Y %I:%M %p",
    "%d/%m/%Y %H:%M:%S",
    "%d-%m-%Y %I:%M:%S %p",
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y",
]


def get_time_group(hour):
    return TIME_GROUP_MAP.get(hour)


def find_column(df, names):
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for name in names:
        if name.strip().lower() in cols_lower:
            return cols_lower[name.strip().lower()]
    return None


@st.cache_data(show_spinner=False)
def load_uploaded_dataframe(file_name: str, file_bytes: bytes) -> pd.DataFrame:
    try:
        if file_name.lower().endswith(".xlsx"):
            return pl.read_excel(BytesIO(file_bytes), engine="calamine").to_pandas()

        return pl.read_csv(
            BytesIO(file_bytes),
            infer_schema_length=10000,
            ignore_errors=True,
            try_parse_dates=False,
        ).to_pandas()
    except Exception:
        if file_name.lower().endswith(".xlsx"):
            return pd.read_excel(BytesIO(file_bytes))

        for encoding in ("utf-8", "utf-8-sig", "latin1"):
            try:
                return pd.read_csv(StringIO(file_bytes.decode(encoding)), low_memory=False)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(BytesIO(file_bytes), low_memory=False)


def parse_drr_datetime_columns(df):
    date_series = df["Date"]
    time_series = df["Time"]
    parsed = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")

    date_direct = pd.to_datetime(date_series, dayfirst=True, errors="coerce")
    parsed.loc[date_direct.notna()] = date_direct.loc[date_direct.notna()]

    date_numeric = pd.to_numeric(date_series, errors="coerce")
    date_excel = pd.to_datetime(date_numeric, unit="D", origin="1899-12-30", errors="coerce")
    parsed.loc[parsed.isna() & date_excel.notna()] = date_excel.loc[parsed.isna() & date_excel.notna()]

    time_direct = pd.to_datetime(time_series, errors="coerce")
    time_numeric = pd.to_numeric(time_series, errors="coerce")
    time_excel = pd.to_datetime(time_numeric, unit="D", origin="1899-12-30", errors="coerce")

    has_time_direct = time_direct.notna()
    parsed.loc[has_time_direct] = parsed.loc[has_time_direct].fillna(pd.Timestamp("1899-12-30")) + (
        time_direct.loc[has_time_direct] - time_direct.loc[has_time_direct].dt.normalize()
    )

    has_time_excel = ~has_time_direct & time_excel.notna()
    parsed.loc[has_time_excel] = parsed.loc[has_time_excel].fillna(pd.Timestamp("1899-12-30")) + (
        time_excel.loc[has_time_excel] - time_excel.loc[has_time_excel].dt.normalize()
    )

    remaining = parsed.isna()
    if remaining.any():
        date_text = date_series.fillna("").astype(str).str.strip()
        time_text = time_series.fillna("").astype(str).str.strip()
        combined = date_text + " " + time_text.where(time_text.ne(""), "12:00AM")
        valid = remaining & date_text.ne("") & date_text.ne("nan")
        combined_valid = combined[valid]

        for fmt in DRR_FORMATS:
            still_missing = parsed[valid].isna()
            if not still_missing.any():
                break
            idx = combined_valid.index[still_missing]
            parsed.loc[idx] = pd.to_datetime(combined_valid.loc[idx], format=fmt, dayfirst=True, errors="coerce")

        still_missing = parsed[valid].isna()
        if still_missing.any():
            idx = combined_valid.index[still_missing]
            parsed.loc[idx] = pd.to_datetime(combined_valid.loc[idx], dayfirst=True, errors="coerce")

    return parsed


def build_time_metric_row(label, metrics_by_group, metric_name):
    return [label] + [int(metrics_by_group[tg].get(metric_name, 0)) for tg in time_groups]


def style_excel(blocks):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Summary Report")

        title_fmt = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#963634",
            "font_color": "white",
        })
        header_fmt = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#D3D3D3",
        })
        label_fmt = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        cell_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })

        row = 0
        max_cols = len(time_groups) + 1
        for block in blocks:
            for r_idx, row_data in enumerate(block):
                col_count = len(row_data)
                if r_idx == 0 and str(row_data[0]).startswith(("TOTAL ", "--- DRR", "DRR Predictive")):
                    worksheet.merge_range(row, 0, row, col_count - 1, row_data[0], title_fmt)
                else:
                    is_time_header = len(row_data) > 1 and row_data[0] == "" and row_data[1] in time_groups
                    for c_idx, value in enumerate(row_data):
                        fmt = header_fmt if is_time_header else (label_fmt if c_idx == 0 else cell_fmt)
                        worksheet.write(row, c_idx, value, fmt)
                row += 1
            row += 2

        worksheet.set_column(0, max_cols - 1, 13)

    output.seek(0)
    return output.getvalue()


def blocks_to_dataframe(blocks):
    rows = []
    width = len(time_groups) + 1
    columns = ["Label"] + time_groups

    for block in blocks:
        for row in block:
            padded = list(row) + [""] * (width - len(row))
            rows.append(padded[:width])
        rows.append([""] * width)

    return pd.DataFrame(rows, columns=columns)


@st.cache_data(show_spinner=False)
def build_csv_bytes(blocks):
    df = blocks_to_dataframe(blocks)
    return df.to_csv(index=False).encode("utf-8-sig")


st.sidebar.markdown("### 1. Call History Files (xlsx/csv)")
call_files = st.sidebar.file_uploader("Upload Call History", type=["xlsx", "csv"], accept_multiple_files=True, key="call")

st.sidebar.markdown("### 2. DRR Files (xlsx/csv)")
drr_files = st.sidebar.file_uploader("Upload DRR Files", type=["xlsx", "csv"], accept_multiple_files=True, key="drr")

call_infos = []
date_totals = {}
drr_by_date_hour = {}
processed_call_files = 0
processed_drr_files = 0
drr_debug_messages = []


if call_files:
    for file in call_files:
        try:
            df = load_uploaded_dataframe(file.name, file.getvalue())
            date_col = find_column(df, ["Call Date", "Date", "call_date"])
            agent_col = find_column(df, ["Collector Name", "Agent", "Collector", "collector_name"])
            status_col = find_column(df, ["Final Dial Status", "Status", "final_status"])

            if not date_col:
                st.warning(f"Missing date column in {file.name}")
                continue

            rename_map = {date_col: "Date"}
            if agent_col:
                rename_map[agent_col] = "Agent"
            if status_col:
                rename_map[status_col] = "Status"
            df = df.rename(columns=rename_map)

            df["datetime"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
            df = df.dropna(subset=["datetime"])
            if df.empty:
                continue

            df["time_group"] = df["datetime"].dt.hour.map(TIME_GROUP_MAP)
            df = df[df["time_group"].notna()].copy()
            if df.empty:
                continue
            processed_call_files += 1

            date_str = df["datetime"].dt.strftime("%Y-%m-%d").iloc[0]
            grouped = df.groupby("time_group", sort=False)

            dials = grouped.size().reindex(time_groups, fill_value=0)
            if "Agent" in df.columns:
                agents = grouped["Agent"].nunique(dropna=True).reindex(time_groups, fill_value=0)
                agent_sets = grouped["Agent"].agg(
                    lambda s: set(s.dropna().astype(str).unique())
                ).reindex(time_groups, fill_value=set())
            else:
                agents = pd.Series(0, index=time_groups)
                agent_sets = pd.Series([set() for _ in time_groups], index=time_groups)

            if "Status" in df.columns:
                connected_mask = df["Status"].astype(str).str.contains("transferred|dropped", case=False, na=False)
                connected = (
                    df.assign(connected=connected_mask.astype("int8"))
                    .groupby("time_group")["connected"]
                    .sum()
                    .reindex(time_groups, fill_value=0)
                )
            else:
                connected = pd.Series(0, index=time_groups)

            metrics_by_group = {
                tg: {
                    "agents": int(agents.loc[tg]),
                    "dials": int(dials.loc[tg]),
                    "connected": int(connected.loc[tg]),
                }
                for tg in time_groups
            }

            block = [
                [f"{date_str} Call History"] + time_groups,
                build_time_metric_row("Active agents", metrics_by_group, "agents"),
                ["Accounts"] + [""] * len(time_groups),
                build_time_metric_row("Dials", metrics_by_group, "dials"),
                build_time_metric_row("Connected", metrics_by_group, "connected"),
            ]
            call_infos.append({"date": date_str, "block": block})

            total_bucket = date_totals.setdefault(
                date_str,
                {tg: {"agents": set(), "dials": 0, "connected": 0} for tg in time_groups},
            )
            for tg in time_groups:
                total_bucket[tg]["agents"].update(agent_sets.loc[tg] if tg in agent_sets.index else set())
                total_bucket[tg]["dials"] += metrics_by_group[tg]["dials"]
                total_bucket[tg]["connected"] += metrics_by_group[tg]["connected"]

        except Exception as e:
            st.error(f"Call History error in {file.name}: {e}")


if drr_files:
    for file in drr_files:
        try:
            df = load_uploaded_dataframe(file.name, file.getvalue())
            if df.empty:
                drr_debug_messages.append(f"{file.name}: file loaded but has no rows.")
                continue

            cols = {
                "Remark Type": find_column(df, ["Remark Type", "remark_type"]),
                "Remark": find_column(df, ["Remark", "remark"]),
                "Remark By": find_column(df, ["Remark By"]),
                "Debtor ID": find_column(df, ["Debtor ID"]),
                "Talk Time Duration": find_column(df, ["Talk Time Duration"]),
                "Date": find_column(df, ["Date"]),
                "Time": find_column(df, ["Time"]),
                "Status": find_column(df, ["Status"]),
                "PTP Amount": find_column(df, ["PTP Amount"]),
            }

            missing = [k for k, v in cols.items() if v is None]
            if missing:
                st.error(f"Missing columns in {file.name}: {missing}")
                continue

            df = df.rename(columns=cols)
            df["datetime"] = parse_drr_datetime_columns(df)
            df = df.dropna(subset=["datetime"])
            if df.empty:
                drr_debug_messages.append(f"{file.name}: no valid Date/Time rows were parsed.")
                continue

            df["time_group"] = df["datetime"].dt.hour.map(TIME_GROUP_MAP)
            df = df[df["time_group"].notna()].copy()
            if df.empty:
                drr_debug_messages.append(f"{file.name}: rows were found, but none are within the supported 7am-8pm hourly groups.")
                continue

            df["date_str"] = df["datetime"].dt.strftime("%Y-%m-%d")
            df["Remark Type"] = df["Remark Type"].astype(str).str.strip().str.lower()
            df["Remark"] = df["Remark"].astype(str).str.lower()
            df["Status"] = df["Status"].astype(str).str.lower()
            df["Talk Time Duration"] = pd.to_numeric(df["Talk Time Duration"], errors="coerce").fillna(0)
            df["PTP Amount"] = pd.to_numeric(df["PTP Amount"], errors="coerce").fillna(0)

            predictive_mask = (
                df["Remark Type"].eq("predictive") |
                (df["Remark Type"].eq("follow up") & df["Remark"].str.contains("predictive", na=False))
            )
            pred_df = df[predictive_mask].copy()
            if pred_df.empty:
                drr_debug_messages.append(f"{file.name}: no predictive rows matched after Remark Type filtering.")
                continue

            rpc_mask = (
                pred_df["Status"].str.contains("bank escalation|rpc|ptp", na=False) |
                ((pred_df["Status"] == "negative callouts - call drop") & pred_df["Remark"].str.contains("confirmed", na=False)) |
                ((pred_df["Status"] == "negative callouts - busy_ooca") & pred_df["Remark"].str.contains("confirmed", na=False))
            )
            pred_df["rpc_flag"] = rpc_mask.astype("int8")
            pred_df["ptp_flag"] = (pred_df["PTP Amount"] > 0).astype("int8")

            base_grouped = pred_df.groupby(["date_str", "time_group"], sort=False).agg(
                agents=("Remark By", lambda s: s.dropna().astype(str).nunique()),
                dials=("date_str", "size"),
                rpc=("rpc_flag", "sum"),
                ptp=("ptp_flag", "sum"),
            )

            connected_grouped = (
                pred_df.loc[pred_df["Talk Time Duration"] > 0]
                .groupby(["date_str", "time_group"])["Debtor ID"]
                .nunique()
                .rename("connected")
            )

            merged = base_grouped.join(connected_grouped, how="left").fillna({"connected": 0}).reset_index()
            if merged.empty:
                drr_debug_messages.append(f"{file.name}: predictive rows were found, but grouping produced no hourly output.")
                continue

            for row in merged.itertuples(index=False):
                bucket = drr_by_date_hour.setdefault(
                    row.date_str,
                    {tg: {"agents": 0, "dials": 0, "connected": 0, "rpc": 0, "ptp": 0} for tg in time_groups},
                )
                bucket[row.time_group]["agents"] += int(row.agents)
                bucket[row.time_group]["dials"] += int(row.dials)
                bucket[row.time_group]["connected"] += int(row.connected)
                bucket[row.time_group]["rpc"] += int(row.rpc)
                bucket[row.time_group]["ptp"] += int(row.ptp)
            processed_drr_files += 1

        except Exception as e:
            st.error(f"DRR processing error in {file.name}: {e}")


final_blocks = []

if call_infos:
    current_date = None
    for info in call_infos:
        date_str = info["date"]
        if current_date is not None and date_str != current_date:
            totals = date_totals[current_date]
            final_blocks.append([
                [f"TOTAL {current_date} (Call History)"] + time_groups,
                ["Active agents"] + [len(totals[tg]["agents"]) for tg in time_groups],
                ["Accounts"] + [""] * len(time_groups),
                ["Dials"] + [totals[tg]["dials"] for tg in time_groups],
                ["Connected"] + [totals[tg]["connected"] for tg in time_groups],
            ])

        final_blocks.append(info["block"])
        current_date = date_str

    if current_date is not None:
        totals = date_totals[current_date]
        final_blocks.append([
            [f"TOTAL {current_date} (Call History)"] + time_groups,
            ["Active agents"] + [len(totals[tg]["agents"]) for tg in time_groups],
            ["Accounts"] + [""] * len(time_groups),
            ["Dials"] + [totals[tg]["dials"] for tg in time_groups],
            ["Connected"] + [totals[tg]["connected"] for tg in time_groups],
        ])


if drr_by_date_hour:
    if final_blocks:
        final_blocks.append([["--- DRR PREDICTIVE PERFORMANCE ---"] + [""] * len(time_groups)])

    for date_str in sorted(drr_by_date_hour.keys()):
        data = drr_by_date_hour[date_str]
        block = [
            [f"DRR Predictive {date_str}"] + time_groups,
            [""] + time_groups,
            build_time_metric_row("Agents", data, "agents"),
            ["Accounts"] + [""] * len(time_groups),
            build_time_metric_row("Dials", data, "dials"),
            build_time_metric_row("Connected", data, "connected"),
            build_time_metric_row("RPC", data, "rpc"),
            build_time_metric_row("PTP", data, "ptp"),
        ]
        final_blocks.append(block)


if processed_call_files or processed_drr_files:
    st.caption(f"Processed call files: {processed_call_files} | Processed DRR files: {processed_drr_files}")

if final_blocks:
    html = """
    <style>
        table {width:100%; border-collapse:collapse; font-family:Arial, sans-serif; margin:25px 0; font-size:14px;}
        td, th {border:1px solid #963634; padding:10px; text-align:center;}
        .red-header {background:#963634; color:white; font-weight:bold; font-size:17px;}
        .gray-header {background:#D3D3D3; font-weight:bold;}
        .bold {font-weight:bold;}
        tr.spacer td {height:20px; border:none; background:none;}
    </style>
    <table>
    """
    for block in final_blocks:
        for i, row in enumerate(block):
            if i == 0 and any(h in str(row[0]) for h in ["TOTAL", "--- DRR", "DRR Predictive"]):
                html += f"<tr><td colspan='{len(row)}' class='red-header'>{row[0]}</td></tr>"
            elif row[0] == "" and len(row) > 1 and row[1] in time_groups:
                html += "<tr class='gray-header'>"
                for cell in row:
                    html += f"<td>{cell}</td>"
                html += "</tr>"
            else:
                html += "<tr>"
                for j, cell in enumerate(row):
                    if j == 0 and i > 0:
                        html += f"<td class='bold'>{cell}</td>"
                    else:
                        html += f"<td>{cell}</td>"
                html += "</tr>"
        html += "<tr class='spacer'><td colspan='15'></td></tr>"

    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

    report_date = pd.Timestamp("today").strftime("%Y%m%d")
    excel_bytes = style_excel(final_blocks)
    csv_bytes = build_csv_bytes(final_blocks)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Download Excel Report",
            data=excel_bytes,
            file_name=f"Hourly_Performance_Report_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            label="Download CSV Report (Fast)",
            data=csv_bytes,
            file_name=f"Hourly_Performance_Report_{report_date}.csv",
            mime="text/csv",
            use_container_width=True,
        )
else:
    if drr_debug_messages:
        st.warning("Files were uploaded, but no report rows were produced.")
        for msg in drr_debug_messages:
            st.write(f"- {msg}")
    else:
        st.info("Please upload Call History and/or DRR files to generate the report.")
