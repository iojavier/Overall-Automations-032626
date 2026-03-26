import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from collections import defaultdict

st.set_page_config(page_title="Hourly Performance Report", layout="wide")
st.title("Call Summary & DRR Predictive Performance")

time_groups = ["7am", "8am", "9am", "10am", "11am", "12pm",
               "1pm", "2pm", "3pm", "4pm", "5pm", "6pm", "7pm", "8pm"]

def get_time_group(hour):
    return {7: "7am", 8: "8am", 9: "9am", 10: "10am", 11: "11am", 12: "12pm",
            13: "1pm", 14: "2pm", 15: "3pm", 16: "4pm", 17: "5pm",
            18: "6pm", 19: "7pm", 20: "8pm"}.get(hour)

def find_column(df, names):
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for name in names:
        if name.strip().lower() in cols_lower:
            return cols_lower[name.strip().lower()]
    return None

def parse_drr_datetime(row):
    date_str = str(row.get("Date", "")).strip()
    time_str = str(row.get("Time", "")).strip()
    if not date_str or date_str in ["nan", ""]: 
        return pd.NaT
    full_str = f"{date_str} {time_str or '12:00AM'}"
    formats = [
        '%d/%m/%Y %I:%M:%S %p', '%d/%m/%Y %I:%M %p', '%d/%m/%Y %H:%M:%S',
        '%d-%m-%Y %I:%M:%S %p', '%d-%m-%Y %H:%M:%S', '%d/%m/%Y'
    ]
    for fmt in formats:
        try:
            return pd.to_datetime(full_str, format=fmt, dayfirst=True)
        except:
            continue
    return pd.to_datetime(full_str, dayfirst=True, errors='coerce')

# ====================== CORRECTED STYLE FUNCTION ======================
def style_excel(blocks):
    """Styles the Excel sheet and saves the blocks."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        # Remove default sheet if it exists, before creating the new one
        if "Sheet" in workbook.sheetnames:
            workbook.remove(workbook["Sheet"])
        ws = workbook.create_sheet("Summary Report")

        # Define styles
        border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
        bold = Font(bold=True)
        center = Alignment(horizontal='center', vertical='center')
        red_fill = PatternFill(start_color="963634", end_color="963634", fill_type="solid")
        white_font = Font(bold=True, color="FFFFFF")
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        row = 1
        for block in blocks:
            for r_idx, row_data in enumerate(block):
                current_row = row + r_idx
                col_count = len(row_data) 
                
                # Check for Red Headers (Merge and Style)
                if r_idx == 0 and str(row_data[0]).startswith(("TOTAL ", "--- DRR", "DRR Predictive")):
                    header_text = row_data[0]
                    
                    # 1. Write the value to the top-left cell (A-column)
                    cell = ws.cell(current_row, 1, header_text)
                    
                    # 2. Apply styling
                    cell.fill = red_fill
                    cell.font = white_font
                    cell.alignment = center
                    
                    # 3. Merge the cells
                    ws.merge_cells(start_row=current_row, start_column=1,
                                   end_row=current_row, end_column=col_count)

                    # 4. Apply border to the entire merged range
                    for c_idx in range(1, col_count + 1):
                         # Note: Use ws.cell() without specifying value here
                         ws.cell(current_row, c_idx).border = border

                else:
                    # Normal data rows
                    for c_idx, value in enumerate(row_data, 1):
                        # This writes the value and retrieves the cell object
                        cell = ws.cell(current_row, c_idx, value) 
                        cell.border = border
                        cell.alignment = center

                        # Time group headers (gray)
                        if len(row_data) > 1 and c_idx > 1 and row_data[1] in time_groups:
                            cell.fill = gray_fill
                            cell.font = bold
                        
                        # First column bold
                        elif c_idx == 1:
                            cell.font = bold

            row += len(block) + 2  # Extra spacing

        # Auto-adjust column widths
        for col in range(1, len(time_groups) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 13

    output.seek(0)
    return output
# ====================== END OF CORRECTED STYLE FUNCTION ======================


# ====================== SIDEBAR ======================
st.sidebar.markdown("### 1. Call History Files (xlsx/csv)")
call_files = st.sidebar.file_uploader("Upload Call History", type=["xlsx", "csv"], accept_multiple_files=True, key="call")

st.sidebar.markdown("### 2. DRR Files (xlsx/csv)")
drr_files = st.sidebar.file_uploader("Upload DRR Files", type=["xlsx", "csv"], accept_multiple_files=True, key="drr")

call_infos = []
drr_by_date_hour = defaultdict(lambda: defaultdict(lambda: {
    "agents": set(), "dials": 0, "connected": 0, "rpc": 0, "ptp": 0
}))

# ====================== PROCESS CALL HISTORY ======================
if call_files:
    date_totals = defaultdict(lambda: {tg: {"agents": set(), "dials": 0, "connected": 0} for tg in time_groups})

    for file in call_files:
        try:
            df = pd.read_excel(file) if file.name.endswith(".xlsx") else pd.read_csv(file)
            date_col = find_column(df, ["Call Date", "Date", "call_date"])
            agent_col = find_column(df, ["Collector Name", "Agent", "Collector", "collector_name"])
            status_col = find_column(df, ["Final Dial Status", "Status", "final_status"])

            if not date_col:
                st.warning(f"⚠️ Missing date column in {file.name}")
                continue

            df = df.rename(columns={date_col: "Date", agent_col: "Agent", status_col: "Status"})
            df["datetime"] = pd.to_datetime(df["Date"], dayfirst=True, errors='coerce')
            df = df.dropna(subset=["datetime"])
            df["hour"] = df["datetime"].dt.hour
            df["time_group"] = df["hour"].apply(get_time_group)
            df = df[df["time_group"].notna()]
            if df.empty:
                continue

            date_str = df["datetime"].dt.strftime("%Y-%m-%d").iloc[0]

            summary = {tg: {"agents": 0, "dials": 0, "connected": 0} for tg in time_groups}
            agents_set = {tg: set() for tg in time_groups}

            for tg in time_groups:
                sub = df[df["time_group"] == tg]
                if sub.empty:
                    continue
                agents = sub["Agent"].dropna().astype(str).unique()
                agents_set[tg].update(agents)
                summary[tg]["agents"] = len(agents)
                summary[tg]["dials"] = len(sub)
                if "Status" in df.columns:
                    summary[tg]["connected"] = sub["Status"].astype(str).str.contains("transferred|dropped", case=False, na=False).sum()

            block = [
                [f"{date_str} Call History"] + time_groups,
                ["Active agents"] + [summary[tg]["agents"] for tg in time_groups],
                ["Accounts"] + [""] * len(time_groups),
                ["Dials"] + [summary[tg]["dials"] for tg in time_groups],
                ["Connected"] + [summary[tg]["connected"] for tg in time_groups]
            ]
            call_infos.append({"date": date_str, "block": block, "agents": agents_set, "summary": summary})

            # Accumulate totals per date
            for tg in time_groups:
                date_totals[date_str][tg]["agents"].update(agents_set[tg])
                date_totals[date_str][tg]["dials"] += summary[tg]["dials"]
                date_totals[date_str][tg]["connected"] += summary[tg]["connected"]

        except Exception as e:
            st.error(f"Call History error in {file.name}: {e}")

# ====================== PROCESS DRR FILES ======================
if drr_files:
    for file in drr_files:
        try:
            df = pd.read_excel(file) if file.name.endswith(".xlsx") else pd.read_csv(file)
            if df.empty:
                continue

            cols = {
                "remark_type": find_column(df, ["Remark Type", "remark_type"]),
                "remark": find_column(df, ["Remark", "remark"]),
                "remark_by": find_column(df, ["Remark By"]),
                "debtor_id": find_column(df, ["Debtor ID"]),
                "talk_time": find_column(df, ["Talk Time Duration"]),
                "date": find_column(df, ["Date"]),
                "time": find_column(df, ["Time"]),
                "status": find_column(df, ["Status"]),
                "ptp": find_column(df, ["PTP Amount"])
            }

            missing = [k for k, v in cols.items() if v is None]
            if missing:
                st.error(f"Missing columns in {file.name}: {missing}")
                continue

            df = df.rename(columns=cols)
            df["datetime"] = df.apply(parse_drr_datetime, axis=1)
            df = df.dropna(subset=["datetime"])
            df["date_str"] = df["datetime"].dt.strftime("%Y-%m-%d")
            df["hour"] = df["datetime"].dt.hour
            df["time_group"] = df["hour"].apply(get_time_group)
            df = df[df["time_group"].notna()]
            if df.empty:
                continue

            df["Remark Type"] = df["Remark Type"].astype(str).str.strip()
            df["Remark"] = df["Remark"].astype(str).str.lower()
            df["Talk Time Duration"] = pd.to_numeric(df["Talk Time Duration"], errors='coerce').fillna(0)
            df["PTP Amount"] = pd.to_numeric(df["PTP Amount"], errors='coerce').fillna(0)

            # Predictive dials
            predictive_mask = (
                (df["Remark Type"] == "Predictive") |
                ((df["Remark Type"] == "Follow Up") & df["Remark"].str.contains("predictive cp", na=False))
            )

            pred_df = df[predictive_mask]

            for (date_str, tg), group in pred_df.groupby(["date_str", "time_group"]):
                agents = group["Remark By"].dropna().astype(str).unique()
                drr_by_date_hour[date_str][tg]["agents"].update(agents)
                drr_by_date_hour[date_str][tg]["dials"] += len(group)
                drr_by_date_hour[date_str][tg]["connected"] += group[group["Talk Time Duration"] > 0]["Debtor ID"].nunique()

                s = group["Status"].astype(str).str.lower()
                r = group["Remark"].astype(str).str.lower()
                rpc_mask = (
                    s.str.contains("bank escalation|rpc|ptp", na=False) |
                    ((s == "negative callouts - call drop") & r.str.contains("confirmed", na=False)) |
                    ((s == "negative callouts - busy_ooca") & r.str.contains("confirmed", na=False))
                )
                drr_by_date_hour[date_str][tg]["rpc"] += rpc_mask.sum()
                drr_by_date_hour[date_str][tg]["ptp"] += (group["PTP Amount"] > 0).sum()

        except Exception as e:
            st.error(f"DRR processing error in {file.name}: {e}")

# ====================== BUILD FINAL BLOCKS ======================
final_blocks = []

# Call History Blocks + Totals
seen_dates = set()
for info in call_infos:
    date_str = info["date"]
    if date_str not in seen_dates:
        # Add total for previous date if exists
        if seen_dates:
            prev_date = list(seen_dates)[-1]
            tot = date_totals[prev_date]
            final_blocks.append([
                [f"TOTAL {prev_date} (Call History)"] + time_groups,
                ["Active agents"] + [len(tot[tg]["agents"]) for tg in time_groups],
                ["Accounts"] + [""] * len(time_groups),
                ["Dials"] + [tot[tg]["dials"] for tg in time_groups],
                ["Connected"] + [tot[tg]["connected"] for tg in time_groups]
            ])
        seen_dates.add(date_str)

    final_blocks.append(info["block"])

# Add final total if any call files
if seen_dates:
    last_date = list(seen_dates)[-1]
    tot = date_totals[last_date]
    final_blocks.append([
        [f"TOTAL {last_date} (Call History)"] + time_groups,
        ["Active agents"] + [len(tot[tg]["agents"]) for tg in time_groups],
        ["Accounts"] + [""] * len(time_groups),
        ["Dials"] + [tot[tg]["dials"] for tg in time_groups],
        ["Connected"] + [tot[tg]["connected"] for tg in time_groups]
    ])

# DRR Blocks
if drr_by_date_hour:
    if final_blocks:
        final_blocks.append([["--- DRR PREDICTIVE PERFORMANCE ---"] + [""] * len(time_groups)])

    for date_str in sorted(drr_by_date_hour.keys()):
        data = drr_by_date_hour[date_str]
        block = [
            [f"DRR Predictive {date_str}"] + time_groups,
            ["", *time_groups],  # Proper time header
            ["Agents"] + [len(data[tg]["agents"]) for tg in time_groups],
            ["Accounts"] + [""] * len(time_groups),
            ["Dials"] + [data[tg]["dials"] for tg in time_groups],
            ["Connected"] + [data[tg]["connected"] for tg in time_groups],
            ["RPC"] + [data[tg]["rpc"] for tg in time_groups],
            ["PTP"] + [data[tg]["ptp"] for tg in time_groups]
        ]
        final_blocks.append(block)

# ====================== DISPLAY ======================
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

    st.download_button(
        label="📥 Download Excel Report",
        data=style_excel(final_blocks),
        file_name=f"Hourly_Performance_Report_{pd.Timestamp('today').strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Please upload Call History and/or DRR files to generate the report.")