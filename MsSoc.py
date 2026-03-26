import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import time

# Streamlit page configuration
st.set_page_config(page_title="Debt Collection Summary", layout="wide")

# Sidebar for file uploader and date filter
st.sidebar.header("Upload XLSX Files")
uploaded_files = st.sidebar.file_uploader(
    "Choose XLSX files", type=["xlsx"], accept_multiple_files=True
)
start_date = st.sidebar.date_input("Start Date", value=date(2025, 8, 1))
end_date = st.sidebar.date_input("End Date", value=date(2025, 8, 31))

def process_files(files, start_date, end_date):
    if not files:
        st.warning("Please upload at least one XLSX file.")
        return None, None

    start_time = time.time()
    required_columns = [
        'Date', 'Debtor ID', 'Remark Type', 'Remark', 'Talk Time Duration',
        'Dialed Number', 'Status', 'Claim Paid Amount', 'PTP Amount', 'Client', 'Remark By'
    ]
    progress_bar = st.progress(0)
    progress_text = st.empty()

    try:
        # Load files with chunking and selective columns for speed
        st.write("Loading files...")
        load_start = time.time()
        dfs = []
        for i, file in enumerate(files):
            # Try chunking; fallback to full load if not supported
            try:
                chunk_iter = pd.read_excel(file, usecols=required_columns, chunksize=10000)
                chunks = [chunk for chunk in chunk_iter]
                dfs.append(pd.concat(chunks, ignore_index=True))
            except TypeError:
                df = pd.read_excel(file, usecols=required_columns)
                dfs.append(df)
            progress_bar.progress((20 + (20 * (i + 1) / len(files))) / 100)
        data = pd.concat(dfs, ignore_index=True)
        load_end = time.time()
        progress_text.text("20% - Files loaded")
        st.write(f"File loading time: {load_end - load_start:.2f} seconds")

        # Pre-filter by date for speed (fix: coerce sidebar dates to datetime)
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce')
        start_dt = pd.to_datetime(start_date)  # Coerce to datetime64[ns]
        end_dt = pd.to_datetime(end_date)      # Coerce to datetime64[ns]
        data = data[(data['Date'] >= start_dt) & (data['Date'] <= end_dt)]
        progress_bar.progress(40 / 100)
        progress_text.text("40% - Data filtered")

        # Validate columns
        if not all(col in data.columns for col in required_columns):
            st.error(f"Missing required columns. Found: {data.columns.tolist()}")
            return None, None

        # -------------------- SUMMARY PER CLIENT --------------------
        combine_start = time.time()
        summary_list_client = []
        for client, client_df in data.groupby("Client"):
            # FILTERED DATASET (Follow Up + Predictive + Predictive CP)
            condition_follow_up = client_df['Remark Type'].str.contains('Follow Up', case=False, na=False)
            condition_predictive = client_df['Remark Type'].str.contains('Predictive', case=False, na=False)
            condition_predictive_cp = client_df['Remark'].str.contains('Predictive CP', case=False, na=False)
            filtered_data = client_df[condition_follow_up | condition_predictive | condition_predictive_cp].copy()
            if filtered_data.empty:
                continue

            # DATE RANGE
            min_date = filtered_data['Date'].min()
            max_date = filtered_data['Date'].max()
            date_range = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}".replace(' 0', ' ')

            # ACCOUNTS
            accounts = filtered_data['Debtor ID'].nunique()

            # DIALS (NEW LOGIC)
            dials_follow_up = client_df[(client_df['Remark Type'].str.contains('Follow Up', case=False, na=False)) & 
                                       (client_df['Remark'].str.contains('Predictive CP', case=False, na=False))]
            dials_predictive = client_df[client_df['Remark Type'].str.contains('Predictive', case=False, na=False)]
            dials_df = pd.concat([dials_follow_up, dials_predictive])
            dials = len(dials_df)

            # Connected (Unique Number/Account)
            connected_numbers = filtered_data[filtered_data['Talk Time Duration'] > 0]['Dialed Number'].nunique()
            connected_accounts = filtered_data[filtered_data['Talk Time Duration'] > 0]['Debtor ID'].nunique()

            # Rates
            penetration = dials / accounts if accounts > 0 else 0
            connected_rate = connected_numbers / accounts if accounts > 0 else 0
            contact_rate = connected_accounts / accounts if accounts > 0 else 0

            # DROPPED (original logic for CLIENT - unchanged)
            dropped_df = filtered_data[filtered_data['Status'].str.contains('DROPPED|Dropped', case=False, na=False)]
            connected_debtors = filtered_data[filtered_data['Talk Time Duration'] > 0]['Debtor ID'].unique()
            dropped = dropped_df[~dropped_df['Debtor ID'].isin(connected_debtors)]['Debtor ID'].nunique()
            dropped_rate = dropped / connected_numbers if connected_numbers > 0 else 0

            # KEPT, PTP, RPC
            kept_df = filtered_data[(filtered_data['Claim Paid Amount'] > 0) & 
                                    (filtered_data['Status'].str.contains('PAYMENT', case=False, na=False))]
            ptp_df = filtered_data[(filtered_data['Status'].str.contains('PTP', case=False, na=False)) & 
                                   (~filtered_data['Status'].str.contains('FF|FOLLOW UP|PTP_FFUP', case=False, na=False)) & 
                                   (filtered_data['PTP Amount'] > 0)]
            rpc_df = filtered_data[filtered_data['Status'].str.contains('RPC -|BANK ESCALATION -', case=False, na=False)]
            rpc_dispo = rpc_df['Debtor ID'].nunique()
            ptp = ptp_df['Debtor ID'].nunique()
            kept = kept_df['Debtor ID'].nunique()
            rpc_pipeline = rpc_dispo + ptp + kept

            # Rates
            rpc_rate = rpc_pipeline / accounts if accounts > 0 else 0
            ptp_rate = (ptp + kept) / rpc_pipeline if rpc_pipeline > 0 else 0
            kept_rate = kept / (ptp + kept) if (ptp + kept) > 0 else 0

            summary_list_client.append({
                'Client': client, 'Date Range': date_range, 'Accounts': accounts, 'Dials': dials, 'Penetration': penetration,
                'Connected (Unique Number)': connected_numbers, 'Connected (Unique Account)': connected_accounts,
                'Connected Rate': connected_rate, 'Contact Rate': contact_rate, 'Dropped': dropped, 'Drop Rate': dropped_rate,
                'RPC (Pipeline)': rpc_pipeline, 'RPC (Dispo only)': rpc_dispo, 'PTP': ptp, 'KEPT': kept,
                'RPC Rate': rpc_rate, 'PTP Rate': ptp_rate, 'KEPT Rate': kept_rate, 'Summary Type': 'Client'
            })
        combine_end = time.time()
        progress_bar.progress(60 / 100)
        progress_text.text("60% - Client summary complete")

        # -------------------- SUMMARY PER AGENT --------------------
        summary_list_agent = []
        for client, client_df in data.groupby("Client"):
            for agent, agent_df in client_df.groupby("Remark By"):
                # FILTERED DATASET (Follow Up + Predictive + Predictive CP + Outgoing)
                condition_follow_up = agent_df['Remark Type'].str.contains('Follow Up', case=False, na=False)
                condition_predictive = agent_df['Remark Type'].str.contains('Predictive', case=False, na=False)
                condition_predictive_cp = agent_df['Remark'].str.contains('Predictive CP', case=False, na=False)
                condition_outgoing = agent_df['Remark Type'].str.contains('Outgoing', case=False, na=False)
                filtered_data = agent_df[condition_follow_up | condition_predictive | condition_predictive_cp | condition_outgoing].copy()
                if filtered_data.empty:
                    continue

                # DATE RANGE
                min_date = filtered_data['Date'].min()
                max_date = filtered_data['Date'].max()
                date_range = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}".replace(' 0', ' ')

                # ACCOUNTS
                accounts = filtered_data['Debtor ID'].nunique()

                # DIALS (NEW LOGIC)
                dials_follow_up = agent_df[(agent_df['Remark Type'].str.contains('Follow Up', case=False, na=False)) & 
                                          (agent_df['Remark'].str.contains('Predictive CP', case=False, na=False))]
                dials_predictive = agent_df[agent_df['Remark Type'].str.contains('Predictive', case=False, na=False)]
                dials_outgoing = agent_df[agent_df['Remark Type'].str.contains('Outgoing', case=False, na=False)]
                dials_df = pd.concat([dials_follow_up, dials_predictive, dials_outgoing])
                dials = len(dials_df)

                # Connected (Unique Number/Account)
                connected_numbers = filtered_data[filtered_data['Talk Time Duration'] > 0]['Dialed Number'].nunique()
                connected_accounts = filtered_data[filtered_data['Talk Time Duration'] > 0]['Debtor ID'].nunique()

                # Rates
                penetration = dials / accounts if accounts > 0 else 0
                connected_rate = connected_numbers / accounts if accounts > 0 else 0
                contact_rate = connected_accounts / accounts if accounts > 0 else 0

                # DROPPED (UPDATED LOGIC FOR AGENT: Use "NEGATIVE CALLOUTS - CALL DROP")
                dropped_df = filtered_data[filtered_data['Status'].str.contains('NEGATIVE CALLOUTS - CALL DROP', case=False, na=False)]
                connected_debtors = filtered_data[filtered_data['Talk Time Duration'] > 0]['Debtor ID'].unique()
                dropped = dropped_df[~dropped_df['Debtor ID'].isin(connected_debtors)]['Debtor ID'].nunique()
                dropped_rate = dropped / connected_numbers if connected_numbers > 0 else 0

                # KEPT, PTP, RPC
                kept_df = filtered_data[(filtered_data['Claim Paid Amount'] > 0) & 
                                        (filtered_data['Status'].str.contains('PAYMENT', case=False, na=False))]
                ptp_df = filtered_data[(filtered_data['Status'].str.contains('PTP', case=False, na=False)) & 
                                       (~filtered_data['Status'].str.contains('FF|FOLLOW UP|PTP_FFUP', case=False, na=False)) & 
                                       (filtered_data['PTP Amount'] > 0)]
                rpc_df = filtered_data[filtered_data['Status'].str.contains('RPC -|BANK ESCALATION -', case=False, na=False)]
                rpc_dispo = rpc_df['Debtor ID'].nunique()
                ptp = ptp_df['Debtor ID'].nunique()
                kept = kept_df['Debtor ID'].nunique()
                rpc_pipeline = rpc_dispo + ptp + kept

                # Rates
                rpc_rate = rpc_pipeline / accounts if accounts > 0 else 0
                ptp_rate = (ptp + kept) / rpc_pipeline if rpc_pipeline > 0 else 0
                kept_rate = kept / (ptp + kept) if (ptp + kept) > 0 else 0

                summary_list_agent.append({
                    'Client': client, 'Agent': agent, 'Date Range': date_range, 'Accounts': accounts, 'Dials': dials, 'Penetration': penetration,
                    'Connected (Unique Number)': connected_numbers, 'Connected (Unique Account)': connected_accounts,
                    'Connected Rate': connected_rate, 'Contact Rate': contact_rate, 'Dropped': dropped, 'Drop Rate': dropped_rate,
                    'RPC (Pipeline)': rpc_pipeline, 'RPC (Dispo only)': rpc_dispo, 'PTP': ptp, 'KEPT': kept,
                    'RPC Rate': rpc_rate, 'PTP Rate': ptp_rate, 'KEPT Rate': kept_rate, 'Summary Type': 'Agent'
                })
        combine_end = time.time()
        progress_bar.progress(100 / 100)
        progress_text.text("100% - Processing complete")
        st.write(f"Total processing time: {combine_end - start_time:.2f} seconds")

        if not summary_list_client or not summary_list_agent:
            st.warning("No data found after processing.")
            return None, None

        summary_client = pd.DataFrame(summary_list_client)
        summary_agent = pd.DataFrame(summary_list_agent)
        return summary_client, summary_agent

    except Exception as e:
        st.error(f"Error during processing: {e}")
        return None, None

def to_excel_with_style(df_client, df_agent):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Client sheet
        df_client.to_excel(writer, sheet_name='Client', index=False)
        worksheet_client = writer.sheets['Client']
        for cell in worksheet_client[1]:
            cell.font = Font(bold=True)
        for idx, column in enumerate(df_client.columns):
            max_length = max(df_client[column].astype(str).str.len().max(), len(column)) + 2
            col_letter = get_column_letter(idx + 1)
            worksheet_client.column_dimensions[col_letter].width = max_length

        # Format percentage columns (Client sheet)
        for col in ['Penetration', 'Connected Rate', 'Contact Rate', 'Drop Rate', 'RPC Rate', 'PTP Rate', 'KEPT Rate']:
            if col in df_client.columns:
                col_idx = df_client.columns.get_loc(col) + 1
                col_letter = get_column_letter(col_idx)
                for row in worksheet_client.iter_rows(min_row=2, max_row=len(df_client) + 1, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '0.00%'

        # Format number columns (Client sheet)
        for col in ['Accounts', 'Dials', 'Connected (Unique Number)', 'Connected (Unique Account)', 'Dropped', 'RPC (Pipeline)', 'RPC (Dispo only)', 'PTP', 'KEPT']:
            if col in df_client.columns:
                col_idx = df_client.columns.get_loc(col) + 1
                col_letter = get_column_letter(col_idx)
                for row in worksheet_client.iter_rows(min_row=2, max_row=len(df_client) + 1, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '#,##0'

        # Agent sheet
        df_agent.to_excel(writer, sheet_name='Agent', index=False)
        worksheet_agent = writer.sheets['Agent']
        for cell in worksheet_agent[1]:
            cell.font = Font(bold=True)
        for idx, column in enumerate(df_agent.columns):
            max_length = max(df_agent[column].astype(str).str.len().max(), len(column)) + 2
            col_letter = get_column_letter(idx + 1)
            worksheet_agent.column_dimensions[col_letter].width = max_length

        # Format percentage columns (Agent sheet)
        for col in ['Penetration', 'Connected Rate', 'Contact Rate', 'Drop Rate', 'RPC Rate', 'PTP Rate', 'KEPT Rate']:
            if col in df_agent.columns:
                col_idx = df_agent.columns.get_loc(col) + 1
                col_letter = get_column_letter(col_idx)
                for row in worksheet_agent.iter_rows(min_row=2, max_row=len(df_agent) + 1, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '0.00%'

        # Format number columns (Agent sheet)
        for col in ['Accounts', 'Dials', 'Connected (Unique Number)', 'Connected (Unique Account)', 'Dropped', 'RPC (Pipeline)', 'RPC (Dispo only)', 'PTP', 'KEPT']:
            if col in df_agent.columns:
                col_idx = df_agent.columns.get_loc(col) + 1
                col_letter = get_column_letter(col_idx)
                for row in worksheet_agent.iter_rows(min_row=2, max_row=len(df_agent) + 1, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '#,##0'

    output.seek(0)
    return output

# Main App
st.title("Debt Collection Summary Report")

if uploaded_files:
    summary_client, summary_agent = process_files(uploaded_files, start_date, end_date)
    if summary_client is not None and summary_agent is not None:
        # Format for display
        for col in ['Penetration', 'Connected Rate', 'Contact Rate', 'Drop Rate', 'RPC Rate', 'PTP Rate', 'KEPT Rate']:
            summary_client[col] = summary_client[col].round(2)
            summary_agent[col] = summary_agent[col].round(2)

        # Client Summary
        st.subheader("Summary Table per Client")
        st.dataframe(
            summary_client.style.set_properties(**{'text-align': 'center'}).format({
                'Penetration': '{:.2%}', 'Connected Rate': '{:.2%}', 'Contact Rate': '{:.2%}',
                'Drop Rate': '{:.2%}', 'RPC Rate': '{:.2%}', 'PTP Rate': '{:.2%}', 'KEPT Rate': '{:.2%}',
                'Accounts': '{:,.0f}', 'Dials': '{:,.0f}', 'Connected (Unique Number)': '{:,.0f}',
                'Connected (Unique Account)': '{:,.0f}', 'Dropped': '{:,.0f}', 'RPC (Pipeline)': '{:,.0f}',
                'RPC (Dispo only)': '{:,.0f}', 'PTP': '{:,.0f}', 'KEPT': '{:,.0f}'
            }),
            use_container_width=True
        )

        # Agent Summary
        st.subheader("Summary Table per Agent")
        for client in summary_agent['Client'].unique():
            with st.expander(f"Client: {client}"):
                agent_data = summary_agent[summary_agent['Client'] == client]
                st.dataframe(
                    agent_data.style.set_properties(**{'text-align': 'center'}).format({
                        'Penetration': '{:.2%}', 'Connected Rate': '{:.2%}', 'Contact Rate': '{:.2%}',
                        'Drop Rate': '{:.2%}', 'RPC Rate': '{:.2%}', 'PTP Rate': '{:.2%}', 'KEPT Rate': '{:.2%}',
                        'Accounts': '{:,.0f}', 'Dials': '{:,.0f}', 'Connected (Unique Number)': '{:,.0f}',
                        'Connected (Unique Account)': '{:,.0f}', 'Dropped': '{:,.0f}', 'RPC (Pipeline)': '{:,.0f}',
                        'RPC (Dispo only)': '{:,.0f}', 'PTP': '{:,.0f}', 'KEPT': '{:,.0f}'
                    }),
                    use_container_width=True
                )

        # Download
        excel_file = to_excel_with_style(summary_client, summary_agent)
        st.download_button(
            label="Download All Summaries as XLSX",
            data=excel_file,
            file_name=f"Debt_Collection_Summary_All_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload XLSX files using the sidebar to generate the summary report.")