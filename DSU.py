import streamlit as st
import pandas as pd
import numpy as np

# -------------------------- Streamlit App --------------------------
st.set_page_config(page_title="Dialer Performance Report", layout="wide")
st.title("📊 Dialer Performance Report Generator")
st.markdown("Upload multiple Excel (.xlsx) files to generate consolidated dialer performance metrics.")

# Sidebar file uploader
uploaded_files = st.sidebar.file_uploader(
    "Upload Excel Files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="You can upload multiple Excel files at once"
)

if not uploaded_files:
    st.info("👆 Please upload one or more Excel files in the sidebar.")
    st.stop()

# -------------------------- Processing --------------------------
@st.cache_data(show_spinner="Processing uploaded files...")
def load_and_process_files(files):
    all_dfs = []
    
    for file in files:
        try:
            # Read all sheets? Or just first? Adjust if needed
            df = pd.read_excel(file, dtype={'Debtor ID': str})  # Keep Debtor ID as string
            df['source_file'] = file.name  # Optional: track source
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")
            continue
    
    if not all_dfs:
        st.error("No valid data could be loaded from the uploaded files.")
        st.stop()
    
    data = pd.concat(all_dfs, ignore_index=True)
    
    # Clean column names (remove extra spaces)
    data.columns = data.columns.str.strip()
    
    # Required columns check
    required_cols = ['Client', 'Debtor ID', 'Remark Type', 'Remark', 'Talk Time Duration', 'Status']
    missing = [col for col in required_cols if col not in data.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        st.stop()
    
    # Convert Talk Time Duration to numeric (in case it's string like "00:01:23")
    if data['Talk Time Duration'].dtype == 'object':
        data['Talk Time Duration'] = pd.to_timedelta(data['Talk Time Duration'], errors='coerce').dt.total_seconds()
    else:
        data['Talk Time'] = pd.to_numeric(data['Talk Time Duration'], errors='coerce')
    
    # Define the two dialer performance criteria
    mask_predictive_cp = (
        (data['Remark Type'] == 'Follow Up') &
        (data['Remark'].str.contains('predictive cp', case=False, na=False))
    )
    
    mask_outgoing_predictive = data['Remark Type'].isin(['Outgoing', 'Predictive'])
    
    # Final dialer performance mask
    dialer_mask = mask_predictive_cp | mask_outgoing_predictive
    dialer_data = data[dialer_mask].copy()
    
    if dialer_data.empty:
        st.warning("No records match the dialer performance criteria.")
        st.stop()
    
    # Pre-define RPC keywords
    rpc_keywords = ['bank escalation', 'rpc', 'ptp', 'payment']
    
    # Group by Client
    summary = []
    
    for client, group in dialer_data.groupby('Client'):
        # WOA: Unique Debtor ID in dialer data
        woa = group['Debtor ID'].nunique()
        
        # TOTAL DIALED: All attempts (count of rows with Debtor ID)
        total_dialed = group['Debtor ID'].count()
        
        # CONNECTED: Talk Time > 0 (and not NaN)
        connected_mask = group['Talk Time'] > 0
        connected_debtors = group.loc[connected_mask, 'Debtor ID']
        connected = connected_debtors.count()
        connected_unique = connected_debtors.nunique()
        
        # TOTAL RPC: Status contains any of the keywords + unique Debtor ID
        rpc_mask = group['Status'].str.contains('|'.join(rpc_keywords), case=False, na=False)
        rpc_unique = group.loc[rpc_mask, 'Debtor ID'].nunique()
        
        summary.append({
            'CLIENT': client,
            'WOA': woa,
            'TOTAL DIALED': total_dialed,
            'CONNECTED': connected,
            'CONNECTED UNIQUE': connected_unique,
            'TOTAL RPC': rpc_unique
        })
    
    result_df = pd.DataFrame(summary)
    
    # Sort by CLIENT
    result_df = result_df.sort_values('CLIENT').reset_index(drop=True)
    
    return result_df

# Run processing
with st.spinner("Analyzing data across all files..."):
    final_report = load_and_process_files(uploaded_files)

# -------------------------- Display Results --------------------------
st.success(f"✅ Processed {len(uploaded_files)} file(s) | Found {len(final_report)} clients")

st.dataframe(final_report, use_container_width=True, hide_index=True)

# Download button
csv = final_report.to_csv(index=False).encode()
st.download_button(
    label="📥 Download Report as CSV",
    data=csv,
    file_name=f"dialer_performance_report_{pd.Timestamp('today').strftime('%Y%m%d')}.csv",
    mime="text/csv"
)

st.caption("Report generated on " + pd.Timestamp('today').strftime('%B %d, %Y'))