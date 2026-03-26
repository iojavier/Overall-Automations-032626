import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration ---
# Set the page configuration at the very start
st.set_page_config(
    page_title="Excel Data Processor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Define the campaigns that use 'Debtor ID' for grouping
DEBTOR_ID_CAMPAIGNS = {
    "BPI RBANK PL SL": "Debtor ID",
    "BPI AUTO CURING SL": "BPI AUTO CURING SL", # Adjusted based on common column names
    "BPI RBANK PL 60DPD SL": "Debtor ID"
}

# Default grouping column for ALL other campaigns
CARD_NO_CAMPAIGN_KEY = "Card No." 

# Define the columns required for the final output
DISPLAY_ORDER = ['Status', 'Card No.', 'Remark']
STATUS_INCLUSIONS = [
    'DROPPED', 'NEGATIVE CALLOUTS', 'BANK ESCALATION',
    'BP', 'PAYMENT', 'PTP', 'RPC', 'TPC'
]
STATUS_REGEX = '|'.join(STATUS_INCLUSIONS)


# --- Core Processing Function (Cached for Performance) ---

@st.cache_data(show_spinner="Processing files... This may take a moment.")
def process_uploaded_files(uploaded_files):
    """
    Processes multiple uploaded XLSX files, filters data based on campaign 
    and status, and dynamically assigns a grouping key ('Debtor ID' or 'Card No.').
    """
    
    # 1. Combine all data first (Omitted for brevity, assuming standard start)
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            df.columns = df.columns.str.strip()
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Error reading file **{uploaded_file.name}**: {e}")
            
    if not all_dfs: return {}
    master_df = pd.concat(all_dfs, ignore_index=True)
    
    required_cols = ['Client', 'Status', 'Card No.', 'Debtor ID']
    if not all(col in master_df.columns for col in required_cols):
        missing_cols = [col for col in required_cols if col not in master_df.columns]
        st.error(f"Missing one or more required columns in the uploaded files: {', '.join(missing_cols)}")
        return {}

    # 2. Filter by Status
    status_mask = master_df['Status'].astype(str).str.contains(STATUS_REGEX, case=False, na=False)
    filtered_df = master_df[status_mask].copy()

    if filtered_df.empty: return {}
        
    final_result = {}
    unique_campaigns = filtered_df['Client'].dropna().unique()

    
    # Helper function for column selection (using set for guaranteed uniqueness)
    def get_final_cols_unique(df, group_col):
        """Constructs a unique, ordered list of columns for the final DataFrame."""
        
        # 1. Start with the unique set of required columns
        required_cols_set = set(DISPLAY_ORDER)
        required_cols_set.add(group_col)
        
        # 2. Define the desired order of columns
        # We start with DISPLAY_ORDER, then append group_col IF it's not already in DISPLAY_ORDER
        ordered_list = list(DISPLAY_ORDER)
        if group_col not in ordered_list:
            ordered_list.append(group_col)
        
        # 3. Create the final list by keeping only columns that exist in the DataFrame
        # and are in our unique ordered list. THIS IS THE CRITICAL CHANGE.
        final_cols = [col for col in ordered_list if col in df.columns]
        
        return final_cols

    # --- A. Process Debtor ID Campaigns ---
    debtor_id_campaign_names = set(DEBTOR_ID_CAMPAIGNS.keys())
    
    for campaign in debtor_id_campaign_names:
        if campaign in unique_campaigns:
            campaign_df = filtered_df[filtered_df['Client'] == campaign].copy()
            group_col = DEBTOR_ID_CAMPAIGNS[campaign]
            
            # Use the corrected helper function
            final_cols = get_final_cols_unique(campaign_df, group_col)
            final_result[campaign] = campaign_df.loc[:, final_cols]

    # --- B. Process Card No. Campaigns (All others) ---
    card_no_campaign_names = [c for c in unique_campaigns if c not in debtor_id_campaign_names]

    for campaign in card_no_campaign_names:
        campaign_df = filtered_df[filtered_df['Client'] == campaign].copy()
        group_col = CARD_NO_CAMPAIGN_KEY # This is 'Card No.'
        
        # Use the corrected helper function
        final_cols = get_final_cols_unique(campaign_df, group_col)
        final_result[campaign] = campaign_df.loc[:, final_cols]

    return final_result

# --- Streamlit Application Layout (Remains the same) ---

def main():
    st.title("📊 Multi-File Excel Campaign Processor")
    st.markdown("""
        Upload your large XLSX files to filter data based on **Campaign** (`Client` column) 
        and **Status** column inclusions.
        """)
    
    # 1. Sidebar Uploader (Handle multiple files)
    with st.sidebar:
        st.header("Upload Files (.xlsx)")
        uploaded_files = st.file_uploader(
            "Select one or more Excel files", 
            type=['xlsx'], 
            accept_multiple_files=True
        )

    # 2. Main Logic
    if uploaded_files:
        st.info(f"Processing **{len(uploaded_files)}** file(s)...")
        
        # Call the cached function to process the files
        processed_data = process_uploaded_files(uploaded_files)
        
        if processed_data:
            st.success("✅ Data processing complete! Results below:")
            
            # --- Results Display and Download ---
            
            # Create the Excel file for download in memory (Fastest method)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for campaign, df in processed_data.items():
                    # Clean the sheet name to prevent Excel errors
                    sheet_name = campaign[:31].replace(':', '_').replace('[', '').replace(']', '').replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_')
                    df.to_excel(writer, sheet_name=sheet_name, index=False) 
            
            # Rewind the buffer
            output.seek(0)
            
            # Display a Download Button
            st.download_button(
                label="⬇️ Download Processed Excel File",
                data=output,
                file_name="processed_campaign_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_button"
            )
            
            st.markdown("---")
            st.subheader("Preview of Processed Data")
            
            # Display data preview using tabs
            tabs = st.tabs(list(processed_data.keys()))
            
            for i, (campaign, df) in enumerate(processed_data.items()):
                with tabs[i]:
                    group_col_used = DEBTOR_ID_CAMPAIGNS.get(campaign, CARD_NO_CAMPAIGN_KEY)
                    st.metric(label=f"Total Rows for **{campaign}** (Grouped by **{group_col_used}**)", value=f"{len(df):,}")
                    st.dataframe(df)

        else:
            st.warning("No matching data found in the uploaded files based on the specified criteria.")

    else:
        st.info("Please upload one or more XLSX files in the sidebar to begin.")
        st.markdown("---")
        st.subheader("Configuration Details")
        st.markdown(f"""
            * **Grouped by 'Debtor ID'**: {', '.join(DEBTOR_ID_CAMPAIGNS.keys())}
            * **Grouped by 'Card No.'**: All other campaigns
            * **Columns Extracted**: {', '.join(DISPLAY_ORDER)} and the relevant grouping ID.
            * **Status Inclusion List**: {', '.join(STATUS_INCLUSIONS)}
            """)

if __name__ == "__main__":
    main()