import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="VB Horizontal", layout="wide")
st.title("Voice Blast – Horizontal Debtor IDs (Memory Safe)")

uploaded_files = st.sidebar.file_uploader(
    "Upload XLSX files", type=["xlsx"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload files →")
    st.stop()

# Process files ONE BY ONE without loading everything at once
@st.cache_data(show_spinner="Processing files...")
def extract_vb_data(files):
    all_vb = []
    for file in files:
        # Read only required columns to save memory
        df = pd.read_excel(file, usecols=["Remark Type", "Remark", "Client", "Status", "Debtor ID"])
        
        # Filter early
        mask = (
            (df["Remark Type"].astype(str).str.strip() == "Follow Up") &
            (df["Remark"].astype(str).str.lower().str.contains("broadcast cp"))
        )
        filtered = df[mask][["Client", "Status", "Debtor ID"]].copy()
        
        if not filtered.empty:
            filtered["Client"] = filtered["Client"].astype(str).str.strip()
            filtered["Debtor ID"] = filtered["Debtor ID"].astype(str).str.strip()
            all_vb.append(filtered)
    
    return pd.concat(all_vb, ignore_index=True) if all_vb else pd.DataFrame()

vb = extract_vb_data(uploaded_files)

if vb.empty:
    st.error("No Voice Blast records found.")
    st.stop()

st.success(f"Found {len(vb):,} VB records")

target_status = ["BUSY", "PM", "PU", "RNA"]

# Build Excel in memory-safe way
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    for campaign in sorted(vb["Client"].unique()):
        camp_data = vb[vb["Client"] == campaign]
        
        # Group once
        grouped = camp_data.groupby("Status")["Debtor ID"].apply(lambda x: x.dropna().unique().tolist())
        
        # Get max length for rows
        lists = {status: grouped.get(status, []) for status in target_status}
        max_len = max(len(lst) for lst in lists.values())
        
        # Build rows efficiently
        rows = []
        for i in range(max_len):
            row = {status: lists[status][i] if i < len(lists[status]) else "" for status in target_status}
            rows.append(row)
        
        result_df = pd.DataFrame(rows, columns=target_status)
        
        sheet_name = str(campaign)[:31]
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Prevent empty workbook
    if len(writer.sheets) == 0:
        pd.DataFrame({"No Data": ["No campaigns found"]}).to_excel(writer, "Summary")

output.seek(0)

st.download_button(
    label="Download Horizontal VB (BUSY PM PU RNA)",
    data=output,
    file_name=f"VB_Horizontal_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Light preview
st.subheader("Preview (Counts only)")
summary = vb.groupby(["Client", "Status"]).size().unstack(fill_value=0)[target_status]
st.dataframe(summary, use_container_width=True)