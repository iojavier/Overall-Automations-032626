import streamlit as st
import pandas as pd

st.title("New Endorsements Per Campaign")
st.sidebar.header("Upload Files")
uploaded_files = st.sidebar.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_records = []

    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.upper()

        # Ensure columns exist
        if 'DEBTOR ID' not in df.columns or 'CLIENT' not in df.columns:
            st.warning(f"Skipping {file.name} - missing DEBTOR ID or CLIENT column")
            continue

        # Filter new endorsements
        mask1 = df['STATUS'].astype(str).str.contains('NEW', case=False, na=False)
        mask2 = df['REMARK'].astype(str).str.contains('New files imported', case=False, na=False)
        df = df[mask1 & mask2]

        if df.empty:
            continue

        # Clean and select
        temp = df[['DEBTOR ID', 'CLIENT']].copy()
        temp['DEBTOR ID'] = temp['DEBTOR ID'].astype(str).str.strip()
        temp['CLIENT'] = temp['CLIENT'].astype(str).str.strip().str.upper()

        all_records.append(temp)

    if all_records:
        combined = pd.concat(all_records, ignore_index=True)

        # Remove duplicate Debtor IDs (count each debtor only once)
        unique = combined.drop_duplicates(subset='DEBTOR ID')

        # Count per campaign
        summary = unique['CLIENT'].value_counts().reset_index()
        summary.columns = ['Campaign', 'New Endo Count']
        summary = summary.sort_values('New Endo Count', ascending=False)

        st.write(f"**Total Unique New Endorsements: {len(unique)}**")
        st.dataframe(summary)

        # Downloads
        st.download_button("Download Full List", unique.to_csv(index=False), "new_endo_full.csv", "text/csv")
        st.download_button("Download Summary", summary.to_csv(index=False), "new_endo_summary.csv", "text/csv")
    else:
        st.info("No new endorsements found in the uploaded files.")

else:
    st.info("Upload your daily remark Excel files to get started.")