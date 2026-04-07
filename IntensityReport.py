from io import BytesIO

import polars as pl
import streamlit as st

st.set_page_config(page_title="Dialer Intensity Report", layout="wide")
st.title("Dialer Intensity Report")
st.markdown("Fast Polars-based processing for large Active List and DRR Excel files.")

ACTIVE_REQUIRED_COLUMNS = {"SUB CAMPAIGN", "CH CODE", "BALANCE BRACKET"}
DRR_REQUIRED_COLUMNS = {
    "Card No.",
    "Remark Type",
    "Remark",
    "PTP Amount",
    "Claim Paid Amount",
    "Status",
}
OUTPUT_COLUMNS = [
    "SUB CAMPAIGN",
    "CH CODE",
    "BALANCE BRACKET",
    "INTENSITY",
    "ACC HAS PTP",
    "ACC HAS KEPT",
    "RPC",
    "TOP 5 STATUS",
]


def read_excel_files(uploaded_files: list, numeric_columns: list[str] | None = None) -> pl.DataFrame:
    frames = []
    numeric_columns = numeric_columns or []

    for uploaded_file in uploaded_files:
        frame = pl.read_excel(BytesIO(uploaded_file.getvalue()), engine="calamine")
        frame = frame.with_columns([pl.col(column).cast(pl.Utf8, strict=False) for column in frame.columns])

        for column in numeric_columns:
            if column in frame.columns:
                frame = frame.with_columns(pl.col(column).cast(pl.Float64, strict=False))

        frames.append(frame)

    return pl.concat(frames, how="diagonal_relaxed") if len(frames) > 1 else frames[0]


def validate_columns(df: pl.DataFrame, required_columns: set[str], label: str) -> None:
    missing_columns = sorted(required_columns - set(df.columns))
    if missing_columns:
        raise ValueError(f"{label} is missing required columns: {', '.join(missing_columns)}")


def build_top_status(dialer_data: pl.DataFrame) -> pl.DataFrame:
    if dialer_data.is_empty():
        return pl.DataFrame(
            schema={
                "CH CODE": pl.Utf8,
                "TOP 5 STATUS": pl.Utf8,
            }
        )

    return (
        dialer_data
        .filter(pl.col("Status").fill_null("").str.strip_chars() != "")
        .with_columns(pl.col("Status").str.strip_chars())
        .group_by(["CH CODE", "Status"])
        .len()
        .sort(["CH CODE", "len", "Status"], descending=[False, True, False])
        .group_by("CH CODE")
        .agg(
            pl.col("Status")
            .head(5)
            .implode()
            .list.join(", ")
            .alias("TOP 5 STATUS")
        )
    )


def process_report(active_files: list, drr_files: list) -> tuple[bytes, int, int]:
    active_list = read_excel_files(active_files)
    merged_drr = read_excel_files(
        drr_files,
        numeric_columns=["PTP Amount", "Claim Paid Amount"],
    )

    validate_columns(active_list, ACTIVE_REQUIRED_COLUMNS, "Active List")
    validate_columns(merged_drr, DRR_REQUIRED_COLUMNS, "DRR")

    follow_up_filter = merged_drr.filter(
        (pl.col("Remark Type").fill_null("").str.to_lowercase() == "follow up")
        & pl.col("Remark").fill_null("").str.to_lowercase().str.contains("predictive cp")
    )

    predictive_outgoing_filter = merged_drr.filter(
        pl.col("Remark Type").fill_null("").str.to_lowercase().is_in(["predictive", "outgoing"])
    )

    dialer_data = (
        pl.concat([follow_up_filter, predictive_outgoing_filter], how="diagonal_relaxed")
        .with_columns(pl.col("Card No.").alias("CH CODE"))
    )

    metrics_df = (
        dialer_data
        .group_by("CH CODE")
        .agg([
            pl.len().alias("INTENSITY"),
            pl.col("PTP Amount").filter(pl.col("PTP Amount") > 0).count().alias("ACC HAS PTP"),
            pl.col("Claim Paid Amount").filter(pl.col("Claim Paid Amount") > 0).count().alias("ACC HAS KEPT"),
            pl.col("Status")
            .filter(
                pl.col("Status")
                .fill_null("")
                .str.to_lowercase()
                .str.contains("rpc|bank escalation|ptp")
            )
            .count()
            .alias("RPC"),
        ])
    )

    top_status_df = build_top_status(dialer_data)

    output_df = (
        active_list
        .select(["SUB CAMPAIGN", "CH CODE", "BALANCE BRACKET"])
        .join(metrics_df, on="CH CODE", how="left")
        .join(top_status_df, on="CH CODE", how="left")
        .with_columns([
            pl.col("INTENSITY").fill_null(0),
            pl.col("ACC HAS PTP").fill_null(0),
            pl.col("ACC HAS KEPT").fill_null(0),
            pl.col("RPC").fill_null(0),
            pl.col("TOP 5 STATUS").fill_null(""),
        ])
        .select(OUTPUT_COLUMNS)
    )

    output_buffer = BytesIO()
    output_df.write_excel(
        output_buffer,
        worksheet="Dialer Intensity Report",
        header_format={
            "bold": True,
            "bg_color": "#1F2329",
            "font_color": "#FFFFFF",
        },
        autofit=True,
    )
    output_buffer.seek(0)
    return output_buffer.getvalue(), active_list.height, dialer_data.height


if "report_bytes" not in st.session_state:
    st.session_state.report_bytes = None
    st.session_state.processed_accounts = 0
    st.session_state.filtered_efforts = 0

active_files = st.file_uploader(
    "Active List XLSX file(s)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="active_files",
)

drr_files = st.file_uploader(
    "DRR XLSX file(s)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="drr_files",
)

if st.button("Process", type="primary"):
    try:
        if not active_files or not drr_files:
            raise ValueError("Please upload both Active List and DRR XLSX files before processing.")

        with st.spinner("Processing files..."):
            report_bytes, processed_accounts, filtered_efforts = process_report(active_files, drr_files)

        st.session_state.report_bytes = report_bytes
        st.session_state.processed_accounts = processed_accounts
        st.session_state.filtered_efforts = filtered_efforts
        st.success("Report generated successfully.")
    except Exception as exc:
        st.session_state.report_bytes = None
        st.session_state.processed_accounts = 0
        st.session_state.filtered_efforts = 0
        st.error(f"Processing failed: {exc}")

if st.session_state.report_bytes is not None:
    st.download_button(
        label="Download Dialer Intensity Report",
        data=st.session_state.report_bytes,
        file_name="Dialer_Intensity_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.info(
        f"Processed **{st.session_state.processed_accounts:,}** accounts and "
        f"**{st.session_state.filtered_efforts:,}** filtered dialer efforts."
    )

st.caption("Required Active List columns: SUB CAMPAIGN, CH CODE, BALANCE BRACKET.")
