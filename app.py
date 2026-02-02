"""Streamlit Rent Roll Parser — AI-powered rent roll normalization and summary."""

from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from anthropic import Anthropic

WORKSPACE_ROOT = Path("workspace")
WORKSPACE_ROOT.mkdir(exist_ok=True)


def create_workspace(filename: str) -> Path:
    """Create a timestamped workspace directory for this upload."""
    stem = Path(filename).stem
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ws_dir = WORKSPACE_ROOT / f"{stem}_{ts}"
    ws_dir.mkdir(parents=True, exist_ok=True)
    return ws_dir

from src.parser import parse_rent_roll
from src.mapper import propose_mapping, mapping_to_df, df_to_mapping
from src.aggregator import prepare_aggregated_data
from src.output import generate_output

st.set_page_config(page_title="Rent Roll Parser", layout="wide")
st.title("Rent Roll Parser")

# ─── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    # API key — try secrets.toml first, fall back to empty
    try:
        default_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        default_key = ""
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=default_key,
        help="Enter your Anthropic API key or set ANTHROPIC_API_KEY in .streamlit/secrets.toml",
    )

    st.divider()
    st.header("Trailing Periods")
    st.caption("Select trailing periods (relative to the rent roll's As-of Date) for move-in date filtering.")
    t30 = st.checkbox("T-30 (Last 30 Days)", value=True, key="t30")
    t60 = st.checkbox("T-60 (Last 60 Days)", value=True, key="t60")
    t90 = st.checkbox("T-90 (Last 90 Days)", value=True, key="t90")

# ─── Main Area ───────────────────────────────────────────────────────────────

# Step 1: File upload
uploaded_file = st.file_uploader("Upload Rent Roll", type=["xlsx", "xls", "csv"])

if uploaded_file and not api_key:
    st.warning("Please enter your Anthropic API key in the sidebar.")
    st.stop()

if uploaded_file and api_key:
    client = Anthropic(api_key=api_key)
    file_bytes = uploaded_file.getvalue()

    # Create workspace directory for this upload
    if "workspace_dir" not in st.session_state or st.session_state.get("_last_filename") != uploaded_file.name:
        ws_dir = create_workspace(uploaded_file.name)
        input_path = ws_dir / uploaded_file.name
        input_path.write_bytes(file_bytes)
        st.session_state.workspace_dir = ws_dir
        st.session_state._last_filename = uploaded_file.name

    st.caption(f"Workspace: `{st.session_state.workspace_dir}`")

    # Step 2: Parse
    if "parsed" not in st.session_state:
        st.session_state.parsed = None
    if "mapping_df" not in st.session_state:
        st.session_state.mapping_df = None
    if "agg_data" not in st.session_state:
        st.session_state.agg_data = None

    if st.button("Parse Rent Roll", type="primary"):
        with st.spinner("Parsing rent roll with AI..."):
            result = parse_rent_roll(client, file_bytes, uploaded_file.name)
            st.session_state.parsed = result
            st.session_state.mapping_df = None
            st.session_state.agg_data = None

    if st.session_state.parsed:
        parsed = st.session_state.parsed
        df = parsed["df"]

        as_of_date = parsed.get("as_of_date")
        if as_of_date:
            st.info(f"As-of Date: {as_of_date.strftime('%m/%d/%Y')}")
        else:
            st.warning("Could not detect an As-of Date in the rent roll. Trailing period filters will be disabled.")

        st.subheader("Extracted Data Preview")
        preview_cols = ["unit", "floorplan", "sqft", "status", "tenant_name",
                        "market_rent", "lease_rent", "total_billing"]
        display_cols = [c for c in preview_cols if c in df.columns]
        st.dataframe(df[display_cols].head(20), use_container_width=True)
        st.caption(f"{len(df)} units extracted from {parsed['sheet_name']}")

        # Step 3: Mapping
        st.divider()
        st.subheader("Floorplan → Unit Type Mapping")

        if st.session_state.mapping_df is None:
            if st.button("Propose Mapping with AI"):
                with st.spinner("Analyzing floorplans..."):
                    mapping = propose_mapping(client, df)
                    st.session_state.mapping_df = mapping_to_df(mapping)
                    st.rerun()
        else:
            st.caption("Review and edit the proposed mapping, then click Generate Summary.")
            edited_df = st.data_editor(
                st.session_state.mapping_df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Floorplan": st.column_config.TextColumn(disabled=True),
                    "Unit Type": st.column_config.TextColumn(),
                    "Reno": st.column_config.CheckboxColumn(),
                },
            )
            st.session_state.mapping_df = edited_df

            # Step 4: Generate
            st.divider()
            if st.button("Generate Summary", type="primary"):
                mapping = df_to_mapping(edited_df)

                with st.spinner("Aggregating and building output..."):
                    agg_data = prepare_aggregated_data(df, mapping)
                    st.session_state.agg_data = agg_data

                    # Build date_ranges from trailing period checkboxes + as-of date
                    date_ranges = []
                    as_of = parsed.get("as_of_date")
                    if as_of:
                        for enabled, days, label in [
                            (t30, 30, "T-30"),
                            (t60, 60, "T-60"),
                            (t90, 90, "T-90"),
                        ]:
                            if enabled:
                                date_ranges.append({
                                    "label": label,
                                    "start": as_of - timedelta(days=days),
                                    "end": as_of,
                                })

                    output_bytes = generate_output(
                        agg_data=agg_data,
                        column_map=parsed["column_map"],
                        raw_wb_bytes=file_bytes,
                        sheet_name=parsed["sheet_name"],
                        date_ranges=date_ranges,
                        as_of_date=as_of,
                    )
                    st.session_state.output_bytes = output_bytes

                    # Save output to workspace directory
                    if st.session_state.get("workspace_dir"):
                        out_name = f"{Path(uploaded_file.name).stem}_Output.xlsx"
                        output_path = st.session_state.workspace_dir / out_name
                        output_path.write_bytes(output_bytes)
                        st.success(f"Output saved to `{output_path}`")

    # Show results if available
    if st.session_state.get("agg_data"):
        agg_data = st.session_state.agg_data

        # Dedup report
        dedup_report = agg_data.get("dedup_report", [])
        if dedup_report:
            st.divider()
            st.subheader("Deduplication Report")
            st.caption(f"{len(dedup_report)} unit(s) had duplicate rows resolved.")
            dedup_df = pd.DataFrame(dedup_report)
            st.dataframe(dedup_df, use_container_width=True)

        # Summary preview
        st.divider()
        st.subheader("Summary Preview")
        active_df = agg_data["active_df"]
        summary_preview = active_df.groupby("display_type").agg(
            Units=("unit", "count"),
            Avg_SF=("sqft", "mean"),
            Total_SF=("sqft", "sum"),
            Avg_Market_Rent=("market_rent", "mean"),
            Avg_Lease_Rent=("lease_rent", "mean"),
        ).round(0)
        # Reorder
        type_order = agg_data["unit_type_order"]
        summary_preview = summary_preview.reindex(
            [t for t in type_order if t in summary_preview.index]
        )
        st.dataframe(summary_preview, use_container_width=True)

    # Download button
    if st.session_state.get("output_bytes"):
        st.divider()
        st.download_button(
            label="Download Output Excel",
            data=st.session_state.output_bytes,
            file_name=f"{Path(uploaded_file.name).stem}_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
