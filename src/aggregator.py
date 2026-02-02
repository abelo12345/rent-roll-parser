"""Step 3: Dedup, classify, and prepare data for formula-driven output."""

import pandas as pd


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    """Deduplicate units: when a unit has both Vacant and Applicant rows, keep Applicant.

    Adds a '_dedup_flag' column: True for rows that were dropped (kept for reference).
    """
    df = df.copy()
    df["_dedup_flag"] = False

    # Group by unit
    grouped = df.groupby("unit")
    drop_indices = []

    for unit, group in grouped:
        if len(group) <= 1:
            continue
        statuses = set(group["status"].values)
        # If both Vacant and Applicant exist, drop the Vacant row
        if "Vacant" in statuses and "Applicant" in statuses:
            vacant_rows = group[group["status"] == "Vacant"].index
            drop_indices.extend(vacant_rows.tolist())

    df.loc[drop_indices, "_dedup_flag"] = True
    return df


def classify_occupancy(status: str) -> str:
    """Classify status into Occupied or Vacant for summary purposes.

    Permissive: anything not clearly vacant is treated as occupied.
    This is safer because phantom vacancy (misclassifying occupied as vacant)
    is worse than the reverse.
    """
    if not status:
        return "Vacant"
    lower = status.lower().strip()
    if any(v in lower for v in ("vacant", "model")):
        return "Vacant"
    return "Occupied"


def apply_mapping(df: pd.DataFrame, mapping: dict[str, dict]) -> pd.DataFrame:
    """Apply the confirmed unit type mapping to the DataFrame."""
    df = df.copy()
    df["unit_type"] = df["floorplan"].map(
        lambda fp: mapping.get(str(fp), {}).get("unit_type", "Unknown")
    )
    df["reno"] = df["floorplan"].map(
        lambda fp: mapping.get(str(fp), {}).get("reno", False)
    )
    df["display_type"] = df.apply(
        lambda r: f"{r['unit_type']} Reno" if r["reno"] else r["unit_type"], axis=1
    )
    df["occupancy"] = df["status"].map(classify_occupancy)
    return df


def _bed_count(unit_type: str) -> int:
    """Extract bedroom count from unit type name for sorting."""
    base = unit_type.replace(" Reno", "").strip()
    if base.lower() == "studio":
        return 0
    # Parse "NxN" or "NxN Den" or "NxN Loft" etc.
    parts = base.split("x")
    if parts and parts[0].isdigit():
        return int(parts[0])
    return 99  # unknown goes last


def get_unit_type_order(df: pd.DataFrame) -> list[str]:
    """Return display types: non-reno sorted by (bed count, avg SF), then reno same way."""
    non_reno = []
    reno = []

    for dt, group in df[~df["_dedup_flag"]].groupby("display_type"):
        avg_sf = group["sqft"].mean() if group["sqft"].notna().any() else 0
        beds = _bed_count(dt)
        if dt.endswith(" Reno"):
            reno.append((dt, beds, avg_sf))
        else:
            non_reno.append((dt, beds, avg_sf))

    non_reno.sort(key=lambda x: (x[1], x[2]))
    reno.sort(key=lambda x: (x[1], x[2]))

    return [t[0] for t in non_reno] + [t[0] for t in reno]


def prepare_aggregated_data(df: pd.DataFrame, mapping: dict[str, dict]) -> dict:
    """Full aggregation pipeline. Returns dict with:
    - 'df': the full DataFrame with mapping, dedup flags, classification
    - 'active_df': only non-deduped rows (the ones used in the summary)
    - 'unit_type_order': ordered list of display types for the summary
    - 'dedup_report': list of deduped units with details
    """
    df = deduplicate(df)
    df = apply_mapping(df, mapping)

    # Applicants with no lease rent — assume market rent
    mask = (df["status"] == "Applicant") & (df["lease_rent"].isna() | (df["lease_rent"] == 0))
    df.loc[mask, "lease_rent"] = df.loc[mask, "market_rent"]

    # Build dedup report
    dedup_report = []
    deduped = df[df["_dedup_flag"]]
    for _, row in deduped.iterrows():
        unit = row["unit"]
        # Find the kept row for this unit
        kept = df[(df["unit"] == unit) & (~df["_dedup_flag"])]
        if not kept.empty:
            kept_row = kept.iloc[0]
            dedup_report.append({
                "unit": unit,
                "dropped_status": row["status"],
                "dropped_row": row["_source_row"],
                "kept_status": kept_row["status"],
                "kept_row": kept_row["_source_row"],
            })

    active_df = df[~df["_dedup_flag"]].copy().reset_index(drop=True)
    unit_type_order = get_unit_type_order(df)

    return {
        "df": df,
        "active_df": active_df,
        "unit_type_order": unit_type_order,
        "dedup_report": dedup_report,
    }
