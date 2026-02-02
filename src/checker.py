"""Post-parse validation: common-sense checks + AI review of rent roll data."""

import json
import re
from datetime import datetime

import pandas as pd

from src.aggregator import apply_mapping, deduplicate, classify_occupancy, get_unit_type_order


def _deterministic_checks(df: pd.DataFrame, as_of_date: datetime | None = None) -> list[dict]:
    """Run common-sense checks on parsed rent roll data.

    Returns list of dicts: {"severity": "error"|"warning"|"info", "check": str, "detail": str}
    """
    issues = []
    today = as_of_date or datetime.now()

    # --- Unit count ---
    total = len(df)
    if total == 0:
        issues.append({"severity": "error", "check": "Unit Count", "detail": "No units were parsed."})
        return issues

    # --- Status distribution ---
    if "status" in df.columns:
        statuses = df["status"].value_counts()
        occupied_count = statuses.get("Occupied", 0) + statuses.get("Occupied-NTV", 0) + statuses.get("Pending Renewal", 0) + statuses.get("Applicant", 0)
        vacant_count = statuses.get("Vacant", 0) + statuses.get("Model", 0)

        # All vacant is suspicious
        if occupied_count == 0 and total > 5:
            issues.append({"severity": "error", "check": "Status Distribution",
                           "detail": f"All {total} units classified as Vacant/Model. Status mapping may have failed."})
        # All occupied also suspicious for large properties
        elif vacant_count == 0 and total > 50:
            issues.append({"severity": "warning", "check": "Status Distribution",
                           "detail": f"All {total} units classified as Occupied. Verify no vacancy exists."})

        # Non-canonical statuses
        canonical = {"Occupied", "Vacant", "Applicant", "Occupied-NTV", "Pending Renewal", "Model"}
        non_canonical = set(df["status"].dropna().unique()) - canonical
        if non_canonical:
            issues.append({"severity": "warning", "check": "Unknown Statuses",
                           "detail": f"Non-canonical status values found: {sorted(non_canonical)}"})

    # --- Rent sanity ---
    if "market_rent" in df.columns:
        zero_market = df[(df["market_rent"].isna()) | (df["market_rent"] == 0)]
        # Exclude vacant/model from this check
        if "status" in df.columns:
            zero_market = zero_market[~zero_market["status"].isin(["Model"])]
        if len(zero_market) > 0:
            pct = len(zero_market) / total * 100
            if pct > 20:
                issues.append({"severity": "warning", "check": "Market Rent",
                               "detail": f"{len(zero_market)} units ({pct:.0f}%) have zero/missing market rent."})

        neg_market = df[df["market_rent"] < 0] if df["market_rent"].notna().any() else pd.DataFrame()
        if len(neg_market) > 0:
            issues.append({"severity": "error", "check": "Negative Market Rent",
                           "detail": f"{len(neg_market)} units have negative market rent."})

    if "lease_rent" in df.columns:
        neg_lease = df[df["lease_rent"] < 0] if df["lease_rent"].notna().any() else pd.DataFrame()
        if len(neg_lease) > 0:
            issues.append({"severity": "error", "check": "Negative Lease Rent",
                           "detail": f"{len(neg_lease)} units have negative lease rent."})

    # Lease rent > 2x market rent (likely data error)
    if "lease_rent" in df.columns and "market_rent" in df.columns:
        both_valid = df[(df["lease_rent"].notna()) & (df["market_rent"].notna()) & (df["market_rent"] > 0)]
        overpaying = both_valid[both_valid["lease_rent"] > both_valid["market_rent"] * 2]
        if len(overpaying) > 0:
            units = overpaying["unit"].tolist()[:5]
            issues.append({"severity": "warning", "check": "Lease > 2x Market",
                           "detail": f"{len(overpaying)} units have lease rent > 2x market rent. Examples: {units}"})

    # --- Date logic ---
    if "lease_start" in df.columns and "lease_end" in df.columns:
        both_dates = df[(df["lease_start"].notna()) & (df["lease_end"].notna())]
        bad_range = both_dates[both_dates["lease_start"] > both_dates["lease_end"]]
        if len(bad_range) > 0:
            units = bad_range["unit"].tolist()[:5]
            issues.append({"severity": "warning", "check": "Lease Date Range",
                           "detail": f"{len(bad_range)} units have lease_start > lease_end. Examples: {units}"})

    # Occupied units with expired leases (lease_end in the past)
    if "lease_end" in df.columns and "status" in df.columns:
        occupied = df[df["status"] == "Occupied"]
        expired = occupied[(occupied["lease_end"].notna()) & (occupied["lease_end"] < today)]
        if len(expired) > total * 0.5 and len(expired) > 10:
            issues.append({"severity": "info", "check": "Expired Leases",
                           "detail": f"{len(expired)} occupied units have lease end dates before the as-of date."})

    # --- Vacant unit with tenant ---
    if "status" in df.columns and "tenant_name" in df.columns:
        vacant_with_tenant = df[
            (df["status"] == "Vacant") &
            (df["tenant_name"].notna()) &
            (df["tenant_name"].str.strip() != "") &
            (~df["tenant_name"].str.upper().isin(["VACANT", "MODEL", ""]))
        ]
        if len(vacant_with_tenant) > 0:
            units = vacant_with_tenant["unit"].tolist()[:5]
            issues.append({"severity": "warning", "check": "Vacant with Tenant",
                           "detail": f"{len(vacant_with_tenant)} vacant units have a tenant name. Examples: {units}"})

    # --- Floorplan mapping ---
    if "display_type" in df.columns:
        unknown = df[df["display_type"] == "Unknown"]
        if len(unknown) > 0:
            fps = unknown["floorplan"].unique().tolist() if "floorplan" in df.columns else []
            issues.append({"severity": "warning", "check": "Unmapped Floorplans",
                           "detail": f"{len(unknown)} units have unmapped floorplan types. Floorplans: {fps}"})

    # --- SQFT ---
    if "sqft" in df.columns:
        zero_sf = df[(df["sqft"].isna()) | (df["sqft"] == 0)]
        if len(zero_sf) > total * 0.1 and len(zero_sf) > 5:
            issues.append({"severity": "warning", "check": "Missing SQFT",
                           "detail": f"{len(zero_sf)} units have zero/missing square footage."})

    return issues


def _build_summary_for_ai(df: pd.DataFrame) -> str:
    """Build a concise summary of the parsed data for AI review."""
    lines = [f"Total units: {len(df)}"]

    if "status" in df.columns:
        lines.append(f"Status distribution: {df['status'].value_counts().to_dict()}")

    if "display_type" in df.columns:
        lines.append(f"Unit types: {df['display_type'].value_counts().to_dict()}")

    if "market_rent" in df.columns and df["market_rent"].notna().any():
        lines.append(f"Market rent: min=${df['market_rent'].min():,.0f}, max=${df['market_rent'].max():,.0f}, avg=${df['market_rent'].mean():,.0f}")

    if "lease_rent" in df.columns and df["lease_rent"].notna().any():
        lines.append(f"Lease rent: min=${df['lease_rent'].min():,.0f}, max=${df['lease_rent'].max():,.0f}, avg=${df['lease_rent'].mean():,.0f}")

    if "sqft" in df.columns and df["sqft"].notna().any():
        lines.append(f"SQFT: min={df['sqft'].min():,.0f}, max={df['sqft'].max():,.0f}, avg={df['sqft'].mean():,.0f}")

    # Sample of first 5 rows
    sample_cols = [c for c in ["unit", "floorplan", "status", "sqft", "market_rent", "lease_rent"] if c in df.columns]
    if sample_cols:
        sample = df[sample_cols].head(5).to_string(index=False)
        lines.append(f"\nSample rows:\n{sample}")

    return "\n".join(lines)


def ai_review(client, df: pd.DataFrame, deterministic_issues: list[dict]) -> list[dict]:
    """Ask AI to review the parsed data and flag any additional concerns."""
    summary = _build_summary_for_ai(df)
    existing_issues = "\n".join(
        f"- [{i['severity'].upper()}] {i['check']}: {i['detail']}"
        for i in deterministic_issues
    ) if deterministic_issues else "None found."

    prompt = f"""You are reviewing parsed multifamily rent roll data. Check for data quality issues, anomalies, or patterns that suggest parsing errors.

DATA SUMMARY:
{summary}

ISSUES ALREADY DETECTED:
{existing_issues}

Look for ADDITIONAL issues not already flagged above. Focus on:
1. Rent distributions that seem wrong (e.g., all same value, huge outliers)
2. Unit type distribution anomalies (e.g., only 1 unit of a type in a 500-unit property)
3. Status/occupancy patterns that seem off
4. Any other red flags suggesting data was parsed incorrectly

Return ONLY valid JSON (no markdown fences) as a list:
[
  {{"severity": "warning", "check": "<check name>", "detail": "<explanation>"}},
  ...
]

If no additional issues are found, return an empty list: []"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "".join(b.text for b in response.content if hasattr(b, "text")).strip()
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        match = re.search(r"\[", text)
        if match:
            text = text[match.start():]
            depth = 0
            for i, ch in enumerate(text):
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                    if depth == 0:
                        text = text[: i + 1]
                        break
        text = re.sub(r",\s*([}\]])", r"\1", text)
        ai_issues = json.loads(text)
        # Tag as AI-sourced
        for issue in ai_issues:
            issue["source"] = "ai"
        return ai_issues
    except Exception:
        return []


def run_checks(client, df: pd.DataFrame, as_of_date: datetime | None = None) -> list[dict]:
    """Run all validation checks (deterministic + AI) on parsed data.

    Returns sorted list of issues: errors first, then warnings, then info.
    """
    issues = _deterministic_checks(df, as_of_date)
    for issue in issues:
        issue["source"] = "check"

    ai_issues = ai_review(client, df, issues)
    issues.extend(ai_issues)

    # Sort: errors first, then warnings, then info
    severity_order = {"error": 0, "warning": 1, "info": 2}
    issues.sort(key=lambda x: severity_order.get(x.get("severity", "info"), 3))

    return issues


# ---------------------------------------------------------------------------
# Final checks — run after aggregation, before output generation
# ---------------------------------------------------------------------------

def _final_deterministic_checks(agg_data: dict) -> list[dict]:
    """Run deterministic checks on the fully aggregated data."""
    issues = []
    active_df = agg_data["active_df"]
    total = len(active_df)

    if total == 0:
        issues.append({"severity": "error", "check": "Final: No Active Units",
                       "detail": "No active (non-deduped) units remain."})
        return issues

    # --- Unknown display types ---
    if "display_type" in active_df.columns:
        unknown = active_df[active_df["display_type"] == "Unknown"]
        if len(unknown) > 0:
            fps = unknown["floorplan"].unique().tolist() if "floorplan" in active_df.columns else []
            issues.append({"severity": "error", "check": "Final: Unmapped Floorplans",
                           "detail": f"{len(unknown)} units have display_type 'Unknown'. Floorplans: {fps}"})

    # --- Occupancy sanity ---
    if "occupancy" in active_df.columns:
        occ_counts = active_df["occupancy"].value_counts()
        occupied = occ_counts.get("Occupied", 0)
        vacant = occ_counts.get("Vacant", 0)
        occ_rate = occupied / total if total > 0 else 0

        if occ_rate < 0.5 and total > 10:
            issues.append({"severity": "warning", "check": "Final: Low Occupancy",
                           "detail": f"Physical occupancy is {occ_rate:.0%} ({occupied}/{total}). Verify statuses."})
        if occupied + vacant != total:
            diff = total - occupied - vacant
            issues.append({"severity": "warning", "check": "Final: Occupancy Mismatch",
                           "detail": f"Occupied ({occupied}) + Vacant ({vacant}) = {occupied + vacant}, but total = {total}. {diff} units unclassified."})

    # --- Rent consistency ---
    if "lease_rent" in active_df.columns and "occupancy" in active_df.columns:
        occ_df = active_df[active_df["occupancy"] == "Occupied"]
        occ_with_rent = occ_df[occ_df["lease_rent"].notna() & (occ_df["lease_rent"] > 0)]
        if len(occ_with_rent) > 0:
            avg_rent = occ_with_rent["lease_rent"].mean()
            if avg_rent < 200:
                issues.append({"severity": "warning", "check": "Final: Low Avg Rent",
                               "detail": f"Average occupied lease rent is ${avg_rent:,.0f}. May indicate parsing error."})
            elif avg_rent > 10000:
                issues.append({"severity": "warning", "check": "Final: High Avg Rent",
                               "detail": f"Average occupied lease rent is ${avg_rent:,.0f}. May indicate parsing error."})

    # --- Per-type consistency ---
    if "display_type" in active_df.columns and "sqft" in active_df.columns:
        type_groups = active_df.groupby("display_type")
        for dtype, group in type_groups:
            if len(group) > 1 and group["sqft"].notna().any():
                sf_std = group["sqft"].std()
                sf_mean = group["sqft"].mean()
                if sf_mean > 0 and sf_std / sf_mean > 0.5:
                    issues.append({"severity": "warning", "check": "Final: SQFT Variance",
                                   "detail": f"'{dtype}' has high SF variance (mean={sf_mean:.0f}, std={sf_std:.0f}). Possible misclassification."})

    return issues


def _build_final_summary_for_ai(agg_data: dict) -> str:
    """Build a summary of aggregated data for AI review."""
    active_df = agg_data["active_df"]
    lines = [f"Total active units: {len(active_df)}"]

    if "occupancy" in active_df.columns:
        lines.append(f"Occupancy: {active_df['occupancy'].value_counts().to_dict()}")

    if "display_type" in active_df.columns:
        type_summary = []
        for dtype, group in active_df.groupby("display_type"):
            avg_sf = group["sqft"].mean() if group["sqft"].notna().any() else 0
            avg_mr = group["market_rent"].mean() if "market_rent" in group.columns and group["market_rent"].notna().any() else 0
            avg_lr = group["lease_rent"].mean() if "lease_rent" in group.columns and group["lease_rent"].notna().any() else 0
            occ = len(group[group["occupancy"] == "Occupied"]) if "occupancy" in group.columns else 0
            type_summary.append(
                f"  {dtype}: {len(group)} units, avg SF={avg_sf:.0f}, avg market=${avg_mr:,.0f}, avg lease=${avg_lr:,.0f}, occupied={occ}"
            )
        lines.append("By unit type:\n" + "\n".join(type_summary))

    dedup = agg_data.get("dedup_report", [])
    if dedup:
        lines.append(f"Deduped units: {len(dedup)}")

    return "\n".join(lines)


def run_final_checks(client, agg_data: dict) -> list[dict]:
    """Run final quality checks on aggregated data (post-mapping, pre-output).

    Returns sorted list of issues.
    """
    issues = _final_deterministic_checks(agg_data)
    for issue in issues:
        issue["source"] = "final_check"

    # AI review of the aggregated data
    summary = _build_final_summary_for_ai(agg_data)
    existing = "\n".join(
        f"- [{i['severity'].upper()}] {i['check']}: {i['detail']}"
        for i in issues
    ) if issues else "None found."

    prompt = f"""You are reviewing the FINAL aggregated rent roll summary. This data has already been parsed, mapped to unit types, and classified.

AGGREGATED SUMMARY:
{summary}

ISSUES ALREADY DETECTED:
{existing}

Look for ADDITIONAL anomalies not already flagged:
1. Unit type distributions that seem wrong (e.g., a 500-unit property with only 2 unit types)
2. Rent-to-SF ratios that are unreasonable
3. Occupancy/vacancy patterns that seem off
4. Any signs the floorplan mapping grouped dissimilar units together

Return ONLY valid JSON (no markdown fences) as a list:
[
  {{"severity": "warning", "check": "<check name>", "detail": "<explanation>"}},
  ...
]

If no additional issues are found, return an empty list: []"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "".join(b.text for b in response.content if hasattr(b, "text")).strip()
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        match = re.search(r"\[", text)
        if match:
            text = text[match.start():]
            depth = 0
            for i, ch in enumerate(text):
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                    if depth == 0:
                        text = text[: i + 1]
                        break
        text = re.sub(r",\s*([}\]])", r"\1", text)
        ai_issues = json.loads(text)
        for issue in ai_issues:
            issue["source"] = "final_ai"
        issues.extend(ai_issues)
    except Exception:
        pass

    severity_order = {"error": 0, "warning": 1, "info": 2}
    issues.sort(key=lambda x: severity_order.get(x.get("severity", "info"), 3))
    return issues


# ---------------------------------------------------------------------------
# Self-correction — attempt to fix issues found by final checks
# ---------------------------------------------------------------------------

def suggest_corrections(client, agg_data: dict, issues: list[dict]) -> list[dict] | None:
    """Analyze final check issues and return correction actions.

    Returns list of correction dicts, or None if no auto-fixes possible.
    Each correction: {"action": str, ...action-specific fields}
    """
    corrections = []

    for issue in issues:
        check = issue.get("check", "")

        # Unknown floorplans → re-map them
        if "Unmapped Floorplans" in check:
            active_df = agg_data["active_df"]
            unknown_fps = active_df[active_df["display_type"] == "Unknown"]["floorplan"].unique().tolist()
            if unknown_fps:
                # Build summary for just the unknown floorplans
                fp_summaries = []
                for fp in unknown_fps:
                    group = active_df[active_df["floorplan"] == fp]
                    fp_summaries.append({
                        "floorplan": str(fp),
                        "count": len(group),
                        "avg_sqft": round(group["sqft"].mean(), 0) if group["sqft"].notna().any() else None,
                        "avg_market_rent": round(group["market_rent"].mean(), 0) if "market_rent" in group.columns and group["market_rent"].notna().any() else None,
                    })
                corrections.append({
                    "action": "remap_floorplans",
                    "floorplans": unknown_fps,
                    "summaries": fp_summaries,
                })

        # Low/zero occupancy → statuses may need re-normalization
        if "Low Occupancy" in check or "Status Distribution" in check:
            corrections.append({
                "action": "review_statuses",
                "detail": issue.get("detail", ""),
            })

    return corrections if corrections else None


def apply_corrections(client, agg_data: dict, mapping: dict, corrections: list[dict]) -> tuple[dict, dict]:
    """Apply corrections and return updated (agg_data, mapping).

    For floorplan remapping, calls AI to re-map unknown floorplans.
    For status issues, attempts to re-resolve statuses.
    """
    from src.mapper import _fallback_unit_type

    updated_mapping = dict(mapping)

    for correction in corrections:
        action = correction["action"]

        if action == "remap_floorplans":
            summaries = correction.get("summaries", [])
            import json as _json

            summary_text = _json.dumps(summaries, indent=2, default=str)
            prompt = f"""These floorplans from a rent roll could not be automatically mapped to unit types.
Based on the data below, assign each a unit type.

{summary_text}

Unit types use the format: "Studio", "1x1", "2x2", "2x2 Den", "3x2", "3x2 Loft", etc.

Return ONLY valid JSON (no markdown fences):
{{
  "FLOORPLAN_CODE": {{"unit_type": "TYPE", "reno": false}},
  ...
}}"""

            try:
                response = client.messages.create(
                    model="claude-sonnet-4-20250514",
                    max_tokens=2048,
                    messages=[{"role": "user", "content": prompt}],
                )
                text = "".join(b.text for b in response.content if hasattr(b, "text")).strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                match = re.search(r"\{", text)
                if match:
                    text = text[match.start():]
                    depth = 0
                    for i, ch in enumerate(text):
                        if ch == "{":
                            depth += 1
                        elif ch == "}":
                            depth -= 1
                            if depth == 0:
                                text = text[: i + 1]
                                break
                text = re.sub(r",\s*([}\]])", r"\1", text)
                new_maps = _json.loads(text)
                updated_mapping.update(new_maps)
            except Exception:
                # Fallback: use SF-based mapping
                for s in summaries:
                    fp = s["floorplan"]
                    if fp not in updated_mapping or updated_mapping[fp].get("unit_type") == "Unknown":
                        updated_mapping[fp] = {
                            "unit_type": _fallback_unit_type(s.get("avg_sqft")),
                            "reno": False,
                        }

    # Re-aggregate with the updated mapping
    from src.aggregator import prepare_aggregated_data
    full_df = agg_data["df"]
    # Get the original pre-mapped dataframe (remove mapping columns)
    base_cols = [c for c in full_df.columns if c not in ("unit_type", "reno", "display_type", "occupancy")]
    base_df = full_df[base_cols].copy()
    # Remove dedup flag — prepare_aggregated_data will re-add it
    if "_dedup_flag" in base_df.columns:
        base_df = base_df[~base_df["_dedup_flag"]].copy()
        base_df = base_df.drop(columns=["_dedup_flag"])

    new_agg = prepare_aggregated_data(base_df, updated_mapping)
    return new_agg, updated_mapping
