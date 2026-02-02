"""Step 2: LLM-proposed floorplan → unit type mapping."""

import json
import re

import pandas as pd


def _build_floorplan_summary(df: pd.DataFrame) -> list[dict]:
    """Build a summary of unique floorplans with stats for the LLM."""
    summaries = []
    for fp, group in df.groupby("floorplan"):
        summaries.append({
            "floorplan": str(fp),
            "count": len(group),
            "avg_sqft": round(group["sqft"].mean(), 0) if group["sqft"].notna().any() else None,
            "min_sqft": int(group["sqft"].min()) if group["sqft"].notna().any() else None,
            "max_sqft": int(group["sqft"].max()) if group["sqft"].notna().any() else None,
            "avg_market_rent": round(group["market_rent"].mean(), 0) if group["market_rent"].notna().any() else None,
            "avg_lease_rent": round(group["lease_rent"].dropna().mean(), 0) if group["lease_rent"].notna().any() else None,
        })
    return sorted(summaries, key=lambda x: (x["avg_sqft"] or 0))


def _repair_truncated_json(text: str) -> str:
    """Attempt to repair JSON that was truncated mid-output.

    Strips the last incomplete entry back to the previous complete one,
    then closes any open braces.
    """
    # Remove any trailing partial key-value pair (after the last complete entry)
    # Look for the last complete "}: ," or "}" pattern
    last_complete = text.rfind("}")
    if last_complete > 0:
        text = text[: last_complete + 1]

    # Remove trailing comma if present
    text = text.rstrip().rstrip(",")

    # Count unclosed braces and close them
    depth = 0
    for ch in text:
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
    text += "}" * depth

    return text


def _fallback_unit_type(avg_sqft: float | None) -> str:
    """Assign unit type based on SF when AI mapping is missing."""
    if avg_sqft is None:
        return "Unknown"
    if avg_sqft < 600:
        return "Studio"
    elif avg_sqft < 950:
        return "1x1"
    elif avg_sqft < 1250:
        return "2x2"
    elif avg_sqft < 1600:
        return "2x2"
    else:
        return "3x2"


def propose_mapping(client, df: pd.DataFrame) -> dict[str, dict]:
    """Ask Claude to propose a floorplan → unit type mapping.

    Returns dict like:
    {
        "B4": {"unit_type": "2x2", "reno": false},
        "S1": {"unit_type": "Studio", "reno": false},
        ...
    }
    """
    summary = _build_floorplan_summary(df)
    summary_text = json.dumps(summary, indent=2, default=str)

    prompt = f"""You are analyzing a multifamily rent roll. Below is a summary of unique floorplan codes with unit counts and rent/SF stats.

{summary_text}

For each floorplan code, propose:
1. A standardized **unit type** name using the format: "Studio", "1x1", "1x1 Loft", "2x2", "2x2 Den", "2x3 Loft", "3x2", "3x2 Loft", etc.
   - The format is BedroomsxBathrooms, optionally followed by "Den" or "Loft" for larger variants
   - Use "Studio" for studio/efficiency units (typically smallest units, ~500-600 SF)
2. Whether the unit appears to be a **renovated** ("reno") variant
   - Reno units typically have higher rents for the same SF compared to non-reno units of the same bed/bath
   - Look for patterns: if two floorplans have similar SF but one has notably higher rents, the higher-rent one may be reno
   - Common reno naming: floorplan codes starting with different letters but same SF may indicate classic vs reno

Return ONLY valid JSON (no markdown fences) in this exact format:
{{
  "FLOORPLAN_CODE": {{"unit_type": "TYPE", "reno": true/false}},
  ...
}}

Order the entries by unit type (Studio first, then 1x1, 2x2, etc.), with non-reno before reno for each type.

IMPORTANT GUIDELINES:
- Every floorplan MUST appear in your output
- Use consistent naming: "Studio", "1x1", "2x2", "2x2 Den", "2x3 Loft", "3x2", "3x2 Loft", etc.
- "Reno" designation should NOT be in the unit_type string - it goes in the "reno" boolean field
- Focus on the SF and rent patterns to determine bed/bath counts:
  - Studio: ~400-650 SF
  - 1x1: ~600-1000 SF
  - 2x2: ~900-1300 SF
  - 2x2 Den / 2x3 Loft: ~1300-1600 SF
  - 3x2: ~1600-2000 SF
  - 3x2 Loft: ~1800-2100+ SF
- These are rough guides; use the data patterns to make your best judgment"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 3}],
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in response.content if hasattr(b, "text")).strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    # Extract JSON object from surrounding commentary (web search may add extra text)
    match = re.search(r"\{", text)
    if match:
        text = text[match.start():]
        depth = 0
        end_found = False
        for i, ch in enumerate(text):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    text = text[: i + 1]
                    end_found = True
                    break
        # If response was truncated (no balanced closing brace), repair it
        if not end_found:
            text = _repair_truncated_json(text)

    # Fix common AI JSON issues: trailing commas before } or ]
    text = re.sub(r",\s*([}\]])", r"\1", text)

    result = json.loads(text)

    # Fill missing floorplans with SF-based fallback
    for fp_summary in summary:
        fp = fp_summary["floorplan"]
        if fp not in result:
            result[fp] = {
                "unit_type": _fallback_unit_type(fp_summary.get("avg_sqft")),
                "reno": False,
            }

    return result


def mapping_to_df(mapping: dict[str, dict]) -> pd.DataFrame:
    """Convert the mapping dict to a DataFrame for display/editing in Streamlit."""
    rows = []
    for fp, info in mapping.items():
        rows.append({
            "Floorplan": fp,
            "Unit Type": info["unit_type"],
            "Reno": info.get("reno", False),
        })
    return pd.DataFrame(rows)


def df_to_mapping(df: pd.DataFrame) -> dict[str, dict]:
    """Convert an edited DataFrame back to a mapping dict."""
    mapping = {}
    for _, row in df.iterrows():
        mapping[row["Floorplan"]] = {
            "unit_type": row["Unit Type"],
            "reno": bool(row["Reno"]),
        }
    return mapping
