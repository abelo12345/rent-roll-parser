"""T12 post-mapping validation: reconciliation checks + AI review + self-correction."""

import json
import re

from src.t12_mapper import (
    REVENUE_WATERFALL, OTHER_INCOME,
    CONTROLLABLE_EXPENSES, UNCONTROLLABLE_EXPENSES,
    BELOW_THE_LINE, ALL_CATEGORIES,
    _CATEGORY_SET,
)
from src.t12_parser import T12ParseResult


# ---------------------------------------------------------------------------
# Compute T12 totals from mapping (mirrors T12 Summary sheet logic)
# ---------------------------------------------------------------------------

_REVENUE_CATS = set(REVENUE_WATERFALL + OTHER_INCOME)
_EXPENSE_CATS = set(CONTROLLABLE_EXPENSES + UNCONTROLLABLE_EXPENSES)
_BTL_CATS = set(BELOW_THE_LINE)


def _compute_t12_totals(mapping: list[dict], leaf_items: list, expenses_positive: bool = True) -> dict:
    """Compute T12 summary totals from the mapping.

    Returns dict with:
      - egi: Effective Gross Income (net rental revenue + other income)
      - total_opex: Total Operating Expenses
      - noi: Net Operating Income
      - cash_flow: Cash Flow After BTL
      - by_category: {category_name: total_value}
      - net_rental_revenue, total_other_income, total_controllable, total_uncontrollable, total_btl
    """
    # Build total_value lookup from leaf items
    leaf_totals = {item.source_row: item.total_value for item in leaf_items}

    # Sum by category
    by_category = {}
    for m in mapping:
        cat = m["category"]
        src_row = m["source_row"]
        val = leaf_totals.get(src_row, 0) or 0
        by_category[cat] = by_category.get(cat, 0) + val

    # Compute group totals
    net_rental_revenue = sum(by_category.get(c, 0) for c in REVENUE_WATERFALL)
    total_other_income = sum(by_category.get(c, 0) for c in OTHER_INCOME)
    egi = net_rental_revenue + total_other_income

    total_controllable = sum(by_category.get(c, 0) for c in CONTROLLABLE_EXPENSES)
    total_uncontrollable = sum(by_category.get(c, 0) for c in UNCONTROLLABLE_EXPENSES)
    total_opex = total_controllable + total_uncontrollable

    # NOI = EGI - Opex (sign depends on whether expenses are positive or negative)
    if expenses_positive:
        noi = egi - total_opex
    else:
        noi = egi + total_opex

    total_btl = sum(by_category.get(c, 0) for c in BELOW_THE_LINE)

    if expenses_positive:
        cash_flow = noi - total_btl
    else:
        cash_flow = noi + total_btl

    return {
        "egi": egi,
        "total_opex": total_opex,
        "noi": noi,
        "cash_flow": cash_flow,
        "net_rental_revenue": net_rental_revenue,
        "total_other_income": total_other_income,
        "total_controllable": total_controllable,
        "total_uncontrollable": total_uncontrollable,
        "total_btl": total_btl,
        "by_category": by_category,
    }


# ---------------------------------------------------------------------------
# T12 Checks
# ---------------------------------------------------------------------------

def run_t12_checks(client, parse_result: T12ParseResult, mapping: list[dict]) -> list[dict]:
    """Run validation checks on T12 mapping by comparing computed totals to source grand totals.

    Returns list of issues: {"severity", "check", "detail", "delta"}.
    """
    gt = parse_result.grand_totals
    issues = []

    # Detect sign convention
    expenses_positive = True
    if gt.get("total_opex") and gt["total_opex"].total_value is not None:
        expenses_positive = gt["total_opex"].total_value > 0

    computed = _compute_t12_totals(mapping, parse_result.leaf_items, expenses_positive)

    # --- EGI vs source total income ---
    if gt.get("total_income") and gt["total_income"].total_value is not None:
        source_egi = gt["total_income"].total_value
        delta = computed["egi"] - source_egi
        if abs(delta) > 1:
            sev = "error" if abs(delta) > abs(source_egi) * 0.05 else "warning"
            issues.append({
                "severity": sev,
                "check": "T12: EGI Reconciliation",
                "detail": f"Our EGI: ${computed['egi']:,.0f}, Source: ${source_egi:,.0f}, Delta: ${delta:,.0f}",
                "delta": delta,
                "source": "t12_check",
            })

    # --- Opex vs source total opex ---
    if gt.get("total_opex") and gt["total_opex"].total_value is not None:
        source_opex = gt["total_opex"].total_value
        delta = computed["total_opex"] - source_opex
        if abs(delta) > 1:
            sev = "error" if abs(delta) > abs(source_opex) * 0.05 else "warning"
            issues.append({
                "severity": sev,
                "check": "T12: Opex Reconciliation",
                "detail": f"Our Opex: ${computed['total_opex']:,.0f}, Source: ${source_opex:,.0f}, Delta: ${delta:,.0f}",
                "delta": delta,
                "source": "t12_check",
            })

    # --- NOI vs source NOI ---
    if gt.get("noi") and gt["noi"].total_value is not None:
        source_noi = gt["noi"].total_value
        delta = computed["noi"] - source_noi
        if abs(delta) > 1:
            sev = "error" if abs(delta) > abs(source_noi) * 0.05 else "warning"
            issues.append({
                "severity": sev,
                "check": "T12: NOI Reconciliation",
                "detail": f"Our NOI: ${computed['noi']:,.0f}, Source: ${source_noi:,.0f}, Delta: ${delta:,.0f}",
                "delta": delta,
                "source": "t12_check",
            })

    # --- Cash flow vs source net income ---
    if gt.get("net_income") and gt["net_income"].total_value is not None:
        source_net = gt["net_income"].total_value
        delta = computed["cash_flow"] - source_net
        if abs(delta) > 1:
            sev = "error" if abs(delta) > abs(source_net) * 0.05 else "warning"
            issues.append({
                "severity": sev,
                "check": "T12: Cash Flow Reconciliation",
                "detail": f"Our Cash Flow: ${computed['cash_flow']:,.0f}, Source: ${source_net:,.0f}, Delta: ${delta:,.0f}",
                "delta": delta,
                "source": "t12_check",
            })

    # --- Low confidence items ---
    low_conf = [m for m in mapping if m.get("confidence") == "low"]
    if low_conf:
        issues.append({
            "severity": "warning",
            "check": "T12: Low Confidence Mappings",
            "detail": f"{len(low_conf)} items have low confidence mappings: "
                      + ", ".join(f"'{m.get('gl_description', '')}' → {m['category']}" for m in low_conf[:5]),
            "source": "t12_check",
        })

    # --- Unmapped / defaulted items ---
    defaulted = [m for m in mapping if "defaulted" in m.get("notes", "").lower() or "not mapped" in m.get("notes", "").lower()]
    if defaulted:
        issues.append({
            "severity": "warning",
            "check": "T12: Unmapped Items",
            "detail": f"{len(defaulted)} items were not mapped by AI and defaulted to Miscellaneous Income.",
            "source": "t12_check",
        })

    # --- AI review of the T12 summary ---
    summary_lines = [
        f"Net Rental Revenue: ${computed['net_rental_revenue']:,.0f}",
        f"Total Other Income: ${computed['total_other_income']:,.0f}",
        f"EGI: ${computed['egi']:,.0f}",
        f"Total Controllable: ${computed['total_controllable']:,.0f}",
        f"Total Uncontrollable: ${computed['total_uncontrollable']:,.0f}",
        f"Total Opex: ${computed['total_opex']:,.0f}",
        f"NOI: ${computed['noi']:,.0f}",
        f"Total BTL: ${computed['total_btl']:,.0f}",
        f"Cash Flow: ${computed['cash_flow']:,.0f}",
        "\nCategory breakdown:",
    ]
    for cat in ALL_CATEGORIES:
        val = computed["by_category"].get(cat, 0)
        if val != 0:
            summary_lines.append(f"  {cat}: ${val:,.0f}")

    existing_issues = "\n".join(
        f"- [{i['severity'].upper()}] {i['check']}: {i['detail']}"
        for i in issues
    ) if issues else "None."

    prompt = f"""You are reviewing a mapped T12 Income Statement for a multifamily property. Check for mapping errors or anomalies.

COMPUTED T12 SUMMARY:
{chr(10).join(summary_lines)}

ISSUES ALREADY DETECTED:
{existing_issues}

Look for ADDITIONAL issues:
1. Categories that seem too large or too small relative to the property
2. Revenue items that may have been mapped to expenses or vice versa
3. Missing expected categories (e.g., property with no insurance, no taxes, no payroll)
4. Any other mapping anomalies

Return ONLY valid JSON (no markdown fences) as a list:
[
  {{"severity": "warning", "check": "<check name>", "detail": "<explanation>"}},
  ...
]

If no additional issues: []"""

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
            issue["source"] = "t12_ai"
        issues.extend(ai_issues)
    except Exception:
        pass

    severity_order = {"error": 0, "warning": 1, "info": 2}
    issues.sort(key=lambda x: severity_order.get(x.get("severity", "info"), 3))
    return issues


# ---------------------------------------------------------------------------
# Self-correction — re-map items to close reconciliation gaps
# ---------------------------------------------------------------------------

def self_correct_t12(client, parse_result: T12ParseResult, mapping: list[dict], issues: list[dict]) -> list[dict] | None:
    """Attempt to fix T12 mapping issues by re-mapping suspect items.

    Returns corrected mapping list, or None if no improvement possible.
    """
    # Only attempt correction if there are reconciliation errors
    recon_errors = [i for i in issues if "Reconciliation" in i.get("check", "") and i.get("severity") == "error"]
    if not recon_errors:
        return None

    # Identify suspect items: low/medium confidence, or items near the delta
    suspects = []
    for m in mapping:
        conf = m.get("confidence", "high")
        if conf in ("low", "medium"):
            suspects.append(m)

    if not suspects:
        # No low-confidence items to re-map
        return None

    # Build context for AI
    leaf_lookup = {item.source_row: item for item in parse_result.leaf_items}
    suspect_lines = []
    for m in suspects:
        item = leaf_lookup.get(m["source_row"])
        total_str = f"${item.total_value:,.0f}" if item and item.total_value is not None else "N/A"
        section = " > ".join(item.section_path) if item and item.section_path else ""
        suspect_lines.append(
            f'  Row {m["source_row"]} | "{m.get("gl_description", "")}" | Section: {section} | '
            f'Total: {total_str} | Current: {m["category"]} (confidence: {m.get("confidence", "")})'
        )

    error_details = "\n".join(f"- {e['detail']}" for e in recon_errors)

    prompt = f"""The T12 mapping has reconciliation errors. Please re-map the suspect items below to fix the deltas.

RECONCILIATION ERRORS:
{error_details}

SUSPECT ITEMS (low/medium confidence, may be miscategorized):
{chr(10).join(suspect_lines)}

AVAILABLE CATEGORIES:
{chr(10).join(f"- {c}" for c in ALL_CATEGORIES)}

For each suspect item, decide if it should be re-mapped. Return ONLY valid JSON (no markdown fences):
[
  {{"source_row": <row>, "category": "<correct category>", "confidence": "high"|"medium", "notes": "<reason>"}},
  ...
]

Only include items that SHOULD change. If an item is already correct, omit it."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
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
        corrections = json.loads(text)
    except Exception:
        return None

    if not corrections:
        return None

    # Apply corrections to a copy of the mapping
    corrected = [dict(m) for m in mapping]
    row_to_idx = {m["source_row"]: i for i, m in enumerate(corrected)}

    changes_made = 0
    for fix in corrections:
        row = fix.get("source_row")
        new_cat = fix.get("category", "")
        if row in row_to_idx and new_cat in _CATEGORY_SET:
            idx = row_to_idx[row]
            if corrected[idx]["category"] != new_cat:
                corrected[idx]["category"] = new_cat
                corrected[idx]["confidence"] = fix.get("confidence", "medium")
                corrected[idx]["notes"] = f"Self-corrected: {fix.get('notes', '')}"
                changes_made += 1

    if changes_made == 0:
        return None

    # Verify the corrections actually improve things
    gt = parse_result.grand_totals
    expenses_positive = True
    if gt.get("total_opex") and gt["total_opex"].total_value is not None:
        expenses_positive = gt["total_opex"].total_value > 0

    old_totals = _compute_t12_totals(mapping, parse_result.leaf_items, expenses_positive)
    new_totals = _compute_t12_totals(corrected, parse_result.leaf_items, expenses_positive)

    # Check if the total absolute delta improved
    def _total_delta(totals):
        d = 0
        if gt.get("total_income") and gt["total_income"].total_value is not None:
            d += abs(totals["egi"] - gt["total_income"].total_value)
        if gt.get("total_opex") and gt["total_opex"].total_value is not None:
            d += abs(totals["total_opex"] - gt["total_opex"].total_value)
        if gt.get("noi") and gt["noi"].total_value is not None:
            d += abs(totals["noi"] - gt["noi"].total_value)
        return d

    old_delta = _total_delta(old_totals)
    new_delta = _total_delta(new_totals)

    if new_delta < old_delta:
        return corrected
    else:
        # Corrections didn't help, discard
        return None
