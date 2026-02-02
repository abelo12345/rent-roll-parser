"""T12 Mapper — AI-powered GL account → underwriting category mapping."""

import json
import re

import pandas as pd

from src.t12_parser import T12LineItem

# ---------------------------------------------------------------------------
# Fixed underwriting category list
# ---------------------------------------------------------------------------

REVENUE_WATERFALL = [
    "Potential Rent",
    "Loss to Lease",
    "Vacancy Loss",
    "Concessions",
    "Bad Debt",
    "Rental Revenue",
]

OTHER_INCOME = [
    "Parking Income",
    "Retail Income",
    "Retail CAM",
    "Signage",
    "Pet Rent",
    "RUBS",
    "Storage",
    "Other Income",
    "Miscellaneous Income",
]

CONTROLLABLE_EXPENSES = [
    "Payroll",
    "Repairs & Maintenance",
    "Turnover / Make Ready",
    "Contract Services",
    "Utilities",
    "Parking Expenses",
    "Trash",
    "Administrative",
    "Legal Fees",
    "Leasing & Marketing",
    "Ground Lease",
    "Signage Expense",
    "Retail CAM Expenses",
]

UNCONTROLLABLE_EXPENSES = [
    "Insurance",
    "Real Estate Taxes",
    "Management Fee",
]

BELOW_THE_LINE = [
    "Capital Reserve Above the Line",
    "Asset Management Fee",
    "Capital Reserve Below the Line",
    "Earthquake Ins Below the Line",
    "Debt Service",
    "Depreciation & Amortization",
    "Other Ownership Expenses",
]

ALL_CATEGORIES = (
    REVENUE_WATERFALL
    + OTHER_INCOME
    + CONTROLLABLE_EXPENSES
    + UNCONTROLLABLE_EXPENSES
    + BELOW_THE_LINE
)

_CATEGORY_SET = set(ALL_CATEGORIES)

_REVENUE_SET = set(REVENUE_WATERFALL + OTHER_INCOME)
_EXPENSE_SET = set(CONTROLLABLE_EXPENSES + UNCONTROLLABLE_EXPENSES)
_BTL_SET = set(BELOW_THE_LINE)


def _format_gl_items(leaf_items: list[T12LineItem]) -> str:
    """Format GL items grouped by section for the AI prompt."""
    lines = []
    current_section = None
    for item in leaf_items:
        section_label = " > ".join(item.section_path) if item.section_path else "(Top Level)"
        if section_label != current_section:
            current_section = section_label
            lines.append(f"\nSection: {section_label}")
        total_str = f"${item.total_value:,.0f}" if item.total_value is not None else "N/A"
        code_str = item.gl_code or "—"
        lines.append(
            f'  Row {item.source_row} | Code: {code_str} | "{item.gl_description}" | Total: {total_str}'
        )
    return "\n".join(lines)


def propose_t12_mapping(client, leaf_items: list[T12LineItem], grand_totals: dict | None = None) -> list[dict]:
    """Ask Claude to map each leaf GL item to a standard underwriting category.

    Args:
        grand_totals: dict from T12ParseResult.grand_totals, used for
            position-based validation (revenue vs expense vs BTL).

    Returns list of dicts with keys:
        source_row, gl_code, gl_description, category, confidence, notes
    """
    gl_text = _format_gl_items(leaf_items)

    category_list = "\n".join(f"- {c}" for c in ALL_CATEGORIES)

    prompt = f"""You are mapping GL (General Ledger) line items from a multifamily property's
Trailing 12-Month Income Statement to standardized underwriting categories.

Here are the GL line items to map. They are in order of appearance, grouped by section:

{gl_text}

Map each GL line item to EXACTLY ONE of these categories:

REVENUE WATERFALL:
- Potential Rent: Market rent, scheduled rent, gross market rent, gross potential rent
- Loss to Lease: Gain/loss to lease, loss to old lease
- Vacancy Loss: Physical vacancy, vacancy loss
- Concessions: All concession types (tenant, move-in, employee, model, construction, preferential rents)
- Bad Debt: Bad debt, write-offs, delinquent rent, prior resident collections, increase/decrease delinquent
- Rental Revenue: ONLY map items here if they are explicitly labeled as net rental revenue or similar calculated revenue items that don't fit above. Most revenue items should go to a more specific category.

OTHER INCOME:
- Parking Income: Parking fees, parking rent (residential parking)
- Retail Income: Base rent retail, commercial rent, retail revenue, signage rent
- Retail CAM: CAM reimbursements, expense recovery, reimb-property taxes, reimb-CAM, prior yr CAM
- Signage: Signage income (if separate from retail rent)
- Pet Rent: Pet rent, pet fees (recurring monthly)
- RUBS: Utility reimbursement/RUBS, water/sewer reimbursement
- Storage: Storage fees/rent (residential)
- Other Income: Lease processing fees, late fees, NSF charges, lease cancellation fees, transfer fees, short-term lease, furniture/corporate units, cable TV, damage/cleaning fees, month-to-month premiums, legal fee income, renters insurance proceeds, credit card processing fees, gate/access fees, returned check fees, rec room fees
- Miscellaneous Income: Interest income, other misc revenue — anything that does not clearly fit above

CONTROLLABLE EXPENSES (typically negative values):
- Payroll: Salaries, wages, overtime, payroll taxes, benefits, insurance, 401k, workers comp, employee housing, training, reimbursements, employee relations, commissions, bonuses, temp agency
- Repairs & Maintenance: All R&M lines (electrical, plumbing, HVAC, appliance, elevator, pest control, pools, painting, fire safety, common area, major repairs, fences, gates, sidewalks, boilers, water heaters, water leak repairs, countertop/cabinet repairs, locks/keys, windows/screens, concrete, landscaping)
- Turnover / Make Ready: Unit preparation, cleaning, carpet, vinyl, paint, drywall, turnover reimbursements, make-ready expenses, carpet cleaning/dye/repair, keys (unit prep context), unit paint, windows/screens damages (unit prep context)
- Contract Services: Maintenance service contracts, landscape contracts, patrol/security service
- Utilities: Electricity, gas, water, sewer, trash, cable TV expense, internet/streaming, resident reimbursements (negative utility offsets), other utility charges, vacant/model electricity
- Parking Expenses: Parking operations expense, parking payroll, parking maintenance
- Trash: ONLY if trash is broken out separately from the utilities section. If trash appears within a "Utilities" section, map to Utilities instead.
- Administrative: Office supplies, equipment, postage, telephone, bank fees, uniforms, credit inquiries (expense side), office maintenance
- Legal Fees: Legal fees (expense side), process service, other professional fees, management fees that are clearly professional service fees
- Leasing & Marketing: Advertising (print, internet), promotions, banners/signage/flags, resident relations, referral fees, resident/website referral, newsletters, special events
- Ground Lease: Ground lease expense
- Signage Expense: Signage expense (expense side)
- Retail CAM Expenses: Retail/commercial CAM expenses

UNCONTROLLABLE EXPENSES:
- Insurance: Property insurance, liability insurance, other insurance (NOT earthquake insurance)
- Real Estate Taxes: Property taxes, unsecured/personal property taxes, licenses and fees
- Management Fee: Management fees (the recurring % based fee, NOT one-time professional fees)

BELOW THE LINE:
- Capital Reserve Above the Line: Capital reserves treated as operating expense
- Asset Management Fee: Asset management fees
- Capital Reserve Below the Line: Capital reserves, capital expenditures, property improvements (building, HVAC, plumbing, mechanical, landscaping, fencing, flooring, pool, etc.)
- Earthquake Ins Below the Line: Earthquake insurance
- Debt Service: Mortgage payments, 1st mortgage, interest expense, loan interest, other debt service
- Depreciation & Amortization: Depreciation (building, furniture, equipment), amortization (loan fees, leasing costs, other)
- Other Ownership Expenses: Appraisal costs, audit/tax costs, bank fees, environmental, legal costs (ownership level), leasing costs (ownership level), permits/licenses, state income tax, insurance claims/expense (ownership level), other ownership expense

IMPORTANT RULES:
1. Every GL line MUST be mapped to exactly one category
2. When uncertain, prefer the more specific category over "Other Income" or "Miscellaneous Income"
3. Adjacent GL lines with similar codes/names likely belong to the same category
4. Revenue items should map to revenue/income categories; expense items to expense categories
5. Negative values in revenue sections can still be revenue (e.g., vacancy loss is negative)
6. Add a note when the mapping is uncertain or when the GL item could fit multiple categories

Return ONLY valid JSON (no markdown fences) as a list:
[
  {{"source_row": <row>, "gl_code": "<code>", "category": "<exact category name>",
    "confidence": "high"|"medium"|"low", "notes": "<explanation if uncertain>"}},
  ...
]"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8192,
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 3}],
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in response.content if hasattr(b, "text")).strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    # Extract JSON array
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

    # Fix common AI JSON issues: trailing commas before } or ]
    text = re.sub(r",\s*([}\]])", r"\1", text)

    ai_mapping = json.loads(text)

    # Build lookup from source_row
    row_lookup = {item.source_row: item for item in leaf_items}

    # Validate and fill gaps
    mapped_rows = {m["source_row"] for m in ai_mapping}
    result = []

    for m in ai_mapping:
        cat = m.get("category", "")
        if cat not in _CATEGORY_SET:
            m["notes"] = f"AI returned unknown category '{cat}'; mapped to Miscellaneous Income"
            m["category"] = "Miscellaneous Income"
            m["confidence"] = "low"
        # Enrich with GL description
        item = row_lookup.get(m["source_row"])
        if item:
            m["gl_description"] = item.gl_description
        result.append(m)

    # Fill unmapped items
    for item in leaf_items:
        if item.source_row not in mapped_rows:
            result.append({
                "source_row": item.source_row,
                "gl_code": item.gl_code,
                "gl_description": item.gl_description,
                "category": "Miscellaneous Income",
                "confidence": "low",
                "notes": "Not mapped by AI; defaulted to Miscellaneous Income",
            })

    # Position-based validation: enforce that items in the revenue section
    # map to revenue categories, expense section to expense, and BTL to BTL.
    if grand_totals:
        income_row = grand_totals["total_income"].source_row if grand_totals.get("total_income") else None
        opex_row = grand_totals["total_opex"].source_row if grand_totals.get("total_opex") else None
        noi_row = grand_totals["noi"].source_row if grand_totals.get("noi") else None

        for m in result:
            cat = m["category"]
            row = m["source_row"]

            if income_row and row < income_row and cat not in _REVENUE_SET:
                old = cat
                m["category"] = "Miscellaneous Income"
                m["confidence"] = "low"
                m["notes"] = f"Corrected: '{old}' is not a revenue category (item is in revenue section)"

            elif income_row and opex_row and income_row < row < opex_row and cat not in _EXPENSE_SET:
                old = cat
                m["category"] = "Administrative"
                m["confidence"] = "low"
                m["notes"] = f"Corrected: '{old}' is not an expense category (item is in expense section)"

            elif noi_row and row > noi_row and cat not in _BTL_SET:
                old = cat
                m["category"] = "Capital Reserve Below the Line"
                m["confidence"] = "low"
                m["notes"] = f"Corrected: '{old}' is not a BTL category (item is below NOI)"

    # Sort by source_row
    result.sort(key=lambda m: m["source_row"])
    return result


def mapping_to_df(mapping: list[dict]) -> pd.DataFrame:
    """Convert mapping list to an editable DataFrame."""
    rows = []
    for m in mapping:
        rows.append({
            "Source Row": m["source_row"],
            "GL Code": m.get("gl_code", ""),
            "GL Description": m.get("gl_description", ""),
            "Category": m["category"],
            "Confidence": m.get("confidence", ""),
            "Notes": m.get("notes", ""),
        })
    return pd.DataFrame(rows)


def df_to_mapping(df: pd.DataFrame) -> list[dict]:
    """Convert edited DataFrame back to mapping list."""
    result = []
    for _, row in df.iterrows():
        result.append({
            "source_row": int(row["Source Row"]),
            "gl_code": row.get("GL Code", ""),
            "gl_description": row.get("GL Description", ""),
            "category": row["Category"],
            "confidence": row.get("Confidence", ""),
            "notes": row.get("Notes", ""),
        })
    return result
