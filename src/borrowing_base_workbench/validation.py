from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from .models import WorkbookDiagnosis


@dataclass
class ValidationIssue:
    severity: str
    field: str
    message: str


def _to_float(value: str) -> float | None:
    if value in (None, ""):
        return None
    try:
        text = str(value).strip()
        had_percent = "%" in text
        cleaned = text.replace("$", "").replace(",", "").replace("x", "").replace("X", "").replace("%", "").strip()
        if not cleaned:
            return None
        number = float(cleaned)
        if had_percent:
            return number / 100.0
        return number
    except ValueError:
        return None


def _to_date(value: str) -> date | None:
    if not value:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def validate_scenario(values: dict[str, str], diagnosis: WorkbookDiagnosis) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []

    required = {
        "company_name": "Company Name",
        "security_type": "Security Type",
        "industry_classification": "Industry Classification",
        "ltm_adj_ebitda": "LTM Adj. EBITDA",
        "bdc_balance": "BDC Balance",
        "investment_date": "Investment Date",
        "maturity_date": "Maturity Date",
    }
    for key, label in required.items():
        if not values.get(key, "").strip():
            issues.append(ValidationIssue("Hard Stop", label, f"{label} is required."))

    revenue = _to_float(values.get("ltm_revenue", ""))
    ebitda = _to_float(values.get("ltm_adj_ebitda", ""))
    bdc_balance = _to_float(values.get("bdc_balance", ""))
    total_sm = _to_float(values.get("total_sm_balance", ""))
    purchase_price = _to_float(values.get("purchase_price", ""))
    spread = _to_float(values.get("spread", ""))
    denom = values.get("loan_denomination", "").strip()
    country = values.get("country", "").strip()
    security_type = values.get("security_type", "").strip()
    industry = values.get("industry_classification", "").strip()
    pik = _to_float(values.get("pik_pct", ""))
    invest_date = _to_date(values.get("investment_date", ""))
    maturity_date = _to_date(values.get("maturity_date", ""))
    revolver = _to_float(values.get("drawn_revolver", "")) or 0.0
    first_out = _to_float(values.get("first_out_balance", "")) or 0.0
    pari_passu = _to_float(values.get("pari_passu", "")) or 0.0
    cash = _to_float(values.get("cash_balance", "")) or 0.0

    if revenue is not None and revenue <= 0:
        issues.append(ValidationIssue("Hard Stop", "LTM Revenue", "Revenue must be positive."))
    if ebitda is not None and ebitda <= 0:
        issues.append(ValidationIssue("Hard Stop", "LTM Adj. EBITDA", "EBITDA must be positive."))
    if bdc_balance is not None and bdc_balance <= 0:
        issues.append(ValidationIssue("Hard Stop", "BDC Balance", "BDC Balance must be positive."))
    if total_sm is not None and bdc_balance is not None and total_sm < bdc_balance:
        issues.append(ValidationIssue("Hard Stop", "Total SM Balance", "Total SM Balance must be at least the BDC Balance."))
    if purchase_price is not None and not (0.01 <= purchase_price <= 1.50):
        issues.append(ValidationIssue("Warning", "Purchase Price", "Purchase price sits outside the modeled range of 0.01 to 1.50."))
    if purchase_price is not None and purchase_price < 0.90:
        issues.append(ValidationIssue("Warning", "Purchase Price", "Purchase price is below the AGENT minimum for eligibility."))
    if spread is not None and not (0.01 <= spread <= 0.25):
        issues.append(ValidationIssue("Warning", "Spread", "Spread sits outside the normal modeled range of 1% to 25%."))
    if security_type == "Second Lien" and ebitda is not None and ebitda < 10_000_000:
        issues.append(ValidationIssue("Warning", "LTM Adj. EBITDA", "Second Lien deals require at least $10MM EBITDA for eligibility."))
    if denom and denom != "USD":
        issues.append(ValidationIssue("Warning", "Loan Denomination", "Non-USD denomination will fail the current eligibility rule set."))
    if country and country != "United States":
        issues.append(ValidationIssue("Warning", "Country", "Non-US obligors will fail the current eligibility rule set."))
    if pik is not None and pik > 0:
        issues.append(ValidationIssue("Warning", "PIK %", "PIK requires manual review because it can fail the non-PIK / DIP eligibility test."))
    if values.get("default_status", "").strip().lower() == "yes":
        issues.append(ValidationIssue("Warning", "Default Status", "Defaulted loans should not auto-run without manual review."))
    if industry and industry not in diagnosis.industries:
        issues.append(ValidationIssue("Hard Stop", "Industry Classification", "Industry must match the approved Concentration Limits taxonomy exactly."))
    if invest_date and maturity_date:
        if maturity_date <= invest_date:
            issues.append(ValidationIssue("Hard Stop", "Maturity Date", "Maturity must be after Investment Date."))
        term_years = (maturity_date - invest_date).days / 365.0
        if term_years > 7:
            issues.append(ValidationIssue("Warning", "Maturity Date", "Original term is longer than the 7-year AGENT threshold."))

    if ebitda:
        leverage = (revolver + first_out + pari_passu + (total_sm or 0.0) - cash) / ebitda
        if leverage > 6.5:
            issues.append(ValidationIssue("Warning", "Leverage", f"Pro forma leverage is {leverage:.2f}x, above the 6.5x threshold."))

    return issues


def build_commentary(values: dict[str, str], issues: list[ValidationIssue]) -> str:
    company = values.get("company_name", "").strip() or "This scenario"
    hard_stops = [issue for issue in issues if issue.severity == "Hard Stop"]
    warnings = [issue for issue in issues if issue.severity != "Hard Stop"]

    lines = [f"{company} scenario summary:"]
    if hard_stops:
        lines.append("The scenario is not ready for a governed run because one or more hard-stop checks failed.")
    else:
        lines.append("The scenario clears the base input gate and is ready for workbook-level execution once the Excel runner uses a trusted workbook copy.")

    if warnings:
        lines.append("Key warning flags:")
        for issue in warnings[:5]:
            lines.append(f"- {issue.field}: {issue.message}")
    else:
        lines.append("No warning-level issues were raised by the current pre-flight rules.")

    return "\n".join(lines)
