from pathlib import Path
import unittest
import sys

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from borrowing_base_workbench.analysis import analyze_workbook
from borrowing_base_workbench.validation import validate_scenario


WORKBOOK = Path(r"C:\Users\Andres.Carrillo\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")


@unittest.skipUnless(WORKBOOK.exists(), "Borrowing base workbook not available")
class AnalysisTests(unittest.TestCase):
    def test_analysis_finds_mapping_layer(self) -> None:
        diagnosis = analyze_workbook(WORKBOOK)
        self.assertGreaterEqual(len(diagnosis.inputs), 20)
        self.assertGreaterEqual(len(diagnosis.outputs), 20)
        self.assertIn("Loan Tape - Settled", diagnosis.monthly_update_tabs)

    def test_validation_catches_bad_inputs(self) -> None:
        diagnosis = analyze_workbook(WORKBOOK)
        issues = validate_scenario(
            {
                "company_name": "BadCo",
                "security_type": "Second Lien",
                "industry_classification": "Not A Real Industry",
                "ltm_revenue": "1000000",
                "ltm_adj_ebitda": "5000000",
                "bdc_balance": "10000000",
                "total_sm_balance": "9000000",
                "loan_denomination": "CAD",
                "country": "Canada",
                "purchase_price": "0.8",
                "investment_date": "2026-03-10",
                "maturity_date": "2035-03-10",
            },
            diagnosis,
        )
        messages = [issue.message for issue in issues]
        self.assertTrue(any("Total SM Balance" in message for message in messages))
        self.assertTrue(any("Industry" in message for message in messages))


if __name__ == "__main__":
    unittest.main()
