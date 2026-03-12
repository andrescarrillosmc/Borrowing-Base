from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class SheetSummary:
    name: str
    max_row: int
    max_column: int
    nonempty_cells: int
    formula_cells: int


@dataclass
class InputField:
    input_id: str
    name: str
    prompt: str
    description: str
    data_type: str
    required: bool
    allowed_values: list[str]
    default_rule: str
    validation_rule: str
    destination_sheet: str
    destination_cell: str
    write_method: str
    may_infer: bool
    human_confirm: bool
    affects_eligibility: str
    affects_concentration: str
    affects_availability: str
    run_types: str
    notes: str


@dataclass
class OutputField:
    output_id: str
    name: str
    business_meaning: str
    source_sheet: str
    source_cell: str
    output_type: str
    required_in_response: bool
    before_after: str
    format_rule: str
    capture_method: str
    notes: str


@dataclass
class ValidationRule:
    rule_id: str
    field_ref: str
    rule_type: str
    logic: str
    error_message: str
    severity: str
    auto_correct: bool
    escalate: bool
    notes: str


@dataclass
class WritePolicy:
    sheet_name: str
    permission: str
    allowed_range: str
    purpose: str
    conditions: str
    risk_level: str
    notes: str


@dataclass
class CapacityDiagnostic:
    area: str
    current_last_row: int
    modeled_ceiling: int | None
    remaining_headroom: int | None
    note: str


@dataclass
class WorkbookDiagnosis:
    workbook_path: str
    sheet_summaries: list[SheetSummary]
    inputs: list[InputField]
    outputs: list[OutputField]
    validation_rules: list[ValidationRule]
    write_policies: list[WritePolicy]
    capacities: list[CapacityDiagnostic]
    industries: list[str]
    monthly_update_tabs: list[str]
    current_risks: list[str]
    recommended_architecture: list[str]
    workbook_version: str
    measurement_date: str
    borrower_name: str
    facility_amount_formula: str
    debug: dict[str, Any] = field(default_factory=dict)

    def summary_lines(self) -> list[str]:
        return [
            f"Workbook: {self.workbook_path}",
            f"Version: {self.workbook_version or 'Unknown'}",
            f"Measurement date: {self.measurement_date or 'Unknown'}",
            f"Borrower: {self.borrower_name or 'Unknown'}",
            f"Inputs mapped: {len(self.inputs)}",
            f"Outputs mapped: {len(self.outputs)}",
            f"Validation rules: {len(self.validation_rules)}",
            f"Write policies: {len(self.write_policies)}",
        ]
