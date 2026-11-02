"""Structured MCP tool result models."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class HeaderInfo(BaseModel):
    name: str
    source_header: str | None
    index: int
    blank: bool
    duplicate: bool


class ColumnProfile(HeaderInfo):
    blank_cells: int
    type_counts: dict[str, int]
    formula_count: int


class SheetProfile(BaseModel):
    sheet_name: str
    state: Literal["visible", "hidden", "veryHidden"]
    data_rows: int
    blank_rows: int
    scanned_cells: int
    columns: list[ColumnProfile]


class InspectResult(BaseModel):
    file_path: str
    file_type: Literal["csv", "xlsx"]
    sheets: list[SheetProfile]
    omitted_sheet_count: int
    limits: list[str]
    notes: list[str]


class ChangedRecord(BaseModel):
    before_row: int | None
    after_row: int | None
    changed_columns: list[str] = Field(default_factory=list)
    value_previews: dict[str, dict[str, str | None]] = Field(default_factory=dict)
    value_previews_truncated: bool = False


class RecordSample(BaseModel):
    row_number: int
    value_previews: dict[str, str | None] = Field(default_factory=dict)
    value_previews_truncated: bool = False


class CompareResult(BaseModel):
    status: Literal["complete", "needs_review"]
    before_path: str
    after_path: str
    before_sheet: str
    after_sheet: str
    key_columns: list[str]
    before_records: int
    after_records: int
    unchanged_records: int | None
    added_records: int | None
    removed_records: int | None
    changed_records: int | None
    changed_cells: int | None
    missing_key_rows_before: int
    missing_key_rows_after: int
    duplicate_key_rows_before: list[int]
    duplicate_key_rows_after: list[int]
    duplicate_key_count_before: int
    duplicate_key_count_after: int
    columns_only_before: list[str]
    columns_only_after: list[str]
    changes: list[ChangedRecord]
    changes_truncated: bool
    added_rows: list[RecordSample]
    removed_rows: list[RecordSample]
    added_rows_truncated: bool
    removed_rows_truncated: bool
    values_included: bool
    notes: list[str]


class ReconcileResult(BaseModel):
    status: Literal["matched", "outside_tolerance", "needs_review"]
    file_path: str
    sheet_name: str
    column: str
    data_rows: int
    numeric_values_included: int
    blank_values_skipped: int
    nonnumeric_values_skipped: int
    formulas_without_cached_values: int
    expected_total: str
    computed_total: str
    difference: str
    tolerance: str
    within_tolerance: bool
    notes: list[str]


class FormulaFinding(BaseModel):
    code: Literal["cached_formula_error", "broken_reference", "formula_pattern_variation"]
    severity: Literal["review", "error"]
    message: str
    cells: list[str]


class FormulaCheckResult(BaseModel):
    file_path: str
    sheet_name: str
    formulas_scanned: int
    formula_runs_compared: int
    formula_pattern_variations: int
    cached_formula_errors: int
    broken_reference_formulas: int
    findings: list[FormulaFinding]
    findings_truncated: bool
    notes: list[str]
