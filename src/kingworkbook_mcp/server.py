"""MCP server entry point and four read-only workbook tools."""

from __future__ import annotations

import logging
from typing import Annotated, Any, Callable, TypeVar

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from .models import CompareResult, FormulaCheckResult, InspectResult, ReconcileResult
from .workbooks import (
    WorkbookInputError,
    check_formulas,
    compare_tables,
    inspect_file,
    reconcile_column,
)

mcp = MCPServer("KingWorkbook MCP")
T = TypeVar("T")
READ_ONLY = ToolAnnotations(
    read_only_hint=True,
    destructive_hint=False,
    idempotent_hint=True,
    open_world_hint=False,
)


def _translate_input_error(action: Callable[..., T], *args: Any, **kwargs: Any) -> T:
    try:
        return action(*args, **kwargs)
    except WorkbookInputError as exc:
        raise ToolError(str(exc)) from exc


@mcp.tool(
    title="Inspect workbook",
    description=(
        "Read-only profile of a CSV or XLSX file inside KINGWORKBOOK_ROOT. Returns sheet names, "
        "headers, row counts, blank counts, inferred types, and formula counts without cell values."
    ),
    annotations=READ_ONLY,
)
def inspect_workbook(
    file_path: Annotated[str, Field(description="Path relative to KINGWORKBOOK_ROOT, such as reports/monthly.xlsx.")],
) -> InspectResult:
    """Inspect workbook structure and bounded column profiles without returning ordinary cell contents."""
    return _translate_input_error(inspect_file, file_path)


@mcp.tool(
    title="Compare workbooks",
    description=(
        "Compare two CSV/XLSX tables using explicit key columns. Duplicate or blank keys are surfaced; "
        "keys are never returned. Changed cell values are excluded unless include_values is true."
    ),
    annotations=READ_ONLY,
)
def compare_workbooks(
    before_path: Annotated[str, Field(description="Earlier CSV/XLSX path relative to KINGWORKBOOK_ROOT.")],
    after_path: Annotated[str, Field(description="Later CSV/XLSX path relative to KINGWORKBOOK_ROOT.")],
    key_columns: Annotated[list[str], Field(description="One or more exact normalized column names from inspect_workbook.")],
    before_sheet: Annotated[str | None, Field(description="Optional sheet name in the earlier XLSX file.")] = None,
    after_sheet: Annotated[str | None, Field(description="Optional sheet name in the later XLSX file.")] = None,
    include_values: Annotated[bool, Field(description="Set true only when the user wants changed cell value previews.")] = False,
    max_reported_changes: Annotated[int, Field(ge=1, le=100, description="Maximum changed-record samples to return.")] = 30,
) -> CompareResult:
    """Compare keyed records and report changes with row references, omitting values by default."""
    return _translate_input_error(
        compare_tables,
        before_path,
        after_path,
        key_columns,
        before_sheet=before_sheet,
        after_sheet=after_sheet,
        include_values=include_values,
        max_reported_changes=max_reported_changes,
    )


@mcp.tool(
    title="Reconcile a total",
    description=(
        "Sum one selected numeric column and compare it with a supplied control total and tolerance. "
        "Formulas are not calculated; usable cached results are labeled and counted."
    ),
    annotations=READ_ONLY,
)
def reconcile_totals(
    file_path: Annotated[str, Field(description="CSV/XLSX path relative to KINGWORKBOOK_ROOT.")],
    column: Annotated[str, Field(description="Exact normalized column name from inspect_workbook.")],
    expected_total: Annotated[str, Field(description="Expected total as a decimal string, for example 1250.00.")],
    tolerance: Annotated[str, Field(description="Allowed absolute difference as a decimal string; defaults to zero.")] = "0",
    sheet_name: Annotated[str | None, Field(description="Optional sheet name in an XLSX workbook.")] = None,
) -> ReconcileResult:
    """Reconcile numeric values without executing formulas or changing the source file."""
    return _translate_input_error(
        reconcile_column,
        file_path,
        column,
        expected_total,
        tolerance,
        sheet_name,
    )


@mcp.tool(
    title="Check formula integrity",
    description=(
        "Review formula cells in an XLSX sheet for cached errors, #REF! text, and unusual relative-pattern "
        "variations. It never evaluates formulas and presents variations as human review leads."
    ),
    annotations=READ_ONLY,
)
def check_formula_integrity(
    file_path: Annotated[str, Field(description="XLSX path relative to KINGWORKBOOK_ROOT.")],
    sheet_name: Annotated[str | None, Field(description="Optional sheet name; defaults to the first visible sheet.")] = None,
    max_findings: Annotated[int, Field(ge=1, le=100, description="Maximum total finding cells returned.")] = 100,
) -> FormulaCheckResult:
    """Inspect cached formula results and relative formula patterns without recalculation."""
    return _translate_input_error(
        check_formulas,
        file_path,
        sheet_name,
        max_findings,
    )


def main() -> None:
    logging.basicConfig(level=logging.WARNING)
    mcp.run()
