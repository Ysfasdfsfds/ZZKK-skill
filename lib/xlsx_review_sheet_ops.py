from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl.cell.cell import MergedCell
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.cell import range_boundaries

from lib.review_logic import ReviewPlan


_WRAP_TOP = Alignment(wrap_text=True, vertical='top')


def render_customer_review_sheet(
    template_path: Path,
    output_path: Path,
    *,
    record_number: str,
    project_name: str,
    work_product_name: str,
    work_product_review_scope: str,
    meeting_date: str,
    host: str,
    scribe: str,
    reviewers: str,
    other_people: str,
    review_plan: ReviewPlan,
    revision_reference: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = work_product_name
    sheet.cell(row=5, column=5).value = work_product_review_scope
    sheet.cell(row=5, column=11).value = '设计阶段'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 1.5
    sheet.cell(row=7, column=5).value = host
    sheet.cell(row=7, column=9).value = scribe
    sheet.cell(row=8, column=5).value = reviewers
    sheet.cell(row=9, column=5).value = other_people
    _write_review_plan(sheet, review_plan, owner=host, completion_date=meeting_date, revision_reference=revision_reference)
    workbook.save(output_path)


def render_product_review_sheet(
    template_path: Path,
    output_path: Path,
    *,
    record_number: str,
    project_name: str,
    work_product_name: str,
    work_product_review_scope: str,
    meeting_date: str,
    host: str,
    scribe: str,
    reviewers: str,
    other_people: str,
    review_plan: ReviewPlan,
    revision_reference: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = work_product_name
    sheet.cell(row=5, column=5).value = work_product_review_scope
    sheet.cell(row=5, column=11).value = '设计阶段'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 2
    sheet.cell(row=7, column=5).value = host
    sheet.cell(row=7, column=9).value = scribe
    sheet.cell(row=8, column=5).value = reviewers
    sheet.cell(row=9, column=5).value = other_people
    _write_review_plan(sheet, review_plan, owner=host, completion_date=meeting_date, revision_reference=revision_reference)
    workbook.save(output_path)


def _write_review_plan(sheet, review_plan: ReviewPlan, *, owner: str, completion_date: str, revision_reference: str) -> None:
    _normalize_issue_rows(sheet)
    _copy_row_style(sheet, source_row=13, target_rows=[14, 15], min_col=2, max_col=13)
    _clear_issue_rows(sheet)
    issue_count = len(review_plan.issues)
    for offset, issue in enumerate(review_plan.issues):
        row = 13 + offset
        values = [
            issue.sequence,
            '1',
            issue.page_scope,
            issue.location,
            issue.description,
            issue.resolution,
            issue.proposer,
            issue.status,
            revision_reference,
        ]
        for column, value in zip([2, 3, 4, 5, 6, 9, 11, 12, 13], values):
            cell = _writable_cell(sheet, row=row, column=column)
            cell.value = value
            cell.alignment = _WRAP_TOP
        sheet.row_dimensions[row].height = 78

    delete_count = _delete_unused_issue_rows(sheet, first_unused_row=13 + issue_count)
    if delete_count:
        _delete_extra_pending_rows(sheet)
    _rebuild_compact_review_layout(sheet)

    pending_row = 18
    conclusion_opinion_row = 20
    conclusion_tracking_row = 21
    conclusion_suggestion_row = 22

    sheet.cell(row=pending_row, column=2).value = '1'
    sheet.cell(row=pending_row, column=3).value = review_plan.pending_problem
    sheet.cell(row=pending_row, column=6).value = review_plan.pending_solution
    sheet.cell(row=pending_row, column=8).value = owner
    sheet.cell(row=pending_row, column=10).value = completion_date
    sheet.cell(row=pending_row, column=12).value = completion_date
    for column in [3, 6, 8, 10, 12]:
        sheet.cell(row=pending_row, column=column).alignment = _WRAP_TOP
    sheet.row_dimensions[pending_row].height = 46

    sheet.cell(row=conclusion_opinion_row, column=5).value = review_plan.conclusion
    sheet.cell(row=conclusion_tracking_row, column=5).value = review_plan.tracking
    sheet.cell(row=conclusion_suggestion_row, column=5).value = review_plan.suggestion
    for row in [conclusion_opinion_row, conclusion_tracking_row, conclusion_suggestion_row]:
        sheet.cell(row=row, column=5).alignment = _WRAP_TOP
        sheet.row_dimensions[row].height = 50


def _clear_issue_rows(sheet) -> None:
    for row in range(13, 23):
        for column in [2, 3, 4, 5, 6, 9, 11, 12, 13]:
            cell = sheet.cell(row=row, column=column)
            if not isinstance(cell, MergedCell):
                cell.value = None
    for row in range(26, 29):
        for column in [2, 3, 6, 8, 10, 12]:
            cell = sheet.cell(row=row, column=column)
            if not isinstance(cell, MergedCell):
                cell.value = None


def _normalize_issue_rows(sheet) -> None:
    for merged_range in list(sheet.merged_cells.ranges):
        if (
            merged_range.min_row >= 13
            and merged_range.max_row <= 22
            and merged_range.min_col >= 2
            and merged_range.max_col <= 13
        ):
            sheet.unmerge_cells(str(merged_range))


def _delete_unused_issue_rows(sheet, *, first_unused_row: int) -> int:
    last_unused_row = 23
    if first_unused_row > last_unused_row:
        return 0
    for merged_range in list(sheet.merged_cells.ranges):
        if merged_range.min_row <= last_unused_row and merged_range.max_row >= first_unused_row:
            sheet.unmerge_cells(str(merged_range))
    delete_count = last_unused_row - first_unused_row + 1
    sheet.delete_rows(first_unused_row, delete_count)
    return delete_count


def _delete_extra_pending_rows(sheet) -> None:
    for merged_range in list(sheet.merged_cells.ranges):
        if merged_range.min_row <= 20 and merged_range.max_row >= 19:
            sheet.unmerge_cells(str(merged_range))
    sheet.delete_rows(19, 2)


def _rebuild_compact_review_layout(sheet) -> None:
    for merged_range in list(sheet.merged_cells.ranges):
        if merged_range.min_row >= 13 and merged_range.max_row <= 22:
            sheet.unmerge_cells(str(merged_range))

    for row in [13, 14, 15]:
        sheet.row_dimensions[row].height = 78
        _merge_and_style(sheet, f'F{row}:H{row}')
        _merge_and_style(sheet, f'I{row}:J{row}')

    sheet.row_dimensions[16].height = 21.7
    _merge_and_style(sheet, 'B16:M16')
    sheet.row_dimensions[17].height = 17.55
    sheet.row_dimensions[18].height = 46
    for row in [17, 18]:
        for range_text in [f'C{row}:E{row}', f'F{row}:G{row}', f'H{row}:I{row}', f'J{row}:K{row}', f'L{row}:M{row}']:
            _merge_and_style(sheet, range_text)

    sheet.row_dimensions[19].height = 21.7
    _merge_and_style(sheet, 'B19:M19')
    for row in [20, 21, 22]:
        sheet.row_dimensions[row].height = 50
        _merge_and_style(sheet, f'B{row}:D{row}')
        _merge_and_style(sheet, f'E{row}:M{row}')


def _copy_row_style(sheet, *, source_row: int, target_rows: list[int], min_col: int, max_col: int) -> None:
    for target_row in target_rows:
        for column in range(min_col, max_col + 1):
            source = sheet.cell(row=source_row, column=column)
            target = sheet.cell(row=target_row, column=column)
            target._style = copy(source._style)
            if source.has_style:
                target.font = copy(source.font)
                target.fill = copy(source.fill)
                target.border = copy(source.border)
                target.alignment = copy(source.alignment)
                target.number_format = source.number_format
                target.protection = copy(source.protection)


def _merge_and_style(sheet, range_text: str) -> None:
    min_col, min_row, max_col, max_row = range_boundaries(range_text)
    source = sheet.cell(row=min_row, column=min_col)
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            target = sheet.cell(row=row, column=column)
            target._style = copy(source._style)
            if source.has_style:
                target.font = copy(source.font)
                target.fill = copy(source.fill)
                target.border = copy(source.border)
                target.alignment = copy(source.alignment)
                target.number_format = source.number_format
                target.protection = copy(source.protection)
    sheet.merge_cells(range_text)


def _writable_cell(sheet, *, row: int, column: int):
    cell = sheet.cell(row=row, column=column)
    if not isinstance(cell, MergedCell):
        return cell
    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell
