from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


def render_customer_review_sheet(
    template_path: Path,
    output_path: Path,
    *,
    record_number: str,
    project_name: str,
    work_product_title: str,
    meeting_date: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = f'《{work_product_title}》'
    sheet.cell(row=5, column=5).value = '第3/4/6/7/8章的场景、需求分解、交付物、术语和参考资料一致性检查'
    sheet.cell(row=5, column=11).value = '需求分析与文档评审'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 1.5
    sheet.cell(row=7, column=5).value = '待补充'
    sheet.cell(row=7, column=9).value = '待补充'
    sheet.cell(row=8, column=5).value = '待补充'
    sheet.cell(row=9, column=5).value = '待补充'
    workbook.save(output_path)


def render_product_review_sheet(
    template_path: Path,
    output_path: Path,
    *,
    record_number: str,
    project_name: str,
    work_product_title: str,
    meeting_date: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = project_name
    sheet.cell(row=5, column=5).value = f'《{work_product_title}》'
    sheet.cell(row=5, column=11).value = '设计阶段'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 2
    sheet.cell(row=7, column=5).value = '待补充'
    sheet.cell(row=7, column=9).value = '待补充'
    sheet.cell(row=8, column=5).value = '待补充'
    sheet.cell(row=9, column=5).value = '待补充'
    workbook.save(output_path)
