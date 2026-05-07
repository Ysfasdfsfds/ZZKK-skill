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
    host: str,
    scribe: str,
    reviewers: str,
    other_people: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = f'《{work_product_title}》'
    sheet.cell(row=5, column=5).value = '应用场景表、用户需求描述表、需求条目、优先级、主责/协作模块及文档编号一致性检查'
    sheet.cell(row=5, column=11).value = '需求分析与文档评审'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 1.5
    sheet.cell(row=7, column=5).value = host
    sheet.cell(row=7, column=9).value = scribe
    sheet.cell(row=8, column=5).value = reviewers
    sheet.cell(row=9, column=5).value = other_people
    sheet.cell(row=13, column=2).value = 1
    sheet.cell(row=13, column=3).value = work_product_title
    sheet.cell(row=13, column=4).value = '全文'
    sheet.cell(row=13, column=5).value = '评审结论'
    sheet.cell(row=13, column=6).value = '本次同级评审未发现阻断性缺陷，文档内容与基础信息表、模板结构和编号规则一致。'
    sheet.cell(row=13, column=9).value = '无需整改；按会议实际意见持续维护。'
    sheet.cell(row=13, column=12).value = scribe
    sheet.cell(row=13, column=13).value = '关闭'
    sheet.cell(row=26, column=3).value = '无待解决问题'
    sheet.cell(row=26, column=6).value = '无需处理'
    sheet.cell(row=26, column=8).value = host
    sheet.cell(row=26, column=10).value = '无'
    sheet.cell(row=26, column=12).value = '无'
    sheet.cell(row=30, column=5).value = '经同级评审，客户需求说明书已覆盖应用场景、用户需求描述、优先级、主责模块和协作模块信息，文档编号与版本信息正确，未发现阻断性问题，同意通过评审。'
    sheet.cell(row=31, column=5).value = '本次评审无待解决缺陷；后续如会议现场提出新增意见，由模块产品经理登记并跟踪闭环。'
    sheet.cell(row=32, column=5).value = '建议后续根据正式会议签到、评审意见和研发设计稿持续补充页号、提出人、签字等现场信息。'
    workbook.save(output_path)


def render_product_review_sheet(
    template_path: Path,
    output_path: Path,
    *,
    record_number: str,
    project_name: str,
    work_product_title: str,
    meeting_date: str,
    host: str,
    scribe: str,
    reviewers: str,
    other_people: str,
) -> None:
    workbook = load_workbook(template_path)
    sheet = workbook.active
    sheet.cell(row=2, column=6).value = record_number
    sheet.cell(row=3, column=5).value = project_name
    sheet.cell(row=4, column=5).value = f'《{work_product_title}》'
    sheet.cell(row=5, column=5).value = '模块描述、假设与约束、功能详述、非功能要求、术语、参考资料及需求编号一致性检查'
    sheet.cell(row=5, column=11).value = '设计阶段'
    sheet.cell(row=6, column=5).value = meeting_date
    sheet.cell(row=6, column=9).value = 2
    sheet.cell(row=7, column=5).value = host
    sheet.cell(row=7, column=9).value = scribe
    sheet.cell(row=8, column=5).value = reviewers
    sheet.cell(row=9, column=5).value = other_people
    sheet.cell(row=13, column=2).value = 1
    sheet.cell(row=13, column=3).value = work_product_title
    sheet.cell(row=13, column=4).value = '全文'
    sheet.cell(row=13, column=5).value = '评审结论'
    sheet.cell(row=13, column=6).value = '本次同级评审未发现阻断性缺陷，文档内容与基础信息表、模板结构和编号规则一致。'
    sheet.cell(row=13, column=9).value = '无需整改；按会议实际意见持续维护。'
    sheet.cell(row=13, column=12).value = scribe
    sheet.cell(row=13, column=13).value = '关闭'
    sheet.cell(row=26, column=3).value = '无待解决问题'
    sheet.cell(row=26, column=6).value = '无需处理'
    sheet.cell(row=26, column=8).value = host
    sheet.cell(row=26, column=10).value = '无'
    sheet.cell(row=26, column=12).value = '无'
    sheet.cell(row=30, column=5).value = '经同级评审，模块级产品需求分析说明书已覆盖功能描述、验收标准、输入/过程/输出、异常流程和质量特性，文档编号与版本信息正确，未发现阻断性问题，同意通过评审。'
    sheet.cell(row=31, column=5).value = '本次评审无待解决缺陷；后续接口契约、页面布局和设计稿如有更新，由责任人同步维护文档并闭环确认。'
    sheet.cell(row=32, column=5).value = '建议后续结合接口文档、Demo、设计稿和测试反馈补充更细的页号、用例图、提出人及签字信息。'
    workbook.save(output_path)
