from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}


def _qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _table_text(tbl: ET.Element) -> str:
    return "".join(node.text or "" for node in tbl.findall('.//w:t', NS))


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall('.//w:t', NS)).strip()


def _cell_text(cell: ET.Element) -> str:
    return "".join(node.text or "" for node in cell.findall('.//w:t', NS))


def _paragraphs(cell: ET.Element) -> list[ET.Element]:
    paragraphs = cell.findall('w:p', NS)
    if paragraphs:
        return paragraphs
    paragraph = ET.SubElement(cell, _qn('p'))
    return [paragraph]


def _copy_run_props(run: ET.Element | None) -> ET.Element | None:
    if run is None:
        return None
    props = run.find('w:rPr', NS)
    return deepcopy(props) if props is not None else None


def _text_runs(paragraph: ET.Element) -> list[ET.Element]:
    runs: list[ET.Element] = []
    for run in paragraph.findall('w:r', NS):
        text = ''.join(node.text or '' for node in run.findall('w:t', NS))
        if text:
            runs.append(run)
    return runs


def _clear_paragraph(paragraph: ET.Element) -> None:
    for child in list(paragraph):
        if child.tag != _qn('pPr'):
            paragraph.remove(child)


def _append_run(paragraph: ET.Element, text: str, template_run: ET.Element | None) -> None:
    run = ET.SubElement(paragraph, _qn('r'))
    props = _copy_run_props(template_run)
    if props is not None:
        run.append(props)
    text_node = ET.SubElement(run, _qn('t'))
    if text.startswith(' ') or text.endswith(' '):
        text_node.set(f'{{{XML_NS}}}space', 'preserve')
    text_node.text = text


def _set_cell_runs(cell: ET.Element, run_specs: list[tuple[str, ET.Element | None]]) -> None:
    paragraphs = _paragraphs(cell)
    paragraph = paragraphs[0]
    for extra in paragraphs[1:]:
        cell.remove(extra)
    _clear_paragraph(paragraph)
    for text, template in run_specs:
        if text:
            _append_run(paragraph, text, template)


def _set_cell_text(cell: ET.Element, text: str) -> None:
    paragraph = _paragraphs(cell)[0]
    runs = _text_runs(paragraph)
    template = runs[0] if runs else None
    _set_cell_runs(cell, [(text, template)])


def _set_labeled_cell(cell: ET.Element, label: str, content: str) -> None:
    paragraph = _paragraphs(cell)[0]
    runs = _text_runs(paragraph)
    label_template = runs[0] if runs else None
    content_template = runs[1] if len(runs) > 1 else label_template
    specs = [(label, label_template)]
    if content:
        specs.append((content, content_template))
    _set_cell_runs(cell, specs)


def _tables(root: ET.Element) -> list[ET.Element]:
    return root.findall('.//w:tbl', NS)


def _find_first_table(root: ET.Element, *markers: str) -> ET.Element:
    for table in _tables(root):
        text = _table_text(table)
        if all(marker in text for marker in markers):
            return table
    joined = ', '.join(markers)
    raise ValueError(f'could not find table with markers: {joined}')


def _find_all_tables(root: ET.Element, *markers: str) -> list[ET.Element]:
    matched: list[ET.Element] = []
    for table in _tables(root):
        text = _table_text(table)
        if all(marker in text for marker in markers):
            matched.append(table)
    return matched


def _row_cells(table: ET.Element) -> list[list[ET.Element]]:
    rows: list[list[ET.Element]] = []
    for row in table.findall('w:tr', NS):
        rows.append(row.findall('w:tc', NS))
    return rows


def _ensure_table_data_rows(table: ET.Element, required_total_rows: int) -> None:
    rows = table.findall('w:tr', NS)
    if len(rows) >= required_total_rows:
        return
    if len(rows) < 2:
        raise ValueError('table does not have a template data row to clone')
    template_row = rows[-1]
    for _ in range(required_total_rows - len(rows)):
        table.append(deepcopy(template_row))


def _find_direct_paragraph_index(root: ET.Element, paragraph_text: str) -> int:
    body = root.find('w:body', NS)
    if body is None:
        raise ValueError('docx body is missing')
    for index, child in enumerate(list(body)):
        if child.tag == _qn('p') and _paragraph_text(child) == paragraph_text:
            return index
    raise ValueError(f'could not find paragraph: {paragraph_text}')


def _find_direct_table_indices(root: ET.Element, *markers: str) -> list[int]:
    body = root.find('w:body', NS)
    if body is None:
        return []
    indices: list[int] = []
    for index, child in enumerate(list(body)):
        if child.tag != _qn('tbl'):
            continue
        text = _table_text(child)
        if all(marker in text for marker in markers):
            indices.append(index)
    return indices


def _ensure_repeated_tables_before_heading(root: ET.Element, *, markers: tuple[str, ...], required_count: int, stop_heading: str) -> None:
    body = root.find('w:body', NS)
    if body is None:
        raise ValueError('docx body is missing')
    table_indices = _find_direct_table_indices(root, *markers)
    if not table_indices:
        raise ValueError(f'could not find repeated tables for markers: {markers}')
    if len(table_indices) >= required_count:
        return
    children = list(body)
    stop_index = _find_direct_paragraph_index(root, stop_heading)
    table_template = deepcopy(children[table_indices[-1]])
    spacer_template = None
    spacer_index = table_indices[0] + 1
    if spacer_index < len(children) and children[spacer_index].tag == _qn('p'):
        spacer_template = deepcopy(children[spacer_index])
    for _ in range(required_count - len(table_indices)):
        insert_at = stop_index
        if spacer_template is not None:
            body.insert(insert_at, deepcopy(spacer_template))
            insert_at += 1
            stop_index += 1
        body.insert(insert_at, deepcopy(table_template))
        stop_index += 1


def _replace_direct_paragraph_text(root: ET.Element, old_text: str, new_text: str) -> bool:
    body = root.find('w:body', NS)
    if body is None:
        return False
    for child in list(body):
        if child.tag != _qn('p'):
            continue
        if _paragraph_text(child) != old_text:
            continue
        runs = _text_runs(child)
        template = runs[0] if runs else None
        _clear_paragraph(child)
        _append_run(child, new_text, template)
        return True
    return False


def _replace_following_paragraph(root: ET.Element, heading_text: str, new_text: str) -> bool:
    body = root.find('w:body', NS)
    if body is None:
        return False
    children = list(body)
    for index, child in enumerate(children):
        if child.tag != _qn('p'):
            continue
        if _paragraph_text(child) != heading_text:
            continue
        if index + 1 >= len(children):
            return False
        sibling = children[index + 1]
        if sibling.tag != _qn('p'):
            return False
        runs = _text_runs(sibling)
        template = runs[0] if runs else None
        _clear_paragraph(sibling)
        _append_run(sibling, new_text, template)
        return True
    return False


def _replace_section_block(root: ET.Element, heading_text: str, stop_heading_text: str, paragraphs: list[str]) -> bool:
    body = root.find('w:body', NS)
    if body is None:
        return False
    children = list(body)
    heading_index = -1
    stop_index = -1
    for index, child in enumerate(children):
        if child.tag != _qn('p'):
            continue
        text = _paragraph_text(child)
        if text == heading_text:
            heading_index = index
        elif heading_index != -1 and text == stop_heading_text:
            stop_index = index
            break
    if heading_index == -1 or stop_index == -1:
        return False
    template = None
    insert_at = heading_index + 1
    removable: list[ET.Element] = []
    for child in children[insert_at:stop_index]:
        if child.tag == _qn('p') and template is None:
            runs = _text_runs(child)
            template = runs[0] if runs else None
        removable.append(child)
    for child in removable:
        body.remove(child)
    insert_pos = insert_at
    for paragraph_text in paragraphs:
        paragraph = ET.Element(_qn('p'))
        _append_run(paragraph, paragraph_text, template)
        body.insert(insert_pos, paragraph)
        insert_pos += 1
    return True


def _write_docx(template_path: Path, output_path: Path, root: ET.Element) -> None:
    document_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    with zipfile.ZipFile(template_path) as source, zipfile.ZipFile(output_path, 'w') as target:
        for member in source.infolist():
            data = document_xml if member.filename == 'word/document.xml' else source.read(member.filename)
            target.writestr(member, data)


@dataclass
class CustomerScenario:
    scenario_name: str
    customer_type: str
    usage_context: str
    scenario_id: str
    description: str


@dataclass
class CustomerRequirementRow:
    user_requirement_id: str
    title: str
    description: str
    priority: str
    owner_module: str
    collaborator_module: str


@dataclass
class CustomerDocData:
    module_name: str
    doc_number: str
    version_display: str
    run_date_display: str
    product_intro: str
    version_changes: str
    target_customers: str
    module_rows: list[tuple[str, str, str, str, str, str, str]]
    scenarios: list[CustomerScenario]
    requirement_rows: list[CustomerRequirementRow]
    delivery_rows: list[tuple[str, str, str, str]]
    glossary_rows: list[tuple[str, str]]
    reference_rows: list[tuple[str, str]]


@dataclass
class ProductRequirement:
    module_name: str
    level: str
    related_module_name: str
    parent_module_name: str
    chipset_special: str
    user_requirement_id: str
    rd_requirement_id: str
    priority_marks: str
    implement_way: str
    feature_description: str
    related_description: str
    acceptance: str
    input_desc: str
    process_desc: str
    output_desc: str
    exception_desc: str
    layout_desc: str
    quality_desc: str


@dataclass
class ProductDocData:
    module_name: str
    doc_number: str
    version_display: str
    run_date_display: str
    module_description: str
    tech_constraints: str
    resource_constraints: str
    assumptions: str
    requirements: list[ProductRequirement]
    quality_rows: list[tuple[str, str, str]]
    glossary_rows: list[tuple[str, str]]
    reference_rows: list[tuple[str, str]]


def render_customer_doc(template_path: Path, output_path: Path, data: CustomerDocData) -> None:
    with zipfile.ZipFile(template_path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))

    doc_number_table = _find_first_table(root, '文件编号：')
    _set_cell_text(doc_number_table.findall('.//w:tc', NS)[0], f'文件编号：{data.doc_number}')

    _replace_direct_paragraph_text(root, '桌面环境模块', f'{data.module_name}模块')
    _replace_direct_paragraph_text(root, '产品客户需求说明书', '产品客户需求说明书')

    version_tables = _find_all_tables(root, '日期', '版本号', '发布说明', '编写者')
    if version_tables:
        rows = _row_cells(version_tables[0])
        if len(rows) > 1 and len(rows[1]) >= 4:
            _set_cell_text(rows[1][0], data.run_date_display)
            _set_cell_text(rows[1][1], data.version_display)
            _set_cell_text(rows[1][2], '自动生成')
            _set_cell_text(rows[1][3], 'Claude')

    overview_table = _find_first_table(root, '产品介绍（用途、范围、主要功能）', '主要面向客户')
    overview_cells = overview_table.findall('.//w:tc', NS)
    _set_cell_text(overview_cells[1], data.product_intro)
    _set_cell_text(overview_cells[3], data.version_changes)
    _set_cell_text(overview_cells[5], data.target_customers)

    module_table = _find_first_table(root, '模块ID (唯一)', '当前稳定版本')
    _ensure_table_data_rows(module_table, len(data.module_rows) + 1)
    module_rows = _row_cells(module_table)
    for row_index in range(1, len(module_rows)):
        values = data.module_rows[row_index - 1] if row_index - 1 < len(data.module_rows) else ('', '', '', '', '', '', '')
        for col_index, value in enumerate(values[: len(module_rows[row_index])]):
            _set_cell_text(module_rows[row_index][col_index], value)

    _ensure_repeated_tables_before_heading(root, markers=('场景名称', '场景描述'), required_count=len(data.scenarios), stop_heading='4 用户需求详述与模块分解')
    scenario_tables = _find_all_tables(root, '场景名称', '场景描述')
    for index, table in enumerate(scenario_tables):
        cells = table.findall('.//w:tc', NS)
        if index < len(data.scenarios):
            scenario = data.scenarios[index]
            values = [scenario.scenario_name, scenario.customer_type, scenario.usage_context, scenario.scenario_id, scenario.description]
        else:
            values = ['', '', '', '', '']
        for cell_index, value in zip([1, 3, 5, 7, 9], values):
            _set_cell_text(cells[cell_index], value)

    demand_table = _find_first_table(root, '用户需求ID', '协作模块')
    _ensure_table_data_rows(demand_table, len(data.requirement_rows) + 1)
    demand_rows = _row_cells(demand_table)
    for row_index in range(1, len(demand_rows)):
        row_values = data.requirement_rows[row_index - 1] if row_index - 1 < len(data.requirement_rows) else CustomerRequirementRow('', '', '', '', '', '')
        values = [
            row_values.user_requirement_id,
            row_values.title,
            row_values.description,
            row_values.priority,
            row_values.owner_module,
            row_values.collaborator_module,
        ]
        for col_index, value in enumerate(values[: len(demand_rows[row_index])]):
            _set_cell_text(demand_rows[row_index][col_index], value)

    delivery_table = _find_first_table(root, '交付物名称', '交付物形态', '说明')
    delivery_rows = _row_cells(delivery_table)
    for row_index in range(1, len(delivery_rows)):
        values = data.delivery_rows[row_index - 1] if row_index - 1 < len(data.delivery_rows) else ('', '', '', '')
        for col_index, value in enumerate(values[: len(delivery_rows[row_index])]):
            _set_cell_text(delivery_rows[row_index][col_index], value)

    glossary_table = _find_first_table(root, '术语/缩写', '定义/解释')
    glossary_rows = _row_cells(glossary_table)
    glossary_capacity = (len(glossary_rows) - 1)
    for index in range(glossary_capacity):
        left = data.glossary_rows[index] if index < len(data.glossary_rows) else ('', '')
        row = glossary_rows[index + 1]
        if len(row) >= 2:
            _set_cell_text(row[0], left[0])
            _set_cell_text(row[1], left[1])

    reference_table = _find_first_table(root, '类别', '文档名称')
    reference_rows = _row_cells(reference_table)
    for index in range(1, len(reference_rows)):
        values = data.reference_rows[index - 1] if index - 1 < len(data.reference_rows) else ('', '')
        if len(reference_rows[index]) >= 2:
            _set_cell_text(reference_rows[index][0], values[0])
            _set_cell_text(reference_rows[index][1], values[1])

    _write_docx(template_path, output_path, root)


def render_product_doc(template_path: Path, output_path: Path, data: ProductDocData) -> None:
    with zipfile.ZipFile(template_path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))

    doc_number_table = _find_first_table(root, '文件编号：')
    _set_cell_text(doc_number_table.findall('.//w:tc', NS)[0], f'文件编号：{data.doc_number}')

    _replace_direct_paragraph_text(root, '桌面环境模块', f'{data.module_name}模块')
    _replace_direct_paragraph_text(root, '产品需求分析说明书', '产品需求分析说明书')

    version_tables = _find_all_tables(root, '日期', '版本号', '发布说明', '编写者')
    if version_tables:
        rows = _row_cells(version_tables[0])
        if len(rows) > 1 and len(rows[1]) >= 4:
            _set_cell_text(rows[1][0], data.run_date_display)
            _set_cell_text(rows[1][1], data.version_display)
            _set_cell_text(rows[1][2], '自动生成')
            _set_cell_text(rows[1][3], 'Claude')

    if not _replace_section_block(root, '模块描述', '假设与约束', [
        data.module_description,
        f'定义模块：{data.module_name}模块提供与桌面环境核心交互、配置体验及相关能力交付有关的产品支撑。',
        '阐述核心职责：承接本轮需求中的界面、交互、配置与集成相关能力，确保版本目标可落地、可评审、可验收。',
    ]):
        raise ValueError('failed to replace 模块描述 section')
    if not _replace_following_paragraph(root, '技术约束', data.tech_constraints):
        raise ValueError('failed to replace 技术约束 paragraph')
    if not _replace_following_paragraph(root, '资源约束', data.resource_constraints):
        raise ValueError('failed to replace 资源约束 paragraph')
    if not _replace_following_paragraph(root, '假设条件', data.assumptions):
        raise ValueError('failed to replace 假设条件 paragraph')

    _ensure_repeated_tables_before_heading(root, markers=('模块名称', '功能描述：', '验收标准：'), required_count=len(data.requirements), stop_heading='模块新增非功能详述')
    requirement_tables = _find_all_tables(root, '模块名称', '功能描述：', '验收标准：')
    for index, table in enumerate(requirement_tables):
        cells = table.findall('.//w:tc', NS)
        if index < len(data.requirements):
            requirement = data.requirements[index]
            values = {
                1: requirement.module_name,
                3: requirement.level,
                5: requirement.related_module_name,
                7: requirement.parent_module_name,
                9: requirement.chipset_special,
                11: requirement.user_requirement_id,
                13: requirement.rd_requirement_id,
                15: requirement.priority_marks,
                17: requirement.implement_way,
            }
            for cell_index, value in values.items():
                _set_cell_text(cells[cell_index], value)
            _set_labeled_cell(cells[18], '功能描述：', requirement.feature_description)
            _set_labeled_cell(cells[19], '关联模块描述：', requirement.related_description)
            _set_labeled_cell(cells[20], '验收标准：', requirement.acceptance)
            _set_labeled_cell(cells[21], '输入：', requirement.input_desc)
            _set_labeled_cell(cells[22], '过程：', requirement.process_desc)
            _set_labeled_cell(cells[23], '输出：', requirement.output_desc)
            _set_labeled_cell(cells[24], '异常流程：', requirement.exception_desc)
            _set_labeled_cell(cells[25], '页面布局/用例图：', requirement.layout_desc)
            _set_labeled_cell(cells[26], '质量特性：', requirement.quality_desc)
        else:
            for cell_index in [1, 3, 5, 7, 9, 11, 13, 15, 17]:
                _set_cell_text(cells[cell_index], '')
            for cell_index, label in [(18, '功能描述：'), (19, '关联模块描述：'), (20, '验收标准：'), (21, '输入：'), (22, '过程：'), (23, '输出：'), (24, '异常流程：'), (25, '页面布局/用例图：'), (26, '质量特性：')]:
                _set_labeled_cell(cells[cell_index], label, '')

    quality_table = _find_first_table(root, '质量特性', '模块级要求', '说明/约束')
    quality_rows = _row_cells(quality_table)
    for row_index in range(1, len(quality_rows)):
        values = data.quality_rows[row_index - 1] if row_index - 1 < len(data.quality_rows) else ('', '', '')
        for col_index, value in enumerate(values[: len(quality_rows[row_index])]):
            _set_cell_text(quality_rows[row_index][col_index], value)

    glossary_table = _find_first_table(root, '术语/缩写', '定义/解释')
    glossary_rows = _row_cells(glossary_table)
    for index in range(1, len(glossary_rows)):
        values = data.glossary_rows[index - 1] if index - 1 < len(data.glossary_rows) else ('', '')
        if len(glossary_rows[index]) >= 2:
            _set_cell_text(glossary_rows[index][0], values[0])
            _set_cell_text(glossary_rows[index][1], values[1])

    reference_table = _find_first_table(root, '参考资料类别', '文档名称')
    reference_rows = _row_cells(reference_table)
    for index in range(1, len(reference_rows)):
        values = data.reference_rows[index - 1] if index - 1 < len(data.reference_rows) else ('', '')
        if len(reference_rows[index]) >= 2:
            _set_cell_text(reference_rows[index][0], values[0])
            _set_cell_text(reference_rows[index][1], values[1])

    _write_docx(template_path, output_path, root)
