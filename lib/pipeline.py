from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook

from lib.doc_numbering import (
    customer_doc_filename,
    customer_doc_number,
    customer_review_filename,
    customer_review_record_number,
    dotted_date,
    ensure_run_date,
    iso_date,
    product_doc_filename,
    product_doc_number,
    product_review_filename,
    product_review_record_number,
    version_token,
)
from lib.docx_ops import (
    CustomerDocData,
    CustomerRequirementRow,
    CustomerScenario,
    ProductDocData,
    ProductRequirement,
    render_customer_doc,
    render_product_doc,
)
from lib.workbook_reader import (
    WorkbookData,
    enrich_workbook,
    first_sentence,
    inspect_workbook_headers,
    planned_enriched_workbook_path,
    read_workbook,
)
from lib.xlsx_review_sheet_ops import render_customer_review_sheet, render_product_review_sheet


@dataclass
class PipelineConfig:
    workbook: Path
    customer_template: Path
    product_template: Path
    review_template: Path
    module_code: str
    module_name: str
    draft_version: str
    final_version: str
    output_dir: Path
    module_product_manager: str
    module_project_manager: str
    reviewers: str
    qa: str
    run_date: str | None = None
    overwrite: bool = False


def plan_outputs(config: PipelineConfig) -> dict[str, object]:
    run_date = ensure_run_date(config.run_date)
    _validate_inputs(config)
    outputs = _output_paths(config)
    workbook_check = inspect_workbook_headers(config.workbook)
    prepared_workbook = planned_enriched_workbook_path(config.workbook, config.output_dir)
    if workbook_check.missing_required_headers:
        missing_text = '、'.join(workbook_check.missing_required_headers)
        raise ValueError(f'原始基础信息表必须包含这些列：{missing_text}')
    return {
        'module_code': config.module_code,
        'module_name': config.module_name,
        'draft_version': config.draft_version,
        'final_version': config.final_version,
        'run_date': run_date,
        'required_source_headers': ['用户需求id', '研发需求id', '研发需求名称', '研发需求描述'],
        'review_record_people': {
            'host': config.module_product_manager,
            'scribe': config.module_project_manager,
            'reviewers': config.reviewers,
            'other_people': config.qa,
        },
        'workbook_preparation': {
            'source_workbook': str(config.workbook),
            'prepared_workbook': str(prepared_workbook),
            'has_user_description_column': workbook_check.has_user_description,
            'has_acceptance_column': workbook_check.has_acceptance,
            'will_add_or_fill': ['用户需求描述', '验收标准'],
        },
        'doc_numbers': {
            'customer_draft': customer_doc_number(config.module_code, config.draft_version),
            'customer_final': customer_doc_number(config.module_code, config.final_version),
            'product_draft': product_doc_number(config.module_code, config.draft_version),
            'product_final': product_doc_number(config.module_code, config.final_version),
            'customer_review': customer_review_record_number(config.module_code, run_date),
            'product_review': product_review_record_number(config.module_code, run_date),
        },
        'content_correspondence': {
            'generation_flow': [
                'customer_draft',
                'customer_review',
                'customer_final',
                'product_draft',
                'product_review',
                'product_final',
            ],
            'customer_review_targets': outputs['customer_draft'].stem,
            'customer_final_must_match_requirement_scope_of': outputs['customer_draft'].stem,
            'product_review_targets': outputs['product_draft'].stem,
            'product_final_must_match_requirement_scope_of': outputs['product_draft'].stem,
        },
        'outputs': {name: str(path) for name, path in outputs.items()},
        'existing_outputs': {name: path.exists() for name, path in outputs.items()},
    }


def run_pipeline(config: PipelineConfig) -> dict[str, object]:
    run_date = ensure_run_date(config.run_date)
    _validate_inputs(config)
    outputs = _output_paths(config)
    prepared_workbook_path = planned_enriched_workbook_path(config.workbook, config.output_dir)
    collisions = [str(path) for path in [prepared_workbook_path, *outputs.values()] if path.exists()]
    if collisions and not config.overwrite:
        raise FileExistsError(f'output files already exist: {collisions}')

    config.output_dir.mkdir(parents=True, exist_ok=True)
    preparation = enrich_workbook(config.workbook, prepared_workbook_path)
    workbook = read_workbook(preparation.output_path)

    customer_draft = _build_customer_doc_data(config, workbook, config.draft_version, run_date, final=False)
    product_draft = _build_product_doc_data(config, workbook, config.draft_version, run_date, final=False)

    render_customer_doc(config.customer_template, outputs['customer_draft'], customer_draft)
    render_product_doc(config.product_template, outputs['product_draft'], product_draft)

    render_customer_review_sheet(
        config.review_template,
        outputs['customer_review'],
        record_number=customer_review_record_number(config.module_code, run_date),
        project_name=workbook.platform_branch or 'UKUI4.22',
        work_product_title=outputs['customer_draft'].stem,
        meeting_date=iso_date(run_date),
        host=config.module_product_manager,
        scribe=config.module_project_manager,
        reviewers=config.reviewers,
        other_people=config.qa,
    )

    customer_final = _build_customer_doc_data(config, workbook, config.final_version, run_date, final=True)
    product_final = _build_product_doc_data(config, workbook, config.final_version, run_date, final=True)
    render_customer_doc(config.customer_template, outputs['customer_final'], customer_final)
    render_product_doc(config.product_template, outputs['product_final'], product_final)
    render_product_review_sheet(
        config.review_template,
        outputs['product_review'],
        record_number=product_review_record_number(config.module_code, run_date),
        project_name=workbook.project_name or '银河麒麟桌面操作系统V11 SP1',
        work_product_title=outputs['product_draft'].stem,
        meeting_date=run_date,
        host=config.module_product_manager,
        scribe=config.module_project_manager,
        reviewers=config.reviewers,
        other_people=config.qa,
    )

    verification = verify_outputs(config)
    return {
        'run_date': run_date,
        'prepared_workbook': {
            'path': str(preparation.output_path),
            'added_user_description_column': preparation.added_user_description_column,
            'added_acceptance_column': preparation.added_acceptance_column,
            'filled_user_description_rows': preparation.filled_user_description_rows,
            'updated_acceptance_rows': preparation.updated_acceptance_rows,
            'total_requirement_rows': preparation.total_requirement_rows,
        },
        'generated_files': {name: str(path) for name, path in outputs.items()},
        'verification': verification,
    }


def verify_outputs(config: PipelineConfig) -> dict[str, object]:
    run_date = ensure_run_date(config.run_date)
    outputs = _output_paths(config)
    strict_review_date = config.run_date is not None
    expected_strings = {
        'customer_draft': customer_doc_number(config.module_code, config.draft_version),
        'customer_final': customer_doc_number(config.module_code, config.final_version),
        'product_draft': product_doc_number(config.module_code, config.draft_version),
        'product_final': product_doc_number(config.module_code, config.final_version),
        'customer_review': customer_review_record_number(config.module_code, run_date),
        'product_review': product_review_record_number(config.module_code, run_date),
    }
    review_prefixes = {
        'customer_review': f'MT-PR-A-{config.module_code}-CRS-',
        'product_review': f'MT-PR-A-{config.module_code}-PRD-',
    }
    results: dict[str, dict[str, object]] = {}
    for name, path in outputs.items():
        item = {
            'path': str(path),
            'exists': path.exists(),
            'expected_identifier': expected_strings[name],
            'identifier_present': False,
        }
        if path.exists() and path.suffix == '.docx':
            text = _read_docx_text(path)
            item['identifier_present'] = expected_strings[name] in text
        elif path.exists() and path.suffix == '.xlsx':
            text = _read_xlsx_text(path)
            if strict_review_date:
                item['identifier_present'] = expected_strings[name] in text
            else:
                item['identifier_present'] = review_prefixes[name] in text
        results[name] = item
    correspondence = _verify_content_correspondence(config, outputs)
    return {
        'run_date': run_date,
        'all_outputs_present': all(item['exists'] for item in results.values()),
        'all_identifiers_present': all(item['identifier_present'] for item in results.values()),
        'all_content_correspondence_passed': correspondence['passed'],
        'content_correspondence': correspondence,
        'outputs': results,
    }


def _validate_inputs(config: PipelineConfig) -> None:
    for path in [config.workbook, config.customer_template, config.product_template, config.review_template]:
        if not path.exists():
            raise FileNotFoundError(path)


def _output_paths(config: PipelineConfig) -> dict[str, Path]:
    return {
        'customer_draft': config.output_dir / customer_doc_filename(config.module_name, config.draft_version),
        'customer_final': config.output_dir / customer_doc_filename(config.module_name, config.final_version),
        'customer_review': config.output_dir / customer_review_filename(config.module_name),
        'product_draft': config.output_dir / product_doc_filename(config.module_name, config.draft_version),
        'product_final': config.output_dir / product_doc_filename(config.module_name, config.final_version),
        'product_review': config.output_dir / product_review_filename(config.module_name),
    }


def _build_customer_doc_data(config: PipelineConfig, workbook: WorkbookData, version: str, run_date: str, *, final: bool) -> CustomerDocData:
    version_display = version_token(version)
    scenarios = []
    requirement_rows = []
    for index, requirement in enumerate(workbook.requirements, start=1):
        scenarios.append(
            CustomerScenario(
                scenario_name=requirement.scenario_name,
                customer_type=requirement.customer_type,
                usage_context=requirement.usage_context,
                scenario_id=f'S-{index:02d}',
                description=_customer_scenario_description(requirement, final=final),
            )
        )
        requirement_rows.append(
            CustomerRequirementRow(
                user_requirement_id=requirement.user_requirement_id,
                title=requirement.title,
                description=_customer_requirement_description(requirement, scenario_id=f'S-{index:02d}'),
                priority=requirement.priority,
                owner_module=requirement.owner_module,
                collaborator_module=requirement.collaborator_module,
            )
        )

    return CustomerDocData(
        module_name=config.module_name,
        doc_number=customer_doc_number(config.module_code, version),
        version_display=version_display,
        run_date_display=dotted_date(run_date),
        release_note=_customer_release_note(config, final=final, run_date=run_date),
        scenarios=scenarios,
        requirement_rows=requirement_rows,
    )


def _build_product_doc_data(config: PipelineConfig, workbook: WorkbookData, version: str, run_date: str, *, final: bool) -> ProductDocData:
    version_display = version_token(version)
    requirements = []
    for requirement in workbook.requirements:
        requirements.append(
            ProductRequirement(
                module_name=config.module_name,
                level='C1',
                related_module_name='无',
                parent_module_name='无',
                chipset_special='/',
                user_requirement_id=requirement.user_requirement_id,
                rd_requirement_id=requirement.rd_requirement_id or requirement.sequence_id,
                priority_marks=_priority_marks(requirement.priority),
                implement_way='☑ 自研   □ 开源修改 □ 开源引入或第三方',
                feature_description=_product_feature_description(requirement),
                related_description='',
                acceptance=' '.join(part for part in requirement.acceptance.splitlines() if part),
                input_desc=first_sentence(requirement.user_description or requirement.rd_description) or f'用户围绕“{requirement.title}”发起操作。',
                process_desc=f'系统围绕“{requirement.title}”执行对应界面展示、状态处理或配置生效流程。',
                output_desc=f'用户可以围绕“{requirement.title}”稳定完成目标操作，并获得一致反馈。',
                exception_desc='异常场景下应给出可理解提示，并保持基础能力可用。',
                layout_desc='沿用现有页面/接口形态，详细图示按研发设计稿和联调结果补充。',
                quality_desc='兼容性、易用性与可靠性要求遵循当前版本基线。',
            )
        )

    quality_rows = [
        ('性能效率', f'{config.module_name}相关功能在高频交互场景下应保持流畅。', '重点关注切换、刷新、配置生效与状态同步时延。'),
        ('安全性', '接口、配置与入口调整不得突破现有权限边界。', '避免因开放能力或隐藏策略导致越权操作。'),
        ('可靠性', '连续调用、重复切换和状态频繁变化场景下应保持结果一致。', '异常时需保留基本能力并给出清晰反馈。'),
        ('兼容性（接口）', '界面能力应兼容当前主题和显示环境；接口能力兼容既有集成方式。', '命名、参数与返回格式保持稳定。'),
        ('可维护性', '配置项、接口项和状态项应遵循模块现有组织方式。', '便于后续版本继续演进与定位问题。'),
        ('可移植性', '输出内容应适配项目当前目标版本与交付环境。', '环境切换时不依赖额外人工修正。'),
        ('易用性', '入口位置、文案和反馈方式应尽量延续既有使用习惯。', '确保用户能快速理解状态并完成目标操作。'),
    ]

    reference_rows = [
        ('输入文档', f'《{config.module_name}-客户需求说明书_{config.draft_version}》'),
        ('依赖的接口契约', ''),
        ('遵循的标准与规范', '项目既有研发与评审规范'),
        ('……', ''),
        ('……', ''),
    ]

    return ProductDocData(
        module_name=config.module_name,
        doc_number=product_doc_number(config.module_code, version),
        version_display=version_display,
        run_date_display=dotted_date(run_date),
        release_note=_product_release_note(config, final=final, run_date=run_date),
        module_description=f'{config.module_name}模块属于桌面环境核心组成部分，负责承接本轮需求中与终端界面、交互入口、配置体验及集成能力相关的产品落地。',
        tech_constraints='需遵循现有桌面环境、设置框架和接口约束，优先复用既有实现方式。',
        resource_constraints='需结合当前版本节奏与研发资源推进，优先保证高价值需求按期落地。',
        assumptions='默认目标环境具备桌面环境基础能力，上下游模块可提供必要接口、配置与联调支持。',
        requirements=requirements,
        quality_rows=quality_rows,
        glossary_rows=_glossary_rows(workbook),
        reference_rows=reference_rows,
    )


def _priority_marks(priority: str) -> str:
    if priority == '高':
        return '☑ 高 □ 中 □ 低'
    if priority == '低':
        return '□ 高 □ 中 ☑ 低'
    return '□ 高 ☑ 中 □ 低'


def _customer_release_note(config: PipelineConfig, *, final: bool, run_date: str) -> str:
    if final:
        return f'根据同级评审记录 {customer_review_record_number(config.module_code, run_date)} 由初稿形成终稿'
    return '初稿，提交同级评审'


def _product_release_note(config: PipelineConfig, *, final: bool, run_date: str) -> str:
    if final:
        return f'根据同级评审记录 {product_review_record_number(config.module_code, run_date)} 由初稿形成终稿'
    return '初稿，提交同级评审'


def _customer_scenario_description(requirement, *, final: bool) -> str:
    if final:
        return f'背景：{first_sentence(requirement.user_description) or first_sentence(requirement.rd_description)} 使用目标：围绕“{requirement.title}”提升操作稳定性与体验一致性。 成功标志：{first_sentence(requirement.acceptance) or "能力按预期生效。"}'
    return first_sentence(requirement.user_description) or first_sentence(requirement.rd_description) or f'围绕“{requirement.title}”提供对应场景支撑。'


def _customer_requirement_description(requirement, *, scenario_id: str) -> str:
    summary = first_sentence(requirement.user_description) or first_sentence(requirement.rd_description)
    role_goal = summary or f'用户围绕“{requirement.title}”发起操作，希望核心能力可用且结果清晰。'
    return f'角色与目标：{role_goal} 业务价值：支撑该需求在版本内清晰落地，并降低理解与执行成本。 归属场景（场景编号）：{scenario_id}'


def _product_feature_description(requirement) -> str:
    summary = ' '.join(part for part in requirement.rd_description.splitlines() if part) or first_sentence(requirement.user_description)
    return f'围绕客户需求 {requirement.user_requirement_id} 与研发需求 {requirement.rd_requirement_id}，落实“{requirement.title}”能力。 {summary}'.strip()


def _glossary_rows(workbook: WorkbookData) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    rows.append((workbook.platform_branch or 'UKUI4.22', '本次需求归属的平台或分支标识。'))
    seen = set()
    for requirement in workbook.requirements:
        if requirement.owner_module not in seen:
            seen.add(requirement.owner_module)
            rows.append((requirement.owner_module, f'与“{requirement.title}”相关的主责能力域。'))
        if len(rows) >= 10:
            break
    rows.extend([
        ('D-Bus', '系统组件之间及第三方客户端进行状态获取和控制调用的常见方式。'),
        ('FTP', '文件传输协议相关远程访问场景。'),
        ('UTF-8', '常见字符编码格式。'),
    ])
    deduped: list[tuple[str, str]] = []
    used = set()
    for key, value in rows:
        if key and key not in used:
            used.add(key)
            deduped.append((key, value))
    return deduped[:10]


def _read_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
    return ''.join(node.text or '' for node in root.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))


def _read_xlsx_text(path: Path) -> str:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook.active
    values: list[str] = []
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                values.append(str(cell.value))
    return '\n'.join(values)


def _verify_content_correspondence(config: PipelineConfig, outputs: dict[str, Path]) -> dict[str, object]:
    prepared_workbook = planned_enriched_workbook_path(config.workbook, config.output_dir)
    result: dict[str, object] = {
        'passed': False,
        'prepared_workbook': str(prepared_workbook),
        'prepared_workbook_exists': prepared_workbook.exists(),
        'customer_review_targets_draft': False,
        'customer_review_targets_final': False,
        'product_review_targets_draft': False,
        'product_review_targets_final': False,
        'customer_final_mentions_review_record': False,
        'product_final_mentions_review_record': False,
        'missing_customer_requirement_ids_in_draft': [],
        'missing_customer_requirement_ids_in_final': [],
        'missing_product_requirement_ids_in_draft': [],
        'missing_product_requirement_ids_in_final': [],
    }
    required_outputs = [
        outputs['customer_draft'],
        outputs['customer_final'],
        outputs['customer_review'],
        outputs['product_draft'],
        outputs['product_final'],
        outputs['product_review'],
    ]
    if not prepared_workbook.exists() or not all(path.exists() for path in required_outputs):
        return result

    try:
        workbook = read_workbook(prepared_workbook)
    except Exception as exc:  # pragma: no cover - returned for CLI diagnostics
        result['error'] = str(exc)
        return result

    customer_ids = [
        requirement.user_requirement_id
        for requirement in workbook.requirements
        if requirement.user_requirement_id
    ]
    product_ids = [
        requirement.rd_requirement_id or requirement.sequence_id
        for requirement in workbook.requirements
        if requirement.rd_requirement_id or requirement.sequence_id
    ]
    result['requirement_count'] = len(workbook.requirements)

    customer_draft_text = _read_docx_text(outputs['customer_draft'])
    customer_final_text = _read_docx_text(outputs['customer_final'])
    product_draft_text = _read_docx_text(outputs['product_draft'])
    product_final_text = _read_docx_text(outputs['product_final'])
    customer_review_text = _read_xlsx_text(outputs['customer_review'])
    product_review_text = _read_xlsx_text(outputs['product_review'])

    result['customer_review_targets_draft'] = outputs['customer_draft'].stem in customer_review_text
    result['customer_review_targets_final'] = outputs['customer_final'].stem in customer_review_text
    result['product_review_targets_draft'] = outputs['product_draft'].stem in product_review_text
    result['product_review_targets_final'] = outputs['product_final'].stem in product_review_text
    result['customer_final_mentions_review_record'] = f'MT-PR-A-{config.module_code}-CRS-' in customer_final_text
    result['product_final_mentions_review_record'] = f'MT-PR-A-{config.module_code}-PRD-' in product_final_text

    result['missing_customer_requirement_ids_in_draft'] = _missing_tokens(customer_draft_text, customer_ids)
    result['missing_customer_requirement_ids_in_final'] = _missing_tokens(customer_final_text, customer_ids)
    result['missing_product_requirement_ids_in_draft'] = _missing_tokens(product_draft_text, product_ids)
    result['missing_product_requirement_ids_in_final'] = _missing_tokens(product_final_text, product_ids)

    result['passed'] = (
        result['customer_review_targets_draft']
        and not result['customer_review_targets_final']
        and result['product_review_targets_draft']
        and not result['product_review_targets_final']
        and result['customer_final_mentions_review_record']
        and result['product_final_mentions_review_record']
        and not result['missing_customer_requirement_ids_in_draft']
        and not result['missing_customer_requirement_ids_in_final']
        and not result['missing_product_requirement_ids_in_draft']
        and not result['missing_product_requirement_ids_in_final']
    )
    return result


def _missing_tokens(text: str, tokens: list[str]) -> list[str]:
    return [token for token in tokens if token and token not in text]
