from __future__ import annotations

from dataclasses import dataclass

from lib.workbook_reader import RequirementEntry, WorkbookData, first_sentence


@dataclass(frozen=True)
class ReviewIssue:
    sequence: str
    document: str
    page_scope: str
    location: str
    description: str
    resolution: str
    proposer: str
    status: str = '关闭'


@dataclass(frozen=True)
class ReviewPlan:
    issues: list[ReviewIssue]
    pending_problem: str
    pending_solution: str
    conclusion: str
    tracking: str
    suggestion: str


def build_customer_review_plan(
    *,
    module_name: str,
    draft_title: str,
    final_version: str,
    workbook: WorkbookData,
    module_product_manager: str,
    product_owner: str,
    test_owner: str,
    security_owner: str,
) -> ReviewPlan:
    focus = _focus_summary(workbook.requirements)
    issues = [
        ReviewIssue(
            '1',
            draft_title,
            '全文',
            '应用场景和用户需求描述',
            f'应用场景描述偏通用，未充分区分{focus}等客户使用场景，触发条件、使用目标和成功标准不够明确。',
            '终稿按需求条目补充“背景/使用目标/成功标准”，并保持场景编号与用户需求 ID 一一对应。',
            product_owner,
        ),
        ReviewIssue(
            '2',
            draft_title,
            '用户需求描述表',
            '需求描述列',
            '部分条目主要由研发需求名称和描述扩展，用户角色、业务价值和验收关注点不足，难以支撑客户侧理解和确认。',
            '终稿在用户需求描述表补充“角色与目标/业务价值/验收关注/归属场景”，使客户诉求与基础信息表闭环。',
            _join_people(product_owner, test_owner),
        ),
        ReviewIssue(
            '3',
            draft_title,
            '非功能相关条目',
            '安全、性能、异常提示',
            f'{_nonfunctional_focus(workbook.requirements)}等非功能约束缺少客户侧汇总说明。',
            '终稿在相关需求描述和场景成功标准中补充安全边界、性能指标、异常提示和可用性要求。',
            _join_people(security_owner, test_owner),
        ),
    ]
    return ReviewPlan(
        issues=issues,
        pending_problem='评审意见已全部按终稿修订闭环。',
        pending_solution=f'对应 {final_version} 终稿已按上述评审意见完成修订，问题状态均为关闭。',
        conclusion=(
            f'本次评审从产品、研发、安全和测试视角发现 {len(issues)} 项需完善意见。'
            f'{final_version} 终稿已补充应用场景差异、用户角色与业务价值、成功标准及非功能关注点，'
            '需求范围与初稿保持一致，同意通过评审。'
        ),
        tracking=f'上述 {len(issues)} 项问题均已在对应 {final_version} 终稿中修订闭环；后续若正式会议新增意见，由模块产品经理登记并跟踪。',
        suggestion='建议后续结合正式会议签到、设计稿、接口文档和测试用例继续补充页号、提出人、签字等现场信息。',
    )


def build_product_review_plan(
    *,
    module_name: str,
    draft_title: str,
    final_version: str,
    workbook: WorkbookData,
    module_project_manager: str,
    rd_owner: str,
    test_owner: str,
    security_owner: str,
) -> ReviewPlan:
    focus = _focus_summary(workbook.requirements)
    issues = [
        ReviewIssue(
            '1',
            draft_title,
            '第3章',
            '功能详述',
            f'多个需求的输入、过程、输出、异常流程描述偏泛化，未充分体现{focus}等差异化业务逻辑。',
            '终稿按每条需求补充关联模块、处理过程、输出结果、异常流程和质量关注点，保持需求 ID 范围不变。',
            rd_owner,
        ),
        ReviewIssue(
            '2',
            draft_title,
            '第4章/验收标准',
            '安全与测试约束',
            f'{_nonfunctional_focus(workbook.requirements)}等内容缺少安全、测试和边界条件验收口径。',
            '终稿补充安全性、可靠性、性能效率、兼容性和测试验证要求，并在条目中明确异常与边界条件。',
            _join_people(security_owner, test_owner),
        ),
        ReviewIssue(
            '3',
            draft_title,
            '术语/参考资料',
            '接口契约与评审追踪',
            '接口清单、Demo、设计稿、策略说明和评审记录之间的引用关系不够完整，后续研发联调和测试追踪成本较高。',
            '终稿补充依赖的接口契约、输入基础表和同级评审记录，确保初稿、评审表、终稿可追溯。',
            module_project_manager,
        ),
    ]
    return ReviewPlan(
        issues=issues,
        pending_problem='评审意见已全部按终稿修订闭环。',
        pending_solution=f'对应 {final_version} 终稿已按上述评审意见完成修订，问题状态均为关闭。',
        conclusion=(
            f'本次评审从项目、研发、安全和测试视角发现 {len(issues)} 项需完善意见。'
            f'{final_version} 终稿已细化功能条目的关联模块、处理过程、异常流程、安全与测试验收口径，'
            '需求范围与初稿保持一致，同意通过评审。'
        ),
        tracking=f'上述 {len(issues)} 项问题均已在对应 {final_version} 终稿中修订闭环；接口契约、Demo、设计稿如有更新，由责任人同步维护。',
        suggestion='建议研发继续补齐接口清单、Demo、策略说明和性能测试证据，并在后续评审中补充页号、提出人与签字信息。',
    )


def customer_scenario_description(requirement: RequirementEntry, *, final: bool) -> str:
    if not final:
        return first_sentence(requirement.user_description) or first_sentence(requirement.rd_description) or f'围绕“{requirement.title}”提供对应场景支撑。'
    profile = _requirement_profile(requirement)
    return f'背景：{profile["scenario"]} 使用目标：{profile["goal"]} 成功标准：{profile["success"]}'


def customer_requirement_description(requirement: RequirementEntry, *, scenario_id: str, final: bool) -> str:
    if not final:
        summary = first_sentence(requirement.user_description) or first_sentence(requirement.rd_description)
        role_goal = summary or f'用户围绕“{requirement.title}”发起操作，希望核心能力可用且结果清晰。'
        return f'角色与目标：{role_goal} 业务价值：支撑该需求在版本内清晰落地，并降低理解与执行成本。 归属场景（场景编号）：{scenario_id}'
    profile = _requirement_profile(requirement)
    return (
        f'角色与目标：{profile["scenario"]} '
        f'业务价值：{profile["goal"]} '
        f'验收关注：{profile["success"]} '
        f'归属场景（场景编号）：{scenario_id}'
    )


def product_requirement_details(requirement: RequirementEntry, *, final: bool) -> dict[str, str]:
    if not final:
        return {
            'related_module_name': '无',
            'feature_description': _product_feature_description(requirement),
            'related_description': '',
            'acceptance': ' '.join(part for part in requirement.acceptance.splitlines() if part),
            'input_desc': first_sentence(requirement.user_description or requirement.rd_description) or f'用户围绕“{requirement.title}”发起操作。',
            'process_desc': f'系统围绕“{requirement.title}”执行对应界面展示、状态处理或配置生效流程。',
            'output_desc': f'用户可以围绕“{requirement.title}”稳定完成目标操作，并获得一致反馈。',
            'exception_desc': '异常场景下应给出可理解提示，并保持基础能力可用。',
            'layout_desc': '沿用现有页面/接口形态，详细图示按研发设计稿和联调结果补充。',
            'quality_desc': '兼容性、易用性与可靠性要求遵循当前版本基线。',
        }
    profile = _requirement_profile(requirement)
    acceptance = ' '.join(part for part in requirement.acceptance.splitlines() if part)
    if profile['success'] and profile['success'] not in acceptance:
        acceptance = f'{acceptance} 评审后补充验收关注：{profile["success"]}'.strip()
    return {
        'related_module_name': profile['related'],
        'feature_description': f'围绕客户需求 {requirement.user_requirement_id} 与研发需求 {requirement.rd_requirement_id}，落实“{requirement.title}”能力。{profile["goal"]}',
        'related_description': f'{profile["related"]}需与模块保持状态、接口、权限和界面反馈一致；涉及外部依赖时按设计稿、接口文档或联调结果闭环。',
        'acceptance': acceptance,
        'input_desc': profile['scenario'],
        'process_desc': profile['process'],
        'output_desc': profile['success'],
        'exception_desc': profile['exception'],
        'layout_desc': '沿用现有页面或接口形态；涉及界面更新、模型列表、法律信息、模型切换和索引状态时，以设计稿/接口文档/联调结果作为终稿依据。',
        'quality_desc': profile['quality'],
    }


def product_quality_rows(module_name: str, *, final: bool) -> list[tuple[str, str, str]]:
    if not final:
        return [
            ('性能效率', f'{module_name}相关功能在高频交互场景下应保持流畅。', '重点关注切换、刷新、配置生效与状态同步时延。'),
            ('安全性', '接口、配置与入口调整不得突破现有权限边界。', '避免因开放能力或隐藏策略导致越权操作。'),
            ('可靠性', '连续调用、重复切换和状态频繁变化场景下应保持结果一致。', '异常时需保留基本能力并给出清晰反馈。'),
            ('兼容性（接口）', '界面能力应兼容当前主题和显示环境；接口能力兼容既有集成方式。', '命名、参数与返回格式保持稳定。'),
            ('可维护性', '配置项、接口项和状态项应遵循模块现有组织方式。', '便于后续版本继续演进与定位问题。'),
            ('可移植性', '输出内容应适配项目当前目标版本与交付环境。', '环境切换时不依赖额外人工修正。'),
            ('易用性', '入口位置、文案和反馈方式应尽量延续既有使用习惯。', '确保用户能快速理解状态并完成目标操作。'),
        ]
    return [
        ('性能效率', f'{module_name}的配置、索引、模型切换和调用能力在高频交互下保持可感知流畅；涉及明确指标的条目按基础表指标验收。', '如查询时延、资源占用、语音响应时延等指标需提供测试证据。'),
        ('安全性', '数据隔离、法律信息、免责声明、模型访问和 SDK 调用不得突破权限边界。', '重点验证跨应用数据不可访问、责任提示、接口错误码和敏感能力边界。'),
        ('可靠性', '模型安装/更新/卸载、开关、优先级、索引任务和自动切换在异常场景下保持状态一致。', '异常时保留上一次有效配置或明确提示，不产生静默失败。'),
        ('兼容性（接口）', 'SDK、模型管理、跳转入口、设计稿入口和历史配置需兼容既有集成方式。', '参数、返回值、错误码和配置迁移需在接口文档或 Demo 中体现。'),
        ('可维护性', '模型分类、开关影响域、索引策略和自动切换规则应可配置、可定位、可追踪。', '便于后续版本扩展能力类型、策略条件和体验优化文案。'),
        ('可移植性', '端侧能力加载、资源回收和本地/云端策略应适配目标交付环境。', '环境差异不得导致核心能力不可用。'),
        ('易用性', '入口、文案、状态和异常提示保持一致，降低用户理解成本。', '关键文案、分类、法律信息和切换提示需经体验走查。'),
    ]


def _requirement_profile(requirement: RequirementEntry) -> dict[str, str]:
    title = requirement.title
    haystack = f'{requirement.title}\n{requirement.rd_description}\n{requirement.user_description}\n{requirement.acceptance}'
    base_summary = first_sentence(requirement.user_description) or first_sentence(requirement.rd_description) or f'用户希望“{title}”能力可用、反馈清晰。'
    related = requirement.owner_module if requirement.owner_module and requirement.owner_module != '桌面环境' else 'AI子系统'
    profile = {
        'scenario': base_summary,
        'goal': f'围绕“{title}”提升操作稳定性、可理解性与体验一致性。',
        'success': first_sentence(requirement.acceptance) or '功能按预期生效，异常场景下给出可理解提示。',
        'related': related,
        'process': f'系统围绕“{title}”执行对应界面展示、状态处理或配置生效流程。',
        'exception': '异常场景下应给出可理解提示，并保持基础能力可用。',
        'quality': '重点验证配置一致性、异常提示、兼容性和高频操作响应。',
    }
    if _contains(haystack, ('向量', '隔离', '数据库')):
        profile.update({
            'scenario': '多应用共用向量能力时，需要保证应用数据隔离，并能判断向量数据生成与查询性能是否达标。',
            'goal': '不同应用的向量数据在逻辑或物理层面隔离，查询延迟与资源占用可度量。',
            'related': '向量数据库、权限隔离、性能监控',
            'process': '系统按应用身份创建或访问独立向量数据空间，执行查询时校验应用边界，并记录查询延迟、资源占用和生成链路指标。',
            'exception': '应用标识缺失、越权访问、索引损坏或资源不足时，应拒绝访问并给出安全提示。',
            'quality': '安全评审关注隔离边界；测试需覆盖跨应用访问、性能基准和资源上限。',
        })
    if _contains(haystack, ('模型管理', '软件商店', '下载安装', '更新', '卸载', '算力卡', '拔除')):
        profile.update({
            'scenario': '用户通过软件商店下载、安装、更新或卸载模型，并在硬件或模型状态变化时需要清楚知道模型是否可用。',
            'goal': '模型操作和状态变化在系统中保持一致，模型生效规则清晰。',
            'related': '软件商店、模型管理、模型状态同步、硬件状态检测',
            'process': '系统监听模型操作和硬件状态，刷新模型列表、可用状态和失效提醒，并按推理状态决定立即或延后生效。',
            'exception': '下载失败、更新失败、卸载中断或硬件状态不可识别时，应保留原状态并提示用户处理方式。',
            'quality': '重点验证模型状态同步、推理中变更、硬件状态提醒、可用模型兜底和列表展示一致性。',
        })
    if _contains(haystack, ('法律信息', '免责声明', '敏感词', '安全')):
        profile.update({
            'scenario': '用户使用本地内容、模型处理或相关能力时，需要看到法律信息和免责声明，理解结果责任边界。',
            'goal': '提供统一法律信息入口，免责声明覆盖相关风险提示。',
            'related': '法律信息、免责声明、安全合规',
            'process': '系统在明确入口展示法律信息和免责声明，并在相关能力处保持责任边界一致。',
            'exception': '法律信息加载失败时，应提示暂不可用并避免用户误认为无免责声明。',
            'quality': '安全评审需确认免责声明覆盖场景、入口可发现性和文案一致性。',
        })
    if _contains(haystack, ('界面', 'UI', '图标', '按钮', '设计稿', '文案', '体验')):
        profile.update({
            'scenario': '用户在界面中管理能力、查看状态或理解功能定位时，需要明确的分类、入口、文案和反馈。',
            'goal': '界面分类、图标、按钮、配置项和跳转规则统一，降低理解成本。',
            'related': '界面、设计稿、体验文案、配置入口',
            'process': '系统按设计稿渲染界面与状态，依据能力来源展示安装、配置、更新、卸载、导入或说明入口。',
            'exception': '设计稿链接不可访问、来源未知或跳转失败时，应降级展示已有能力并提示原因。',
            'quality': '产品与测试需按设计稿走查空列表、有数据、异常跳转和文案展示。',
        })
    if _contains(haystack, ('索引', 'GPU', 'CPU', '闲时', '锁屏', '屏保', '显存')):
        profile.update({
            'scenario': '系统执行后台索引或资源密集任务时，需要兼顾用户使用状态、GPU/CPU 资源占用和任务连续性。',
            'goal': '索引任务遵循忙闲状态、资源配额和中断规则，避免影响桌面体验。',
            'related': '索引服务、GPU资源、CPU资源、系统状态检测',
            'process': '系统区分不同索引任务，按忙闲状态、显存、CPU 配额和任务队列执行。',
            'exception': '显存不足、CPU 配额异常、任务中断或系统状态判断失败时，应记录原因并避免后台反复重试。',
            'quality': '测试需覆盖忙时、闲时、锁屏、屏保、单核/多核 CPU 口径和资源不足场景。',
        })
    if _contains(haystack, ('开关', '优先级', '影响域', '覆盖范围')):
        profile.update({
            'scenario': '用户希望统一控制能力开关和优先级，同时明确不受控制的例外对象。',
            'goal': '开关和优先级覆盖范围清晰，例外对象有明确提示。',
            'related': '开关控制、优先级策略、状态同步',
            'process': '系统按能力类型应用开关和优先级策略，对例外对象只展示提示不执行控制。',
            'exception': '能力类型识别失败或优先级冲突时，应阻止保存并提示用户调整。',
            'quality': '重点验证覆盖范围、例外提示、历史配置兼容和优先级保存。',
        })
    if _contains(haystack, ('加载', '卸载', '端侧', '响应时间')):
        profile.update({
            'scenario': '用户首次使用或长时间闲置后再次使用端侧能力时，需要缩短等待时间，并理解加载、卸载和再次加载状态。',
            'goal': '加载和卸载机制按策略执行，关键响应时间满足基础表要求。',
            'related': '端侧加载、资源回收、响应时延',
            'process': '系统按空闲策略释放资源，再次调用时触发加载流程并反馈状态；关键调用路径需优先保证响应时延。',
            'exception': '加载失败、超时、资源不足或策略链接不可达时，应提示并保留可用能力。',
            'quality': '测试需覆盖首次加载、闲置卸载、再次加载、响应时延和资源回收。',
        })
    if _contains(haystack, ('切换', '断网', '超时', '密钥', '秘钥', '不可用')):
        profile.update({
            'scenario': '当前能力因密钥错误、关闭、不可用、请求超时或网络状态导致不可服务时，用户希望系统按规则切换到可用能力，且不误切换。',
            'goal': '自动切换策略区分触发条件和不触发条件，保证连续性并避免错误兜底。',
            'related': '自动切换、网络状态检测、优先级策略',
            'process': '系统在调用开始前检查可用性、密钥、网络和优先级，满足触发条件时选择可用能力；调用过程中的异常按不触发规则处理。',
            'exception': '全部能力不可用、切换失败或优先级配置缺失时，应保持原状并明确提示。',
            'quality': '测试需覆盖触发/不触发边界、网络时序、优先级选择、提示文案和切换时延。',
        })
    if _contains(haystack, ('SDK', '接口', 'Function Call', 'node.js', 'js 接口', 'Demo', 'demo')):
        profile.update({
            'scenario': '第三方应用、内部模块或研发调用方需要通过稳定接口调用能力，并获得清晰的参数、返回值、错误码和示例。',
            'goal': '接口能力清单清晰，调用方式、参数、返回值和 Demo 可支撑研发联调与测试验收。',
            'related': 'SDK、接口契约、推理服务、模型服务、Demo',
            'process': '系统通过统一接口暴露调用入口，按任务优先级或模型配置调度资源，并在接口文档中明确参数、错误码、返回结构和示例。',
            'exception': '模型不可用、参数非法、任务超时、资源不足或接口版本不匹配时，应返回明确错误码且不影响其他任务。',
            'quality': '研发需补充接口清单和 Demo；测试需覆盖同步/异步、异常码、并发任务和接口一致性。',
        })
    return profile


def _product_feature_description(requirement: RequirementEntry) -> str:
    summary = ' '.join(part for part in requirement.rd_description.splitlines() if part) or first_sentence(requirement.user_description)
    return f'围绕客户需求 {requirement.user_requirement_id} 与研发需求 {requirement.rd_requirement_id}，落实“{requirement.title}”能力。 {summary}'.strip()


def _focus_summary(requirements: list[RequirementEntry]) -> str:
    labels = []
    checks = [
        ('模型配置/管理', ('模型配置', '模型管理', '模型开关', '优先级')),
        ('向量数据隔离', ('向量', '隔离')),
        ('SDK 接口', ('SDK', '接口', 'Function Call', 'Demo', 'demo')),
        ('界面与体验', ('界面', 'UI', '图标', '文案', '体验')),
        ('索引与资源占用', ('索引', 'GPU', 'CPU', '显存')),
        ('自动切换策略', ('切换', '断网', '超时')),
        ('安全合规', ('法律信息', '免责声明', '安全')),
    ]
    haystack = '\n'.join(f'{item.title}\n{item.rd_description}\n{item.acceptance}' for item in requirements)
    for label, keywords in checks:
        if _contains(haystack, keywords):
            labels.append(label)
    return '、'.join(labels[:5]) or '核心需求'


def _nonfunctional_focus(requirements: list[RequirementEntry]) -> str:
    labels = []
    haystack = '\n'.join(f'{item.title}\n{item.rd_description}\n{item.acceptance}' for item in requirements)
    checks = [
        ('数据隔离', ('隔离', '越权')),
        ('性能指标', ('1秒', '资源占用', 'CPU', 'GPU', '显存', '延迟')),
        ('异常提示', ('异常', '失败', '超时', '不可用')),
        ('安全免责声明', ('免责声明', '法律信息', '敏感词')),
        ('兼容与接口稳定性', ('接口', 'Demo', 'demo', '参数', '返回值')),
    ]
    for label, keywords in checks:
        if _contains(haystack, keywords):
            labels.append(label)
    return '、'.join(labels[:5]) or '安全、性能、可靠性和异常提示'


def _contains(text: str, keywords: tuple[str, ...]) -> bool:
    return any(keyword in text for keyword in keywords)


def _join_people(*names: str) -> str:
    cleaned = []
    for name in names:
        value = name.strip()
        if value and value not in cleaned:
            cleaned.append(value)
    return '、'.join(cleaned)
