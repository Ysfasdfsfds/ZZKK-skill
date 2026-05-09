# ZZKK-PRD

`ZZKK-PRD` 是一个面向 Claude Code 的本地 skill，用来把一份原始《基础信息表》先补充整理，再生成成套需求文档交付物。

它会先基于用户提供的原始基础信息表：

- 补充 `用户需求描述`
- 完善 `验收标准`
- 另存一份新的基础信息表副本

然后继续生成 6 个交付物：

- `<模块名称>-客户需求说明书_<初稿版本号>.docx`
- `<模块名称>-客户需求说明书_<终稿版本号>.docx`
- `<模块名称>-客户需求说明书_同级评审会议记录表.xlsx`
- `<模块名称>-模块级产品需求分析说明书_<初稿版本号>.docx`
- `<模块名称>-模块级产品需求分析说明书_<终稿版本号>.docx`
- `<模块名称>-模块级产品需求分析说明书_同级评审会议记录表.xlsx`

## 适合什么场景

这个 skill 适合下面这类工作流：

- 你已经有一份原始《基础信息表》
- 你手里有既定格式的客户/产品/评审模板
- 客户需求说明书使用简化版 DOCX 模板，只需要补“应用场景表”和“用户需求描述表”
- 你希望 Claude Code 按固定流程补充字段并批量生成交付物
- 你希望保留现有模板格式，而不是重新生成一套全新版式

## 快速开始

### 1. 安装到 Claude Code 本地 skills 目录

SSH：

```bash
mkdir -p ~/.claude/skills
git clone git@github.com:Ysfasdfsfds/ZZKK-skill.git ~/.claude/skills/ZZKK-PRD
```

HTTPS：

```bash
mkdir -p ~/.claude/skills
git clone https://github.com/Ysfasdfsfds/ZZKK-skill.git ~/.claude/skills/ZZKK-PRD
```

### 2. 安装 Python 依赖

```bash
python3 -m pip install -r ~/.claude/skills/ZZKK-PRD/requirements.txt
```

### 3. 在 Claude Code 中调用 skill

```text
/ZZKK PRD
```

安装后建议新开一个 Claude Code 会话，确认 skill 已被正确发现。

## 你需要准备什么

运行前需要准备：

1. 一份原始《基础信息表》
2. 一个《客户需求说明书》DOCX 模板
3. 一个《产品需求分析说明书》DOCX 模板
4. 一个《评审表》XLSX 模板
5. Claude Code 可用的本地 Python 3 环境

这个仓库 **不包含** 上述模板文件，也不包含真实工作簿样例。

## 原始基础信息表要求

原始基础信息表至少必须包含这些列：

- `用户需求id`
- `研发需求id`

为了实际生成文档，还需要这些列：

- `研发需求名称`
- `研发需求描述`

Skill 在运行时会继续检查：

- 是否需要新增 `用户需求描述`
- 是否需要新增 `验收标准`
- 是否需要对已有 `验收标准` 做完善

## Skill 会怎么和你交互

调用后，skill 会主动询问：

- 原始基础信息表路径
- 模块编号
- 模块名称
- 产线名称英文缩写
- 产品名称
- 初稿版本号
- 终稿版本号
- 3 个模板文件分别是什么角色
- 负责该模块的模块产品经理姓名（评审表主持人）
- 负责该模块的项目经理（评审表书记员）
- 评审员名单（应包括研发负责人、项目经理、产品级产品经理、安全负责人、测试负责人）
- 研发负责人姓名、产品级产品经理姓名、安全负责人姓名、测试负责人姓名（写入评审问题“提出人”）
- 对应 QA（评审表其他人员）
- 输出目录
- 是否覆盖现有文件

其中有几个固定规则：

- 模板角色 **不能依赖文件名猜测**
- 默认 **不覆盖** 已存在文件
- 原始基础信息表会先被补充、另存，再用于最终生成
- 文件编号以模板为准：客户说明书为 `产线名称英文缩写-模块ID-CRS-版本号`，产品说明书为 `产线名称英文缩写-产品名称-模块ID-PRD-版本号`
- “无/不涉及”类占位统一输出 `无`，不得用 `/` 代替
- 初稿、评审表、终稿必须形成流程闭环：先生成初稿，再基于初稿生成评审表，最后根据评审表形成终稿；终稿基于同一份补充后的基础信息表生成，并覆盖同一批需求 ID

## 手动 CLI 调试

如果你想在 Claude Code 之外排查问题，可以直接运行 `skill.py`。

### dry-run

```bash
python3 ~/.claude/skills/ZZKK-PRD/skill.py plan \
  --workbook "<基础信息表路径>" \
  --customer-template "<客户模板路径>" \
  --product-template "<产品模板路径>" \
  --review-template "<评审表模板路径>" \
  --module-code "<模块编号>" \
  --module-name "<模块名称>" \
  --product-line-code "<产线名称英文缩写>" \
  --product-name "<产品名称>" \
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
  --module-product-manager "<模块产品经理>" \
  --module-project-manager "<模块项目经理>" \
  --reviewers "<评审员名单>" \
  --rd-owner "<研发负责人>" \
  --product-owner "<产品级产品经理>" \
  --test-owner "<测试负责人>" \
  --security-owner "<安全负责人>" \
  --qa "<对应QA>" \
  --output-dir "<输出目录>"
```

### 生成文件

```bash
python3 ~/.claude/skills/ZZKK-PRD/skill.py run \
  --workbook "<基础信息表路径>" \
  --customer-template "<客户模板路径>" \
  --product-template "<产品模板路径>" \
  --review-template "<评审表模板路径>" \
  --module-code "<模块编号>" \
  --module-name "<模块名称>" \
  --product-line-code "<产线名称英文缩写>" \
  --product-name "<产品名称>" \
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
  --module-product-manager "<模块产品经理>" \
  --module-project-manager "<模块项目经理>" \
  --reviewers "<评审员名单>" \
  --rd-owner "<研发负责人>" \
  --product-owner "<产品级产品经理>" \
  --test-owner "<测试负责人>" \
  --security-owner "<安全负责人>" \
  --qa "<对应QA>" \
  --output-dir "<输出目录>"
```

如果你确认允许覆盖现有文件，再追加：

```bash
--overwrite
```

### 校验输出

```bash
python3 ~/.claude/skills/ZZKK-PRD/skill.py verify \
  --workbook "<基础信息表路径>" \
  --customer-template "<客户模板路径>" \
  --product-template "<产品模板路径>" \
  --review-template "<评审表模板路径>" \
  --module-code "<模块编号>" \
  --module-name "<模块名称>" \
  --product-line-code "<产线名称英文缩写>" \
  --product-name "<产品名称>" \
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
  --module-product-manager "<模块产品经理>" \
  --module-project-manager "<模块项目经理>" \
  --reviewers "<评审员名单>" \
  --rd-owner "<研发负责人>" \
  --product-owner "<产品级产品经理>" \
  --test-owner "<测试负责人>" \
  --security-owner "<安全负责人>" \
  --qa "<对应QA>" \
  --output-dir "<输出目录>"
```

校验会同时检查：

- 评审表“工作产品/标识”指向对应说明书初稿，而不是终稿
- 同一类说明书的初稿与终稿包含同一批需求 ID
- 终稿版本记录包含对应评审记录编号，体现由初稿经评审形成终稿
- 6 个输出文件存在且文档编号/评审记录编号正确

## 模板兼容性说明

这个 skill 不是一个“任意模板都能直接适配”的通用文档引擎。

当前版本依赖这些前提：

- DOCX 模板中存在既定标题、表格结构和标签文本
- 模板中的加粗标签样式需要能够被原结构复用
- XLSX 评审表模板使用固定单元格位置写入记录编号、项目名、会议日期等信息

如果模板结构和当前实现假定差异过大，生成可能失败，或者需要你再做针对性调整。

## 常见问题

### 缺少 `openpyxl`

安装依赖：

```bash
python3 -m pip install -r ~/.claude/skills/ZZKK-PRD/requirements.txt
```

### 缺少工作簿列

请先确认原始基础信息表至少包含：

- `用户需求id`
- `研发需求id`
- `研发需求名称`
- `研发需求描述`

### 模板不兼容

如果报错显示缺少表格、标签或定位失败，通常说明你的模板结构与当前 skill 假定的不一致，需要调整模板，或修改 `lib/` 下的处理逻辑。

### 输出文件已存在

默认不会覆盖。只有在你明确希望覆盖时，才在 CLI 中传入 `--overwrite`，或在 Claude Code 交互里明确允许覆盖。

## 仓库结构

```text
ZZKK-PRD/
├── SKILL.md
├── README.md
├── LICENSE
├── requirements.txt
├── skill.py
└── lib/
    ├── doc_numbering.py
    ├── docx_ops.py
    ├── pipeline.py
    ├── workbook_reader.py
    └── xlsx_review_sheet_ops.py
```

## License

MIT
