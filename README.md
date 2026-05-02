# ZZKK-PRD

`ZZKK-PRD` 是一个面向 Claude Code 的本地 skill，用来把一份原始《基础信息表》整理为一份补充后的基础信息表副本，并进一步生成 6 个交付物：

- `<模块名称>-客户需求说明书_<初稿版本号>.docx`
- `<模块名称>-客户需求说明书_<终稿版本号>.docx`
- `<模块名称>-客户需求说明书_同级评审会议记录表.xlsx`
- `<模块名称>-模块级产品需求分析说明书_<初稿版本号>.docx`
- `<模块名称>-模块级产品需求分析说明书_<终稿版本号>.docx`
- `<模块名称>-模块级产品需求分析说明书_同级评审会议记录表.xlsx`

在生成这 6 个文件之前，skill 会先基于用户提供的原始基础信息表：

- 补充 `用户需求描述`
- 完善 `验收标准`
- 另存一份新的基础信息表副本

## 适用前提

你需要准备：

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

同时，为了实际生成文档，还需要这些列：

- `研发需求名称`
- `研发需求描述`

Skill 会在执行时继续检查：

- 是否需要新增 `用户需求描述`
- 是否需要新增或完善 `验收标准`

## 安装方式

先把仓库 clone 到 Claude Code 的本地 skills 目录：

```bash
mkdir -p ~/.claude/skills
git clone <your-repo-url> ~/.claude/skills/ZZKK-PRD
```

然后安装 Python 依赖：

```bash
python3 -m pip install -r ~/.claude/skills/ZZKK-PRD/requirements.txt
```

安装完成后，重新打开 Claude Code 会话，或在新会话里确认 skill 已被发现。

## 使用方式

在 Claude Code 中通过 `/ZZKK PRD` 调用这个 skill。

Skill 会主动询问这些内容：

- 原始基础信息表路径
- 模块编号
- 模块名称
- 初稿版本号
- 终稿版本号
- 用户提供的 3 个模板文件分别是什么角色
- 输出目录
- 是否覆盖现有文件

其中：

- 模板角色 **不能依赖文件名猜测**
- 默认 **不覆盖** 已存在文件
- 原始基础信息表会先被补充、另存，再用于最终生成

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
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
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
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
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
  --draft-version "<初稿版本号>" \
  --final-version "<终稿版本号>" \
  --output-dir "<输出目录>"
```

## 模板兼容性说明

这个 skill 不是一个“任意模板都能直接适配”的通用文档引擎。

当前版本依赖这些前提：

- DOCX 模板中存在既定标题、表格结构和标签文本
- 模板中的加粗标签样式需要能够被原结构复用
- XLSX 评审表模板使用固定单元格位置写入记录编号、项目名、会议日期等信息

如果模板结构和当前实现假定差异过大，生成会失败，或者生成结果需要你再做针对性调整。

## 常见问题

### 1. 缺少 `openpyxl`

安装依赖：

```bash
python3 -m pip install -r ~/.claude/skills/ZZKK-PRD/requirements.txt
```

### 2. 缺少工作簿列

请先确认原始基础信息表至少包含：

- `用户需求id`
- `研发需求id`
- `研发需求名称`
- `研发需求描述`

### 3. 模板不兼容

如果报错显示缺少表格、标签或定位失败，通常说明你的模板结构与当前 skill 假定的不一致，需要调整模板，或修改 `lib/` 下的处理逻辑。

### 4. 输出文件已存在

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
