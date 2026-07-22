# 重庆新东方校区运营 Skill 合集

面向校区管理者的可复算、可审计运营工具集。当前包含两个独立 Skill：

- `xdf-normalize-teacher-schedule`：教师课表占档与班级容量规整。
- `xdf-plan-campus-capacity`：未来 12 个月教师与教室产能规划。

两个 Skill 都以 Python 确定性脚本计算数字。AI 只负责交互、解释、生成受限规则补丁和组织人工确认，不凭语言直觉补数。

代码按 [MIT License](LICENSE) 开源。许可证不授予真实业务数据、学员/教师信息或新东方品牌素材的使用权；这些内容不包含在仓库中，也不得上传。

## 这能干啥

| Skill | 你提供 | 你得到 |
|---|---|---|
| 教师课表占档规整 | 教务系统导出的班级明细，可选人工回传课表 | 教师 × 时段课表、容量总览/明细、冲突与待确认清单 |
| 校区产能规划 | 在读、续费率、招生目标、班容、教师与教室供给 | 12 个月三情景预测、教师/教室缺口、招聘扩租建议草稿 |

合集以标准多 Skill 目录打包：

```text
.agents/skills/
├── xdf-normalize-teacher-schedule/
│   ├── SKILL.md
│   ├── scripts/
│   └── references/
└── xdf-plan-campus-capacity/
    ├── SKILL.md
    ├── scripts/
    ├── references/
    └── assets/
```

支持 Skill 的 Agent 可直接发现 `.agents/skills/`；不使用 Codex 时，可直接运行各 Skill 的 Python CLI。

## 交互式使用引导

1. 说明目标：“整理教师课表”或“测算未来 12 个月产能”。
2. 提供本地 `.xlsx`，并确认数据日期、校区范围和使用目的。
3. AI 核对工作表和列映射，不猜测缺失字段。
4. 脚本先生成基准结果，AI 解释 `ready`、`needs_confirmation` 或 `blocked`。
5. 课表结果可人工修改后回传，再明确选择“规则优先”或“模板优先”重跑。
6. 如需改业务规则，直接用自然语言提出。AI 展示“旧值 → 新值”后，只写入本地覆盖；说“恢复初始规则”即可清空覆盖并重跑。
7. 管理结论、排课调整、招聘、扩租或消息发送始终由人确认。

### 课表 Skill 的触发条件

以下请求会触发 `xdf-normalize-teacher-schedule`：

- 首次上传教务班级明细，要求整理教师占档、空段、班型供给或容量差额。
- 把第一版课表人工修改后回传，要求识别关班、新班、换教师或移动时段。
- 要求在“规则优先”和“模板优先”之间切换重跑。
- 用自然语言修改允许覆盖的业务规则，或要求恢复初始规则。
- 复核课表的冲突、待确认、模板对比和运行状态。

首次生成或重跑必须同时具备原始班级明细 `.xlsx`、数据日期和校区/管理单元。只解释已有结果时可读取结果工作簿与 `schedule-review.json`，但不能据此重算。未来 12 个月教师/教室缺口改用产能 Skill；自动排课、修改教务系统或范围外课程不触发本 Skill。

## 快速开始

需求：Python 3.11+ 和可用的 `pip`。

```bash
git clone https://github.com/zhusq5-gif/xdf-campus-ops-skills.git
cd xdf-campus-ops-skills
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

用合成数据运行课表 Skill：

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/normalize_schedule.py \
  --input evals/fixtures/synthetic-schedule-input.xlsx \
  --config config/teacher-schedule.json \
  --campus 合成校区 \
  --data-as-of 2026-07-22 \
  --output-dir outputs/demo-schedule
```

用合成数据运行产能 Skill：

```bash
python3 .agents/skills/xdf-plan-campus-capacity/scripts/plan_capacity.py \
  --input evals/fixtures/synthetic-capacity-input.xlsx \
  --config config/capacity-planning.json \
  --template .agents/skills/xdf-plan-campus-capacity/assets/capacity-output-template.xlsx \
  --output-dir outputs/demo \
  --backtest
```

`--backtest` 仅用于历史回测，实时决策不得使用。

## 规则优先模式（稳定）

`rule_first` 是课表 Skill 的默认模式。教务导出和确定性业务规则优先，人工模板中的删班、换教师、换时段只生成差异提示。

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/normalize_schedule.py \
  --input <班级明细.xlsx> \
  --config config/teacher-schedule.json \
  --campus <完整校区名称> \
  --data-as-of YYYY-MM-DD \
  --manual-template <人工修改后课表.xlsx> \
  --template-policy rule_first \
  --output-dir <输出目录>
```

适用于首次运行、教务快照复核、模板修改尚未确认或希望避免误删班级的场景。

`rule_first` 只保护教务正式班的位置和保留状态；回传模板中的合规计划新班仍会以当前人数 0 纳入。若只是演练，不希望计划新班参与容量和冲突计算，应先从模板中移除。

## 模板优先模式（个性化）

`template_first` 适用于使用者已在本 Skill 输出课表中确认关班、换教师、换时段或计划新班后的个性化结果。

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/normalize_schedule.py \
  --input <班级明细.xlsx> \
  --config config/teacher-schedule.json \
  --campus <完整校区名称> \
  --data-as-of YYYY-MM-DD \
  --manual-template <人工修改后课表.xlsx> \
  --template-policy template_first \
  --output-dir <输出目录>
```

模板优先只能改变合规正式班的“是否保留、教师、时段”；不能绕过输入安全、年级/科目范围、标准课次、班号、容量和季节约束。已回传季节中，模板未出现的源班不会入表。

### 第一版产出回传后的处理闭环

1. AI 复用原始班级明细、数据日期和校区范围，确认回传文件来自本 Skill。
2. 脚本读取正式班号与 `新班...`，不把底色、批注或任意文本当成规则。
3. 对比“数据源有/模板无、模板有/数据源无、同班号教师或时段变化”。
4. 使用者明确选择模式；未明确时保持 `rule_first`。
5. `template_first` 只应用通过硬约束的正式班保留、教师和时段，并重新执行冲突与容量计算。
6. 新结果列明已应用、未应用和待确认的变化。

模板改动不会自动变成永久业务规则。人工移动或删班是“模板记录”，只在继续传入该模板时生效；“以后都按新班容/新时段规则处理”才会由 AI 转成受限本地 JSON 覆盖。覆盖文件显式参与后续运行，并可通过“恢复初始规则”清空。

## 提示词触发速查表

| 你可以说 | 触发的 Skill/动作 |
|---|---|
| “把这份班级明细整理成教师课表” | 运行 `xdf-normalize-teacher-schedule` 基准模式 |
| “这个班有两个老师，按主带课老师” | 校验多教师与主带课老师匹配 |
| “以系统规则为准重跑” | `--template-policy rule_first` |
| “以我改过的模板为准” | `--template-policy template_first` |
| “思维一二年级最大班容改为 17” | AI 生成受限本地规则补丁，校验后重跑 |
| “恢复初始规则” / “回到初始规则” | 清空本地规则覆盖并重跑 |
| “按未来 12 个月测算教师和教室缺口” | 运行 `xdf-plan-campus-capacity` |
| “用历史周期回测” | 产能 Skill 增加 `--backtest` |
| “只检查数据，不给招聘扩租结论” | 说明只读评估，不执行外部动作 |

## 产能处理规则速查表

课表 Skill 的班级容量默认值：

| 科目 | 年级 | 最低开班人数 | 最大班容 |
|---|---:|---:|---:|
| 思维 / 博学 | 1–2 | 8 | 16 |
| 思维 / 博学 | 3–6 | 10 | 20 |
| 双语 STEAM | 1–2 | 10 | 18 |
| 双语 STEAM | 3–6 | 12 | 20 |

- 允许超售；差额 = `最大班容 - 当前人数`，可为负数。
- 未达开班线为蓝色加粗，已开未满为绿色加粗，满班/超售为红色加粗。
- 同一教师同一时段多班：人数唯一较大的班临时入表；最高人数并列时必须人工确认。

产能规划 Skill 的默认处理链：

| 环节 | 默认口径 |
|---|---|
| 续费后在读 | `在读 × min(1, 基准续费率 × 情景系数)` |
| 新增学员 | `本期目标 × (往期实际 / 往期目标) × 情景系数` |
| 三种情景 | 保守 `0.9`、基准 `1.0`、进取 `1.1` |
| 班级数 | 先满足最大班容，再在最小班容可行范围内选最接近目标班容的班级数 |
| 教师需求 | `计划班级数 × 每班教师数` |
| 教室需求 | 按校区、月份、峰值时段、教室类型汇总 `班级数 × 每班教室数` |
| 缺口 | `max(0, 需求 - 供给)` |
| 建议门槛 | 只用基准情景，默认连续缺口 2 个月才生成建议 |
| 默认提前期 | 教师 2 个月，教室 6 个月 |

以上数值都来自版本化配置，业务调整后必须改配置、重跑测试和重算结果。

## 核心规范

1. **数字只由脚本计算**：AI 不补写缺失人数、续费率、班容、时段或供给。
2. **规则版本化**：字段映射、容量和阈值进入 `config/`，不散落在提示词中。
3. **数据日期可追溯**：输出保留数据日期、配置版本、输入哈希和证据键。
4. **安全失败则阻断**：Excel 公式、外部链接、疑似公式注入、缺列、关键重复或越界值不得带病计算。
5. **人工确认闭环**：不自动修改教务系统，不自动排课、招聘、扩租或发消息。
6. **最小化数据**：Git 只保存合成数据和经复核的规则；真实导出、输出、姓名、工号和可逆映射不得上传。
7. **结果可对抗验证**：交付前同时检查正常算例、恶意单元格、日期错位、容量冲突和模板重复位置。

## 深入阅读

- [教师课表 Skill](.agents/skills/xdf-normalize-teacher-schedule/SKILL.md)
- [课表数据合同](.agents/skills/xdf-normalize-teacher-schedule/references/data-contract.md)
- [课表业务规则](.agents/skills/xdf-normalize-teacher-schedule/references/business-rules.md)
- [课表输出规范](.agents/skills/xdf-normalize-teacher-schedule/references/output-spec.md)
- [校区产能 Skill](.agents/skills/xdf-plan-campus-capacity/SKILL.md)
- [产能数据合同](.agents/skills/xdf-plan-campus-capacity/references/data-contract.md)
- [产能业务规则](.agents/skills/xdf-plan-campus-capacity/references/business-rules.md)
- [管理输出规范](.agents/skills/xdf-plan-campus-capacity/references/management-output.md)
- [项目规格](docs/PROJECT_SPEC.md)
- [评测与对抗性验收](docs/EVALUATION.md)
- [安全策略](SECURITY.md)

## 功能对比表

| 能力 | 课表规整 | 产能规划 |
|---|---|---|
| 决策时间尺度 | 当前导出快照 | 未来 12 个月 |
| 核心维度 | 季节 × 教师 × 轮次/星期时段 | 情景 × 月份 × 校区 × 班型 × 峰值时段 |
| 主要输入 | 班级明细、可选回传课表 | 在读、续费、招生、班型、教师、教室 |
| 主要输出 | 教师课表、容量差额、冲突/待确认 | 三情景预测、资源缺口、管理建议草稿 |
| 人工覆盖 | `rule_first` / `template_first` | 修改版本化配置后重跑 |
| 自然语言规则覆盖 | 支持受限本地覆盖与一键恢复 | 暂不支持；必须修改版本化配置 |
| 外部执行 | 不修改教务系统，不自动调课 | 不自动招聘、扩租或发消息 |

## 错误处理

| 状态/错误 | 处理方式 |
|---|---|
| `ready` | 数据与规则通过；仍需管理者确认数据快照后的业务变化 |
| `needs_confirmation` | 已生成安全子集；查看待确认、冲突和模板对比后修改数据或规则再跑 |
| `blocked` | 只输出校验问题；不给出容量、招聘或扩租结论 |
| 缺工作表/列 | 按数据合同修复源表或显式字段映射，不做相似列名猜测 |
| 公式、外部链接、注入文本 | 阻断，将来源工作簿转为经复核的静态值后再运行 |
| 多教师无有效主带课老师 | 班级不入表，在教务源中补齐唯一主带课老师 |
| 模板同班号多位置/跨季移动 | 进入待确认，不会由 `template_first` 强制覆盖 |
| 本地规则覆盖哈希过期 | 阻断旧覆盖；先审核新初始配置，再生成新补丁 |

开发与发布前验收：

```bash
python3 -m unittest discover -s tests -v
python3 tools/scan_sensitive_data.py .
python3 tools/validate_skill.py .agents/skills/xdf-plan-campus-capacity
python3 tools/validate_skill.py .agents/skills/xdf-normalize-teacher-schedule
```

## 使用限制

- 课表 Skill 当前只支持小学 1–6 年级的思维（NL/SW）、双语 STEAM（ST）和博学（BX）。
- 当前可运行源表只覆盖暑假和秋季；寒假/春季口径虽有记录，没有源表时不会虚构输出。
- 复习班、补课班、定金班、取消班、S2/S3 和非标准临时班不进入正式课表。
- 产能 Skill 的三情景是资源决策辅助，不是对学员、续费或招生结果的保证。
- 建议中的“教师容量单位”不等于具体编制或劳动合同人数，需人工确认工时和岗位口径。
- 本合集不连接实时业务系统，不代替教务、人事、财务或租赁审批。
- 真实业务数据和生成输出必须位于 `data/local/` 或 `outputs/` 等 Git 忽略路径；上传 GitHub 前必须通过自动扫描和人工复核。
