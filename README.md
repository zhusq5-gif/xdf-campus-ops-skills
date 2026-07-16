# 新东方重庆学校校区运营 Skills

本仓库把儿童素养智学部可重复、可验证的日常管理动作沉淀为 Codex Skill。首个 Skill 是 `xdf-plan-campus-capacity`：读取脱敏 Excel，清洗并校验数据，按未来 12 个月、峰值时段和班型测算教师与教室产能，输出可追溯的管理建议。

## 第一性原理

产能决策只由可复算的数据、规则和约束产生：

1. 先把业务导出转换为统一数据合同。
2. 用确定性公式计算学员、班级、教师和教室需求。
3. 仅在数据与规则通过校验时生成招聘或扩租建议。
4. 用智能体解释证据、比较情景和组织沟通，不让模型凭语言直觉编造数字。

## 当前交付

- 仓库级 Skill：`.agents/skills/xdf-plan-campus-capacity/`
- 默认规则：`config/capacity-planning.json`
- 合成评测数据：`evals/fixtures/`
- 本机业务数据：`data/local/`，已被 Git 忽略
- 演示输出：`outputs/demo/`，已被 Git 忽略

## 已合并的候选管理动作

`秋季纳新名额测算` 的完整源项目已迁入 Git 忽略的 `data/local/autumn-enrollment-quota/`，可复用口径、数据合同和风险边界已脱敏沉淀为 [候选 Skill 发现记录](docs/discovery/AUTUMN_ENROLLMENT_QUOTA.md)。

当前不将它直接封装为第二个 Skill：源实现含真实姓名、单次手工覆盖和固定日期/单元格，且“续班预估占用”不等于“可纳新名额”。待补齐可售容量与预留名额口径、建立合成评测后再封装。

## 快速开始

```bash
python3 -m pip install -r requirements.txt
python3 .agents/skills/xdf-plan-campus-capacity/scripts/plan_capacity.py \
  --input evals/fixtures/synthetic-capacity-input.xlsx \
  --config config/capacity-planning.json \
  --template .agents/skills/xdf-plan-campus-capacity/assets/capacity-output-template.xlsx \
  --output-dir outputs/demo \
  --backtest
```

运行后生成：

- `campus-capacity-result.xlsx`
- `management-recommendations.md`
- `message-payload.json`

## 验证

```bash
python3 -m unittest discover -s tests -v
python3 tools/scan_sensitive_data.py .
python3 tools/validate_skill.py .agents/skills/xdf-plan-campus-capacity
```

在接入真实数据前，先阅读 [项目规格](docs/PROJECT_SPEC.md)、[数据合同](docs/DATA_CONTRACT.md)、[评测与验收](docs/EVALUATION.md) 和 [安全策略](SECURITY.md)。

## 明确边界

- 首期不自动招聘、不自动签约、不直接向飞书或钉钉发送消息。
- 首期不解决具体排课；只测算峰值时段的资源容量。
- 业务导出默认只在本机处理。GitHub 只保存合成数据和通过专项复核的最小样本。
- 第二个管理动作形成独立触发条件和数据合同后，再评估拆分 Skill 合集；需要团队安装时再封装 Plugin。
