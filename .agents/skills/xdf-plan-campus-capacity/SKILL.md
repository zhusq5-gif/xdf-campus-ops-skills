---
name: xdf-plan-campus-capacity
description: 校验和清洗新东方校区运营 Excel，基于在读学员、续费率、招生目标与往期达成、招生/开课时段、班容与搭班规则、教师和教室供给，按未来12个月与峰值时段测算三种情景下的教师/教室产能缺口，并生成可审计的招聘扩租建议、Excel 报表及飞书/钉钉消息载荷。用于校区产能规划、师资招聘测算、教室扩租分析、经营例会资源预测或相关表格复算；不要用于具体排课、自动招聘、自动租赁或未经人工确认的消息发送。
---

# 校区产能规划

把业务数据转换为可复算的产能判断。始终让脚本计算数字，只根据脚本输出解释结果。

## 工作流

1. 确认输入是脱敏 `.xlsx`，并确认业务数据存放在 Git 忽略目录。
2. 读取 [data-contract.md](references/data-contract.md)，核对源表和配置映射；不要猜测缺失或近似字段。
3. 读取 [business-rules.md](references/business-rules.md)，确认续费、招生、班级、教师和教室口径。
4. 使用版本化配置运行确定性脚本：

```bash
python3 .agents/skills/xdf-plan-campus-capacity/scripts/plan_capacity.py \
  --input <input.xlsx> \
  --config config/capacity-planning.json \
  --template .agents/skills/xdf-plan-campus-capacity/assets/capacity-output-template.xlsx \
  --output-dir <output-dir>
```

历史回测增加 `--backtest`；实时决策不要使用该参数。

5. 检查脚本退出码和 `message-payload.json` 的 `status`：
   - `ready`：可以解释结果并组织人工审批。
   - `blocked`：只报告校验问题，不得输出招聘或扩租结论。
6. 读取 [management-output.md](references/management-output.md)，把证据化建议交付给管理者。

## 强制边界

- 拒绝含公式、外部链接、疑似公式注入、关键重复或越界比例的输入。
- 不用模型补齐人数、续费率、时段、班容、供给或阈值。
- 不删除校验问题，不把警告改写成确定结论。
- 不自动调用飞书、钉钉、招聘、采购或租赁工具。
- 消息载荷必须保留 `requires_human_confirmation: true`。
- 每个结论必须带数据日期、配置版本、输入哈希或证据键。

## 输出

交付以下三个文件：

- `campus-capacity-result.xlsx`
- `management-recommendations.md`
- `message-payload.json`

优先解释首次缺口月份、峰值缺口、持续期、提前期和情景差异。若用户要求修改口径，先更新配置或参考规则，再重跑全部评测。
