---
name: xdf-normalize-teacher-schedule
description: 校验重庆新东方小学素养班级明细，把暑假与秋季一维导出转换为教师时段占档表、容量差额表和待确认清单。支持主带课老师、回传模板优先级及可回退的本地规则覆盖；不用于自动排课或越过硬约束的人工调整。
---

# 教师排课占档规整

把教务系统的班级明细转换为校区主管可读、可筛选、可复核的课表。所有人数、容量、冲突选择和汇总均由脚本确定性计算。

## 适用范围

- 重庆新东方校区主管查看小学一至六年级素养课程。
- 科目只包括：思维（NL/SW）、双语 STEAM（ST）、博学（BX）。
- 首个可运行版本读取暑假和秋季明细；寒假、春季口径已记录，但没有对应源表时不得虚构输出。
- 输出用于查看教师占档、可排空段和班型供给多样性，不替代人工排课决策。

## 工作流

1. 确认源文件是本地脱敏或受控的 `.xlsx`，不得提交到 Git。
2. 阅读 [data-contract.md](references/data-contract.md) 和 [business-rules.md](references/business-rules.md)。
3. 按实际源表列名核对版本化初始配置 `config/teacher-schedule.json`，不得让模型猜列，也不得把自然语言覆盖直接写回该文件。
4. 加载 Skill 后先运行一次基准课表：

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/normalize_schedule.py \
  --input <班级明细.xlsx> \
  --config config/teacher-schedule.json \
  --campus <完整校区名称> \
  --data-as-of YYYY-MM-DD \
  --output-dir <输出目录>
```

同一管理口径包含多个系统校区时，优先使用配置中的管理单元：

```bash
  --campus-group 礼嘉管理单元
```

临时组合也可以重复提供 `--campus <完整校区名称>`。输出摘要必须同时展示管理口径和实际纳入的源校区名称。

若使用者在上次生成的课表中按约定写入计划新班，增加 `--manual-template`；暑假和秋季模板分开保存时可重复提供：

```bash
  --manual-template <人工修改后的课表.xlsx>
```

人工模板必须执行三向对比：数据源新增、模板遗留、同班号教师/时间变更。只比较当前支持的 ST/BX/NL/SW，小学范围外编码不得制造陈旧班误报。

回传模板后必须显式选择开关并重新运行：

```bash
  --template-policy rule_first
  --template-policy template_first
```

- `rule_first`（默认）：教务数据与确定性规则优先，模板中的删班、换教师、换时段只提示差异。
- `template_first`：对已回传季节，模板中正式班的保留、教师和时段优先；未出现的源班不入表。安全、支持范围、课次、班号、容量和季节是硬约束，模板不能覆盖。

5. 检查退出码和 `schedule-review.json`：
   - `ready`：无异常，仍需使用者确认是否存在系统导出后的人工变化。
   - `needs_confirmation`：已生成安全子集，但冲突、非标准课次或人工候选必须确认。
   - `blocked`：输入安全、数据合同或唯一性失败；只解释校验问题，不输出正式课表和容量结论。
6. 阅读 [output-spec.md](references/output-spec.md) 做视觉和业务复核。
7. 生成基准课表后必须询问使用者：是否有关班、新开班、换教师或移动时段等人工调整。使用者回传修改结果后，按选定开关复用同一数据源重跑，不得仅凭任意底色推断业务规则。

## 自然语言规则变更与回退

AI 先把使用者的自然语言要求转换为最小 JSON 补丁，展示“旧值→新值”，然后只能通过受限管理脚本写入 Git 已忽略的本地覆盖：

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/manage_schedule_rules.py \
  --base config/teacher-schedule.json \
  --overlay data/local/teacher-schedule-rules/active-overrides.json \
  apply --patch <AI生成并校验的补丁.json> --request "<使用者原话>"
```

重跑课表时增加：

```bash
  --rules-overlay data/local/teacher-schedule-rules/active-overrides.json
```

只允许覆盖校区组、标准课次、容量、时段/轮次、排除值和配色。源字段映射、安全、班号分类、冲突策略和计划新班语法不得由自然语言覆盖。

使用者说“恢复初始规则”、“回到初始规则”时，调用：

```bash
python3 .agents/skills/xdf-normalize-teacher-schedule/scripts/manage_schedule_rules.py \
  --base config/teacher-schedule.json \
  --overlay data/local/teacher-schedule-rules/active-overrides.json \
  restore --request "<使用者原话>"
```

初始规则始终保留在版本化配置中；恢复操作只清空本地覆盖，随后必须重跑并报告差异。基础配置哈希变化时阻断旧覆盖，不得静默迁移。

## 计划新班格式

只识别带 `新班` 前缀的单元格：

- `新班ST3A`
- `新班SW3B`
- `新班NL3A`
- `新班BX3A`
- `新班NL1`、`新班NL2`：低年级思维基础班型。
- `新班NL1A`、`新班NL2A`：低年级思维拔高班型。

计划新班当前数量为 0，容量按科目和年级规则取得。其他 `NL1/NL2` 后缀进入待确认。

## 强制边界

- 输入含公式、外部链接、疑似公式注入文本时阻断。
- 复习班、补课班、定金班、取消班、S2/S3 和非小学课程不进入正式结果。
- 非标准课次进入待确认，确认前不进入课表和容量汇总。
- 暑假同时包含周课和连续课时，优先用带“每天/每日/连续”的密集片段识别轮次与时段；多个密集片段结论不一致时进入待确认。
- 班型以班号为主、班级名称交叉验证；冲突时采用班号并提示。
- `授课教师`只有一人时直接入表；有多人时必须从 `主带课老师`取唯一且存在于授课教师列表的人选，否则进入待确认。
- 同一教师同一时段多班时，人数唯一较大的班临时入表，其余进入冲突清单；最高人数相同时必须人工确认。
- 人工模板中的正式班教师/时段变化只能按显式 `template_first` 开关应用；默认 `rule_first` 只做差异提示。
- 不自动发送消息，不修改教务系统，不自动确认人工排课。
- 输出、日志和公开仓库不得包含真实样本、工号或可逆身份映射。

## 输出

- `teacher-schedule-result.xlsx`
- `schedule-review.json`

工作簿包含运行摘要、校验问题、排除记录、各季课表、容量总览、容量汇总、容量明细、排课冲突、待确认、模板对比和假设与版本。容量指标仅展示总容量、当前数量和差额。
