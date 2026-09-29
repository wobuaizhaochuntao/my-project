# Kimi Code Skill 清单

> 最后更新：2026-09-29。用户级 skill 位于 `C:\Users\Administrator\.kimi-code\skills\`，项目级位于 `d:\vscode\skills\`（同名时项目级优先）。

## 业务 Skill（共 5 个）

### 1. mark-pushable-orders — 未推送订单标记
- **什么时候用**：给出洋钱罐「未推送订单」导出表，要求判断/标记/筛选"可以推送"、"哪些单子能推"。
- **规则**：`_主表错误` 列非空的行不推送；同一借款人（按身份证号码，缺失按被申请人姓名）任一案件有错，名下全部连坐不推送。
- **输入**：未推送订单 xlsx（需含 `_主表错误`、`身份证号码`、`被申请人姓名` 列）
- **输出**：追加「是否可推送」（红绿色底）「不推送原因」两列 + 「标记说明」汇总表的新 xlsx，默认命名 `<输入名>_已标记.xlsx`。
- **脚本**：`scripts/mark_pushable.py`（列名可用 `--error-col`/`--id-col`/`--name-col` 指定）
- **注意**：同名不同人可能被一并排除；只依据主表错误列，`_明细错误` 不参与。

### 2. spv-debt-transfer — SPV 集中度化解方案
- **什么时候用**：要求按「批次×法院×SPV 单元格 ≤5 件/月」卡控做债转方案、化解 SPV 集中度、对比 4/8/12/16/22 次档位达标率、估算债转后多立案数量。
- **输入**：案件明细 CSV（批次/法院/当前状态/spv/揽收时间/诉请代偿款/案号时间/还款计划创建时间）+ 订单明细 XLSX（订单号/反担保主体/担保方名称）。
- **输出**：债转明细/路径汇总/法院×SPV 分布表 xlsx（三色标记），脚本打印迁移件数、整户率、MSCORE 分布、达标率等。
- **脚本**：`scripts/spv_allocator.py` — `--maxp 2`（=8次档）`--sweep`（全档位对比）`--cutoff-batch`（截止批次）`--exclude-spv`（禁用SPV）。
- **关键口径**（勿擅自改）：占容=未完结；可迁移池五状态（待提交/回收待处理/拒不受理/待审查/已审查）；邮寄优先级 未揽收>超期揽收>近期揽收；同人整户优先；不可改善组达标率剔除。
- **详述**：`references/caliber.md`（含节奏估算三口径公式、Excel 交付规范、8 条已踩过的坑）。

### 3. debt-transfer-contracts — 债转合同生成
- **什么时候用**：说"生成债转协议/做债转合同/第N批一转二转协议"，或提供债转文件 Excel。
- **做什么**：从债转文件 Excel 生成法久（Fajiu）债转 Word 合同，含一转/二转/三转/补充协议、SPV 当事方更新、附件表格、校验、供应商客户录入明细。

### 4. debt-transfer-sms — 债转短信表生成
- **什么时候用**：说"生成债转短信/FJ-债转短信表/短信表格"，或提供债转短信、全部订单、短信源表。
- **做什么**：生成短信上传 Excel（FJ-债转短信-*.xlsx），含一转二转拆分、当事方简称、众邦银行资金方字段、Excel SaveAs 格式修复。

### 5. call-quality-inspection — 通话录音质检分析
- **什么时候用**：给出「通话质检记录」.xls（HTML 表格）/ 录音文件夹，要求质检、风险排查、清理录音、生成质检报告。
- **流程**：`parse_export.py`（解析导出→records.csv+transcripts）→ `scan.py`（五类关键词扫描：辱骂/投诉监管/情绪激动/还款能力/身份核实）→ **逐条精读全文定级（不可跳过，关键词只是线索）** → `build_xlsx.py`（生成质检分析版 xlsx，补全复核列，风险等级红黄绿着色）→ `clean_recordings.py`（按审定后的表格清理录音，走回收站可恢复）→ 可选 Markdown 报告。
- **注意**：文字版为空的标「无法判断」，不得编造；夜间外呼（21点后）单独检查；还款能力类关键词只作备注。
- **定级标准**：`references/rubric.md`。

## 流水线关系

```
未推送订单标记(mark-pushable-orders)
        │
SPV 集中度化解方案(spv-debt-transfer)  ←── 出债转明细/分配方案
        │
        ├─→ 债转合同(debt-transfer-contracts)
        └─→ 债转短信(debt-transfer-sms)
```

通话质检（call-quality-inspection）为独立业务线。

## 项目级副本（d:\vscode\skills\）

| 目录 | 说明 |
|---|---|
| debt-transfer-contracts | 与用户级同名，项目级优先生效 |
| debt-transfer-sms | 与用户级同名，项目级优先生效 |

⚠️ 改用户级这两份时注意与项目级副本保持一致，否则项目里跑的是旧版。

## 内置通用 Skill

- `check-kimi-code-docs` — 查 Kimi Code 官方文档
- `update-config` — 查看/修改 config.toml、tui.toml 配置
- `write-goal` — 辅助编写 /goal 目标
