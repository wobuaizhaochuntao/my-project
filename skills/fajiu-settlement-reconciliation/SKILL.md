---
name: fajiu-settlement-reconciliation
description: >-
  Reconciles Fajiu monthly settlement application Excel files against posting-success Excel files,
  including exact repayment-amount aggregation by loan order, one-to-one date matching,
  bidirectional marking, and balance validation.
  Use when the user asks to核对结算、入账成功、流水导出、法久_YYYYMM.xlsx、月度结算表，
  或标记未在结算表体现的入账记录。
type: prompt
whenToUse: 当用户说"核对结算"/"入账核对"/"结算体现标记"，或提供结算申请表、法久入账明细 Excel 时
---

# 法久月度结算入账核对

将“结算申请表”与“法久入账明细”双向核对，并生成两份带标记的 Excel。

## 执行

1. 确认两个原始文件：
   - 结算表（法久月度结算确认）：常见名称为 `法久_202609.xlsx`（订单号/还款总金额/repaid_date）
   - 入账表（律助流水复核导出）：常见名称为 `流水导出-10月结算.xlsx`、`9月结算.xlsx`（借款单号/还款金额/复核日期）
2. 不要把此前生成的 `_已核对` 或 `_体现标记` 文件误当成原始输入，除非用户明确指定。
3. 执行本 Skill 的脚本：

```bash
./.venv/bin/python scripts/reconcile.py --settlement "<法久结算表.xlsx>" --posting "<流水导出表.xlsx>"
```

如需指定输出：

```bash
./.venv/bin/python scripts/reconcile.py \
  --settlement "<法久结算表.xlsx>" \
  --posting "<流水导出表.xlsx>" \
  --settlement-output "<结算核对结果.xlsx>" \
  --posting-output "<入账体现标记.xlsx>"
```

列名与默认口径不一致时用 `--settlement-cols` / `--posting-cols` 覆盖（如 `order=借款单号,amount=还款金额,review_date=复核日期`）；
时间口径用 `--date-order` 指定（见下）。

脚本依赖 `openpyxl`，已安装在本 skill 目录的 `.venv` 虚拟环境中，用 `./.venv/bin/python` 运行即可。
若虚拟环境损坏，重建：`python3 -m venv .venv && ./.venv/bin/pip install openpyxl`。

## 固定核对口径

- 身份字段：结算表“订单号”对应入账表“借款单号”。
- 金额字段：结算表“还款总金额”；入账表“还款金额”（勿用“流水金额”）。
- 金额精确到分，不允许误差。
- 同一订单号下，多条入账“还款金额”可以加总后对应一条结算记录。
- 每条结算记录和每条入账记录最多使用一次。
- 正常时间口径（先复核、后结算确认）：入账表“复核日期”不晚于结算表 repaid_date；
  复核日期晚于 repaid_date 的组合标为时间异常。
- 当存在多种组合时，依次最大化：
  1. 匹配的入账记录数
  2. 覆盖的结算记录数
  3. 时间最接近程度

## 输出标记

结算表新增：

- 核对结果
- 入账成功日期
- 匹配还款金额
- 时间差（天）
- 合并匹配组
- 合并后金额
- 复核口径

结算记录分类：

- `入账成功`
- `时间异常待复核`
- `重复申请/无剩余成功记录`
- `金额不一致`
- `未找到入账成功记录`

入账表“明细”新增：

- 结算体现情况
- 对应结算行
- 对应还款金额合计
- 时间核对

入账记录分类：

- `已在结算表体现`
- `已体现-时间异常待复核`
- `未在结算表体现（无此订单）`——订单号在结算表中未出现（红色）
- `未在结算表体现（金额不一致）`——订单在结算表有，但金额加总对不上或已被其他记录占用（橙色）

## 验证与汇报

脚本结束后必须检查 JSON 中：

- `balanced` 必须为 `true`
- 结算表正常成功金额必须等于入账表正常匹配金额
- 各分类数量之和必须分别等于两份输入明细的数据行数

输出文件名自带时间戳后缀（如 `法久_202609_已核对_20261009_231606.xlsx`），重复运行或源文件正被 Excel 打开也不会覆盖冲突。如需固定文件名再用 `--settlement-output` / `--posting-output` 指定。

向用户简要汇报：

- 入账成功的结算记录数
- 对应入账记录数
- 正常匹配金额
- 时间异常数量
- 未在结算表体现的入账数量及金额（区分「无此订单」和「金额不一致」两种）
- 两份输出文件的可点击路径

保留原始文件，默认只生成新文件。
