---
name: debt-transfer-sms
description: >-
  Generate Fajiu (法久) debt-transfer SMS upload Excels (FJ-债转短信-*.xlsx)
  from债转短信/全部订单/短信源表, including一转二转 split, party short names,
  众邦银行 funder field, and Excel SaveAs format fix. Use when the user asks
  to generate债转短信/短信表格/FJ-债转短信, mentions短信发送, or provides
  债转短信*.xlsx / 全部订单 / 逾期客户信息汇总 for SMS batching.
type: prompt
whenToUse: 当用户说"生成债转短信"/"FJ-债转短信表"/"短信表格"，或提供债转短信、全部订单、短信源表 Excel 时
---

# 法久债转短信生成

把源订单表拆成可上传法久的 `FJ-债转短信-{出让}转{受让}-{MMDD}.xlsx`。

## 用户需提供

1. **源 Excel**（二选一）
   - `债转短信*.xlsx`：含 `一转` / `二转` 表
   - 扁平案件表：`全部订单*.xlsx` / `短信 *.xlsx`（含担保方、受让方、反担保、可选 `_主表错误`）
2. **文件名日期 MMDD**（发送日/批次日，如 `0911`）
3. 可选：**模板**（一份已能上传法久的 `FJ-债转短信-*.xlsx`）；不传则自动找最近模板
4. 可选：核对表 `逾期客户信息汇总*.xlsx`

## 脚本（优先执行）

目录：`${KIMI_SKILL_DIR}/scripts/`（依赖：`openpyxl`；格式另存还需本机 Excel + `pywin32`）

| 脚本 | 用途 |
|------|------|
| `generate_sms.py` | 读取源表并生成全部 FJ 短信 xlsx |
| `common.py` | 主体简称、列号、Excel SaveAs |

```bash
python "${KIMI_SKILL_DIR}/scripts/generate_sms.py" -i "短信 0910.xlsx" -o "./短信0910" -d 0911
python "${KIMI_SKILL_DIR}/scripts/generate_sms.py" -i 债转短信0904.xlsx -o ./out -d 0904 --dry-run
python "${KIMI_SKILL_DIR}/scripts/generate_sms.py" -i 全部订单.xlsx -o ./out -d 0908 -t "已能上传的模板.xlsx"
```

Agent 应直接跑脚本，不要每次手写生成逻辑。

## 总流程

```
Task Progress:
- [ ] 1. 确认源表路径与 --date(MMDD)
- [ ] 2. dry-run 核对分组与笔数
- [ ] 3. 正式生成（默认 Excel SaveAs）
- [ ] 4. 如有逾期汇总表，按订单号核对担保/反担保/受让
- [ ] 5. 交付文件清单
```

## 拆分规则

### A. `债转短信*`（一转/二转）

| 表 | FROM | TO |
|----|------|----|
| 一转 | 担保方名称 | 反担保主体 |
| 二转 | 反担保主体 | 受让方名称 |

反担保名去掉尾部数字 ID（如 `北京锋泰科技有限公司3660` → 锋泰全称）。

### B. 扁平案件表

1. **每笔都生成**：反担保 → 受让方
2. **额外一转**（需两次通知）：
   - 优先：`_主表错误` 含 `0/2`
   - 否则：无该列且 `反担保 ≠ 担保方` 时，补 `担保方 → 反担保`

## 填写与命名

命名：`FJ-债转短信-{FROM简称}转{TO简称}-{MMDD}.xlsx`

只填这些列（其余留空）：

| 列 | 字段 | 值 |
|----|------|----|
| A | order_id | 订单号 |
| G | 当前债权归属主体 | FROM 全称 |
| M | 二次债转债权受让方 | FROM 全称 |
| X | 第三次债转债权出让方 | TO 全称 |
| AE | 资金方名称 | 默认 **众邦银行** |

- **不要**加 `短信发送时间` 列（法久校验用 31 列到 AE）
- 表头 sheet 名保持模板的 `一转`
- 生成后必须 **Excel SaveAs**（openpyxl 直出常被法久格式校验拒绝）

## 主体简称

见 `${KIMI_SKILL_DIR}/scripts/common.py` 的 `SHORT_PARTIES`。常见：

SNX/YR/FT/HR，序嘉洪/序洪/序普洪/启心程/洲序/灵机一动/臻行远/竣洪/洪均/万妙汇新

遇到未知主体：先问用户简称，再写入 `SHORT_PARTIES` 后重跑。

## 格式校验要点

上传失败时对照：

- 需有 `sharedStrings.xml`
- workbook 关系用相对路径 `worksheets/sheet1.xml`
- XML 带 `<?xml ...?>`
- 不要多出第 32 列
- 空单元格不要标 `t="n"` 却无值

脚本默认用本机 Excel COM `SaveAs(FileFormat=51)` 修复。

## 核对（可选）

有 `逾期客户信息汇总*.xlsx` 时：

1. 订单号集合一致（源表 ∩ 生成文件 ∩ 逾期表）
2. 逾期表担保/反担保/受让 = 源表
3. 生成文件 FROM/TO 符合拆分规则
4. `0/1` 只出一份；`0/2` 出两份（担保→反担保 + 反担保→受让）

## 更多

- 示例命令见 [examples.md](examples.md)
