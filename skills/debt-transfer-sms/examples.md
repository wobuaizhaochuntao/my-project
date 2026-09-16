# 债转短信生成示例

## 1. 扁平案件表（短信 / 全部订单）

```bash
cd "C:/Users/Administrator/Desktop/周报/短信"
python "%USERPROFILE%/.kimi-code/skills/debt-transfer-sms/scripts/generate_sms.py" ^
  -i "短信0910/短信 0910.xlsx" ^
  -o "短信0910" ^
  -d 0911
```

预期：按反担保→受让方拆分；`_主表错误` 含 `0/2` 的再补担保→反担保。

## 2. 债转短信一转/二转表

```bash
python "%USERPROFILE%/.kimi-code/skills/debt-transfer-sms/scripts/generate_sms.py" ^
  -i "短信0904/债转短信0904.xlsx" ^
  -o "短信0904" ^
  -d 0904
```

## 3. 先预览分组

```bash
python "%USERPROFILE%/.kimi-code/skills/debt-transfer-sms/scripts/generate_sms.py" ^
  -i "短信0910/短信 0910.xlsx" ^
  -o "短信0910" ^
  -d 0911 ^
  --dry-run
```

## 4. 指定模板 / 跳过 Excel 另存

```bash
python .../generate_sms.py -i src.xlsx -o out -d 0911 ^
  -t "短信0904/FJ-债转短信-SNX转序嘉洪-0904.xlsx"

python .../generate_sms.py -i src.xlsx -o out -d 0911 --skip-excel-saveas
```

## 输出样例

```
FJ-债转短信-SNX转竣洪-0911.xlsx
FJ-债转短信-SNX转FT-0911.xlsx
FJ-债转短信-FT转万妙汇新-0911.xlsx
```
