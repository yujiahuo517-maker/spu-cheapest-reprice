# spu-cheapest-reprice

京喜「导出 SPU 下最便宜商品链接并改价」的 JoyCode Skill：把京喜商家后台导出的「导出商品JX-SKU信息」xlsx 填进底表后，自动算出每个 SPU（商品编码）下京东价最低的商品链接和改价建议，复现原 Excel Power Query 模板的全部逻辑，且**不依赖 Excel / Power Query**。

## 逻辑（复现 PQ 模板查询「0」）

1. 底表 DATA 读入导出数据
2. 建议改动价格 = `可控采购价 > 1000 ? 可控采购价*1.01 + 2 : 可控采购价 + 2`
3. 按 `商品编码` 分组，取组内 `京东价` 最小的一行
4. 输出改嫁表 10 列，按可控采购价降序

## 用法

### 最快：拖放（无需打开 JoyCode）

把「导出商品JX-SKU信息....xlsx」直接拖到根目录的 `一键改价.bat` 上，自动生成结果并弹出资源管理器定位到输出文件。前提是装了 Python 且 `pip install openpyxl`。

### 命令行 / JoyCode

```powershell
python scripts/fill_reprice.py "<导出商品JX-SKU信息....xlsx>" [-o 输出路径.xlsx]
```

- 默认输出：导出文件同目录 `<原名>_改价结果.xlsx`（含已填的「底表」+ 算好的「改嫁表」）
- 原 PQ 模板保留在 `assets/PQ模板.xlsx`，生成结果以它为底本，不会修改原文件

依赖：Python 3 + `openpyxl`。

## 性能

- 十万行内基本秒级；实测 20 万 SKU / 6 万 SPU 端到端约 2~3 分钟
- 上限受 Excel sheet 行数限制（1,048,576 行）

## 目录

```
spu-cheapest-reprice/
├── SKILL.md            # JoyCode skill 定义（触发词 + 流程）
├── agents/openai.yaml  # UI 元数据
├── assets/PQ模板.xlsx  # 原 PQ 模板（作为生成底本）
└── scripts/fill_reprice.py
```
