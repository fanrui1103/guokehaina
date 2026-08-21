# Case: Label Photo OCR → Structured Excel

**工具目录：** [`tools/label-ocr`](../tools/label-ocr)

## Problem

仓库/跟单拿到物料标签照片或截图后，要填进固定格式的物料清单 Excel。字段多（批次、产地、料号、条码、有效期等），人工抄写慢，且不同标签印刷用词不一致。

## Approach

1. 用户先填本批共用信息（订单号、送货单号、供方等）  
2. 选择多张标签图或整个文件夹  
3. OCR 识别后，按可配置字段别名映射到清单列  
4. 界面里可双击改错格，再导出，不覆盖原始模板文件  
5. 换标签版式时，优先改 `field_map.json` 别名，而不是改死代码  

## Tech

- Python 桌面应用  
- OCR 流水线（识别 + 字段抽取）  
- JSON 可配置字段别名  
- Excel 模板写出  

## Outcome

- 「看图抄表」变成批量识别 + 人工复核  
- 配置与代码分离，方便适应新标签文案  
- 体现：**非结构化输入 → 结构化表格**，以及给业务同事留复核入口  

## What to look at in code

- `src/extractor.py` — 识别与抽取  
- `src/field_map.json` — 字段别名  
- `src/excel_writer.py` / `src/app.py` — 导出与界面  
