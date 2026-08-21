# Case: Shipping Report PDF Generation

**工具目录：** [`tools/shipping-report`](../tools/shipping-report)  
**相关：** [`tools/delivery-report`](../tools/delivery-report)（送货单 Excel 驱动的同类流程）

## Problem

跟单同事需要按采购订单生成多份出货检验报告 PDF。每种物料有自己的报告模板；订单数量、样本信息等字段要改，其余版式必须与客户模板一致。手工改 PDF 慢且容易漏改。

## Approach

1. 解析采购订单 PDF，识别料号与数量需求  
2. 匹配各料号对应的出货报告模板  
3. 按规则改写关键字段（日期、订单数量、样本相关字段等），**其余版面保持模板原样**  
4. 浏览器/桌面表单让同事勾选料号、填写发货数量后批量导出  

`delivery-report` 把输入换成「送货单 Excel + 物料 PDF 模板」，解决另一条真实出货路径。

## Tech

- Python  
- PDF 读写与版面保持（PyMuPDF 等）  
- 轻量 Web/桌面操作页（Flask 等）  
- 可打包为同事免安装使用的桌面程序  

## Outcome

- 同事可自助生成，减少手工改 PDF  
- 规则与模板约束被写进程序，降低「改错不该改的地方」的风险  
- 体现：**真实业务约束下的自动化**，而不是演示用假数据 Demo  

## What to look at in code

- `po_parser.py` / `excel_parser.py` — 输入解析  
- `report_gen.py` — 按模板生成  
- `app.py` + `templates/` — 给业务同事的操作界面  
