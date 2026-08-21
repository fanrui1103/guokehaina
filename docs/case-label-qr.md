# Case: Label Serial + QR Batch Generation

**工具目录：** [`tools/label-qr`](../tools/label-qr)

## Problem

物料标示单以 Excel 模板形式存在。批量出货时要生成很多份，每份**只有序列号和二维码不同**，厂商 / 料号 / 品名等必须完全跟模板走。手工改容易漏改二维码，或改坏不该动的单元格。

## Approach

1. 用户选择模板工作簿与工作表  
2. 从样例行读取固定字段，并解析样例二维码结构  
3. 用户输入序列号范围（支持 `1-24,20-30` 这类写法）  
4. 批量生成：旧料号留空；仅替换表格序列号与二维码末段序列  
5. 提供预览与「检查已生成文件的二维码」，确认每张码末尾序列确实在变  

## Tech

- Python + openpyxl  
- 二维码编解码校验  
- Tkinter 桌面界面  
- PyInstaller 打包脚本（`build_exe.py`）  

## Outcome

- 把「复制模板改号」变成可选范围批量生成  
- 用检查功能降低「表上号变了、码没变」的事故  
- 体现：**模板驱动生成 + 可验证的输出**  

## What to look at in code

- `generator.py` — 读模板、生成、验码  
- `app.py` — 交互与预览  
