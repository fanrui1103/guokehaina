# Manufacturing Ops Automation Toolkit

**制造业跟单效率工具集** — Python 桌面工具，把重复的 Excel / PDF / 标签工作自动化，并交付给零基础业务同事使用。

> Internship / portfolio repository. Built during manufacturing ops internship work: clarify messy real-world requirements → ship usable tools → package for non-technical users.

[简体中文说明](#中文) · [English](#english)

---

<a id="english"></a>

## English

### What this is

A collection of **small, production-facing desktop utilities** for manufacturing order-fulfillment workflows:

| Area | What it automates |
| --- | --- |
| Shipping / delivery reports | Fill PDF report templates from PO / delivery Excel |
| Labels | Batch-change serials & QR codes; OCR labels into Excel |
| Planning / inventory | Enrich reservation-stock reports; shortage comparison |

### Skills demonstrated

- Turning vague shop-floor needs into concrete rules (edge cases, remainders, alerts)
- Python automation for **Excel / PDF / images**
- Desktop UX for non-engineers (forms, previews, one-click exe packaging)
- Cross-file matching, validation, and exception highlighting
- End-to-end delivery: source → usable tool for colleagues

### Flagship case studies

1. [Shipping report PDF generation](docs/case-shipping-report.md)
2. [Reservation & stock report enrichment](docs/case-reservation.md)
3. [Label serial + QR batch generation](docs/case-label-qr.md)
4. [Label photo OCR → structured Excel](docs/case-label-ocr.md)

### Repository layout

```text
tools/
  shipping-report/        # OQC-style shipping PDF from PO + templates
  delivery-report/        # Shipping PDF from delivery-note Excel
  packing-label-excel/    # Inner/outer/pallet label Excel generator
  label-qr/               # Template-based serial + QR batch labels
  label-ocr/              # Label image OCR → material list Excel
  reservation-enrich/     # Reservation/stock report + MPS enrichment
  shortage-compare/       # Shortage vs inventory vs production plan
docs/                     # Case write-ups for recruiters / interviewers
```

Source only on `main`. Packaged `.exe` builds are **not** kept in git (they bury the code signal). Local packaging scripts remain where useful.

### Quick start

```bash
git clone https://github.com/fanrui1103/guokehaina.git
cd guokehaina/tools/<tool-folder>
pip install -r requirements.txt   # if present
# then run the .bat or: python app.py / python <entry>.py
```

Each tool folder has a short `README.md`.

### Note on anonymization

Client / vendor brand names in UI copy or sample templates are treated as **anonymized client cases** for portfolio presentation. Do not treat sample files as production data.

---

<a id="中文"></a>

## 中文

### 这是什么

实习期间围绕**制造业跟单**场景做的一组 Python 小工具：把同事每天重复的 Excel / PDF / 标签操作自动化，并做成非技术同事也能用的桌面程序。

### 能体现的能力

- 把含糊的业务口述，问清楚成可执行规则（尾数、预警、模板约束）
- Excel / PDF / 图像 OCR 等办公自动化
- 面向业务同事的界面与免安装交付思路
- 跨表匹配、校验、异常高亮
- 从需求 → 实现 → 交付的闭环

### 主打案例（建议面试官优先看）

1. [出货报告 PDF 自动生成](docs/case-shipping-report.md)
2. [预约与库存报表自动补数](docs/case-reservation.md)
3. [物料标签：只改序列号与二维码](docs/case-label-qr.md)
4. [标签照片 OCR 填入物料清单](docs/case-label-ocr.md)

### 目录

见上方 `tools/` 结构。仓库主分支**只保留源码与说明**，不把绿色版 `_internal` 大包堆在 Git 里，方便审阅代码。

### 本地运行

进入 `tools/某工具/`，按该目录 `README.md` 安装依赖并启动。

### 说明

对外展示时按「客户项目 / 脱敏案例」理解即可；样例模板仅用于演示格式，不代表可公开的业务数据。
