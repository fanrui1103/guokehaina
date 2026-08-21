"""智能排产（改良版）

用法和原来差不多：选「客户文件」+「公司文件」→ 开始处理 → 导出结果。
在原合并结果基础上，于「订单数量」右侧增加「投产减入库」列。
"""

from __future__ import annotations

import re
import traceback
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pandas as pd

NEW_COL = "投产减入库"
APP_TITLE = "智能排产（改良版）"


def _to_number(value: Any) -> float:
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _norm_key(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _excel_date(value: Any) -> str | None:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)) and not pd.isna(value) and float(value) > 40000:
        return (datetime(1899, 12, 30) + timedelta(days=int(value))).strftime("%Y-%m-%d")
    return None


def pre_customer(df: pd.DataFrame) -> pd.DataFrame:
    """解析预约与库存报表：多层表头，只保留「需求」行。"""
    if df.shape[0] < 5:
        raise ValueError("客户文件行数太少，请确认选的是「预约与库存报表」")

    cols: list[str] = []
    for c in range(df.shape[1]):
        top = df.iloc[1, c]
        bot = df.iloc[2, c]
        top_s = "" if pd.isna(top) else str(top).strip()
        dated = _excel_date(bot)
        if dated:
            cols.append(dated)
            continue
        bot_s = "" if pd.isna(bot) else str(bot).strip()
        if bot_s == "类型":
            cols.append("类型")
        elif bot_s == "订单量" and "在途" in top_s:
            cols.append("在途订单量")
        elif bot_s == "交货量":
            cols.append("已预约交货量")
        elif bot_s == "库存" and "当前" in top_s:
            cols.append("库存")
        elif bot_s == "采购量":
            cols.append("采购量")
        elif top_s:
            cols.append(top_s)
        elif bot_s:
            cols.append(bot_s)
        else:
            cols.append(f"col{c}")

    body = df.iloc[4:].copy()
    body.columns = cols
    body = body.loc[:, ~pd.Index(body.columns).duplicated()].copy()
    body = body.where(body.notna(), "")
    if "类型" not in body.columns:
        raise ValueError("客户文件里找不到「类型」列，请确认文件格式")
    body = body[body["类型"].astype(str).str.strip() == "需求"].copy()
    body = body.reset_index(drop=True)

    date_cols = [c for c in body.columns if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(c))]
    if date_cols and "昨日" not in body.columns:
        body.insert(list(body.columns).index(date_cols[0]), "昨日", 0)
    return body


def fuzzy_match(row: pd.Series, company_sorted: pd.DataFrame) -> pd.Series:
    empty = pd.Series(
        {"订单数量": 0.0, "仓库结存": 0.0, "产品编号": "未知", "查询状态": "false"}
    )
    sku = _norm_key(row.get("供应料号", ""))
    if not sku:
        return empty

    matches = company_sorted[company_sorted["客户货号"].map(_norm_key) == sku]
    if matches.empty:
        matches = company_sorted[
            company_sorted["客户货号"].astype(str).str.contains(sku, regex=False, na=False)
        ]
    if matches.empty:
        return empty

    unfinished = matches[matches["完工状态"].astype(str).str.strip() != "已完工"]
    total = pd.to_numeric(unfinished["订单数量"], errors="coerce").fillna(0).sum()
    warehouse = _to_number(matches.iloc[0]["仓库结存"])
    product = _norm_key(matches.iloc[0]["产品编号"]) or "未知"
    return pd.Series(
        {
            "订单数量": float(total),
            "仓库结存": warehouse,
            "产品编号": product,
            "查询状态": "true",
        }
    )


def calculate_deadline(row: pd.Series, date_cols: list[str]) -> str:
    """按客户库存滚动扣减日期需求，推算排产截至日。"""
    if not date_cols:
        return ""
    today = datetime.strptime(date_cols[0], "%Y-%m-%d")
    tomorrow = today + timedelta(days=1)
    supply = _to_number(row.get("库存", 0))
    cumulative = 0.0
    demand_cols = ["昨日"] + date_cols if "昨日" in row.index else date_cols
    for date_str in demand_cols:
        cumulative += _to_number(row.get(date_str, 0))
        if cumulative > supply:
            if date_str == "昨日":
                return tomorrow.strftime("%Y-%m-%d")
            demand_date = datetime.strptime(date_str, "%Y-%m-%d")
            deadline = demand_date - timedelta(days=1)
            if deadline < tomorrow:
                deadline = tomorrow
            return deadline.strftime("%Y-%m-%d")
    return ""


def build_diff_map_by_sku(company: pd.DataFrame) -> dict[str, float]:
    """按客户货号汇总未完工制令的 Σ(投产数量 − 生产入库数)。已完工不计入。"""
    by_sku: dict[str, float] = {}
    for _, row in company.iterrows():
        if _norm_key(row.get("完工状态")) != "未完工":
            continue
        sku = _norm_key(row.get("客户货号"))
        if not sku:
            continue
        diff = _to_number(row.get("投产数量")) - _to_number(row.get("生产入库数"))
        by_sku[sku] = by_sku.get(sku, 0.0) + diff
    return by_sku


def lookup_diff(row: pd.Series, by_sku: dict[str, float]) -> Any:
    """合并表「供应料号」对应主生产计划「客户货号」。"""
    sku = _norm_key(row.get("供应料号", ""))
    if sku and sku in by_sku:
        return by_sku[sku]
    return pd.NA


def process_data(customer: pd.DataFrame, company: pd.DataFrame) -> pd.DataFrame:
    needed = ["制令单号", "客户货号", "完工状态", "订单数量", "仓库结存", "产品编号", "投产数量", "生产入库数"]
    missing = [c for c in needed if c not in company.columns]
    if missing:
        raise ValueError("公司文件（主生产计划）缺少列：" + "、".join(missing))

    company_sorted = company.sort_values("制令单号")
    matched = customer.apply(lambda r: fuzzy_match(r, company_sorted), axis=1)
    result = customer.copy()
    for col in matched.columns:
        result[col] = matched[col]

    date_cols = sorted(c for c in result.columns if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(c)))
    result["是否需要追加生产"] = (
        pd.to_numeric(result["采购量"], errors="coerce").fillna(0)
        > pd.to_numeric(result["订单数量"], errors="coerce").fillna(0)
        + pd.to_numeric(result["仓库结存"], errors="coerce").fillna(0)
    ).map({True: "是", False: "否"})
    result["排产截至"] = result.apply(lambda r: calculate_deadline(r, date_cols), axis=1)

    # 列名与原程序导出对齐
    result = result.rename(
        columns={
            "库存": "客户库存",
            "采购量": "当前可采购量",
        }
    )

    by_sku = build_diff_map_by_sku(company)
    result[NEW_COL] = result.apply(lambda r: lookup_diff(r, by_sku), axis=1)

    front = [
        "交货地点",
        "供应料号",
        "品名规格",
        "在途订单量",
        "已预约交货量",
        "客户库存",
        "当前可采购量",
        "订单数量",
        NEW_COL,
        "仓库结存",
        "产品编号",
        "是否需要追加生产",
        "排产截至",
    ]
    front = [c for c in front if c in result.columns]
    middle = date_cols
    tail = ["查询状态"] if "查询状态" in result.columns else []
    other = [c for c in result.columns if c not in front + middle + tail + ["类型", "昨日", "序号", "趋势", "单位", "供应商料号", "5周订货量", "5周计用量合计", "14天计划用量合计"]]
    ordered = front + middle + tail + other
    result = result.loc[:, ordered]

    # 数值列转数字，方便 Excel
    for col in [
        "在途订单量",
        "已预约交货量",
        "客户库存",
        "当前可采购量",
        "订单数量",
        NEW_COL,
        "仓库结存",
        *date_cols,
    ]:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce")

    return result


def read_customer_raw(path: Path) -> pd.DataFrame:
    """预约与库存报表带多层表头，始终按无表头读取。"""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return pd.read_excel(path, header=None, dtype=object, engine="xlrd")
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, header=None, dtype=object, engine="openpyxl")
    raise ValueError("客户文件请使用 Excel（.xls 或 .xlsx）")


def read_company(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return pd.read_excel(path, dtype=object, engine="xlrd")
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, dtype=object, engine="openpyxl")
    raise ValueError("公司文件请使用 Excel（.xlsx）")


def run_merge(customer_path: Path, company_path: Path, export_path: Path) -> Path:
    customer = pre_customer(read_customer_raw(customer_path))
    company = read_company(company_path)
    if "客户货号" not in company.columns:
        raise ValueError("公司文件不像「主生产计划」，请确认选对了文件（需含「客户货号」列）")

    result = process_data(customer, company)
    export_path.parent.mkdir(parents=True, exist_ok=True)
    if export_path.suffix.lower() == ".csv":
        result.to_csv(export_path, index=False, encoding="utf-8-sig")
    else:
        if export_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            export_path = export_path.with_suffix(".xlsx")
        result.to_excel(export_path, index=False, engine="openpyxl")
    return export_path

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x360")
        self.resizable(False, False)

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.export_path = tk.StringVar(value="选择导出路径")

        pad = {"padx": 12, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        ttk.Label(frm, text="客户文件（预约与库存报表 .xls/.xlsx）：").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(frm, textvariable=self.file1_path, width=72).grid(row=1, column=0, columnspan=2, sticky="we")
        ttk.Button(frm, text="浏览", command=lambda: self.select_file(self.file1_path)).grid(row=1, column=2, padx=6)

        ttk.Label(frm, text="公司文件（主生产计划 .xlsx）：").grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Entry(frm, textvariable=self.file2_path, width=72).grid(row=3, column=0, columnspan=2, sticky="we")
        ttk.Button(frm, text="浏览", command=lambda: self.select_file(self.file2_path)).grid(row=3, column=2, padx=6)

        ttk.Label(frm, text="导出路径：").grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Entry(frm, textvariable=self.export_path, width=72).grid(row=5, column=0, columnspan=2, sticky="we")
        ttk.Button(frm, text="选择路径", command=self.select_export_path).grid(row=5, column=2, padx=6)

        ttk.Button(frm, text="开始处理", command=self.process_files).grid(row=6, column=0, sticky="w", pady=16)
        ttk.Label(
            frm,
            text="说明：「投产减入库」= 未完工制令的(投产−入库)；按供应料号=客户货号对应；可为负",
            foreground="#444",
        ).grid(row=7, column=0, columnspan=3, sticky="w")

        self.status = ttk.Label(frm, text="准备就绪", relief=tk.SUNKEN, anchor="w")
        self.status.grid(row=8, column=0, columnspan=3, sticky="we", pady=(18, 0))
        frm.columnconfigure(0, weight=1)

    def select_file(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[
                ("Excel", "*.xls;*.xlsx;*.xlsm"),
                ("全部", "*.*"),
            ],
        )
        if path:
            var.set(path)

    def select_export_path(self) -> None:
        default_name = datetime.now().strftime("%Y-%m-%d") + "-智能排产.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择导出路径",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("全部", "*.*")],
        )
        if path:
            self.export_path.set(path)

    def process_files(self) -> None:
        f1 = self.file1_path.get().strip()
        f2 = self.file2_path.get().strip()
        export = self.export_path.get().strip()
        if not f1 or not f2:
            messagebox.showwarning("提示", "请先选择客户文件和公司文件")
            return
        if not export or export == "选择导出路径":
            messagebox.showwarning("提示", "请先选择导出路径")
            return

        try:
            self.status.config(text="处理中，请稍候…")
            self.update_idletasks()
            out = run_merge(Path(f1), Path(f2), Path(export))
            self.status.config(text="处理完成")
            messagebox.showinfo("成功", f"文件已成功导出到：\n{out}")
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            self.status.config(text="错误发生")
            messagebox.showerror("错误", f"处理过程中发生错误：\n{exc}")


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
