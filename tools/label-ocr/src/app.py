"""标签提取填表：选图片 → 识别 → 核对 → 导出 Excel。"""

from __future__ import annotations

import json
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

SRC = Path(__file__).resolve().parent
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from paths import app_dir, bundle_dir  # noqa: E402

ROOT = app_dir()

from excel_writer import default_output_name, write_excel  # noqa: E402
from extractor import extract_many, is_label_image  # noqa: E402

SETTINGS_PATH = ROOT / "上次填写.json"
TEMPLATE_NAME = "雅达物料清单模板.xlsx"
TEMPLATE_CANDIDATES = [
    ROOT / TEMPLATE_NAME,
    ROOT / "模板" / TEMPLATE_NAME,
    bundle_dir() / TEMPLATE_NAME,
]

DISPLAY_COLS = [
    ("name", "图片", 180),
    ("CLID", "标签条形码", 180),
    ("ArtesynPN", "雅达料号", 110),
    ("ManufacturerPN", "制造商料号", 110),
    ("UnitQty", "数量", 70),
    ("DateCode", "生产日期", 100),
    ("ExpDate", "到期日期", 100),
    ("LotNo", "批次号", 90),
    ("COO", "原产地", 60),
    ("Manufacturer", "生产厂家", 90),
    ("MaterialGroup", "物料组", 90),
    ("PONo", "订单号", 100),
    ("POLine", "行号", 50),
]

def find_template() -> Path | None:
    for p in TEMPLATE_CANDIDATES:
        if p.exists():
            return p
    return None


def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(data: dict) -> None:
    SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def fmt_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{value.year}/{value.month}/{value.day}"
    return str(value)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("标签提取填表")
        self.geometry("1180x720")
        self.minsize(960, 600)
        self.configure(bg="#f4f1ea")

        self.image_paths: list[Path] = []
        self.rows: list[dict] = []
        self.template_path = find_template()

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("TLabel", background="#f4f1ea", font=("Microsoft YaHei UI", 10))
        style.configure("TButton", font=("Microsoft YaHei UI", 10))
        style.configure("TLabelframe", background="#f4f1ea")
        style.configure("TLabelframe.Label", background="#f4f1ea", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Hint.TLabel", foreground="#5c5346", font=("Microsoft YaHei UI", 9))
        style.configure("Title.TLabel", background="#f4f1ea", font=("Microsoft YaHei UI", 16, "bold"))

        self._build()
        self._load_defaults()

    def _build(self):
        pad = {"padx": 12, "pady": 6}

        ttk.Label(self, text="标签提取填表", style="Title.TLabel").pack(anchor="w", padx=16, pady=(12, 0))
        ttk.Label(
            self,
            text="把标签照片/截图里的字段读出来，填进雅达物料清单。换一批标签也能用，不依赖某一张标签上的具体数字。",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=16)

        box1 = ttk.LabelFrame(self, text="① 这批共用信息（识别后每行都会带上，表格里还能改）")
        box1.pack(fill="x", padx=12, pady=8)

        grid = ttk.Frame(box1)
        grid.pack(fill="x", **pad)

        self.var_po = tk.StringVar()
        self.var_line = tk.StringVar()
        self.var_invoice = tk.StringVar()
        self.var_unit = tk.StringVar(value="EA")
        self.var_pkg = tk.StringVar(value="1")
        self.var_supplier = tk.StringVar()
        self.var_scode = tk.StringVar()

        def add(row, col, label, var, width=16):
            ttk.Label(grid, text=label).grid(row=row, column=col * 2, sticky="e", padx=(8, 4), pady=4)
            ttk.Entry(grid, textvariable=var, width=width, font=("Microsoft YaHei UI", 10)).grid(
                row=row, column=col * 2 + 1, sticky="w", pady=4
            )

        add(0, 0, "订单号", self.var_po, 18)
        add(0, 1, "订单行号", self.var_line, 10)
        add(0, 2, "送货单号", self.var_invoice, 22)
        add(1, 0, "单位", self.var_unit, 8)
        add(1, 1, "每个条码箱数", self.var_pkg, 8)
        add(1, 2, "供方代码", self.var_scode, 12)
        ttk.Label(grid, text="供方").grid(row=2, column=0, sticky="e", padx=(8, 4), pady=4)
        ttk.Entry(grid, textvariable=self.var_supplier, width=48, font=("Microsoft YaHei UI", 10)).grid(
            row=2, column=1, columnspan=5, sticky="we", pady=4
        )

        box2 = ttk.LabelFrame(self, text="② 选择标签图片，然后点开始识别")
        box2.pack(fill="x", padx=12, pady=4)
        row2 = ttk.Frame(box2)
        row2.pack(fill="x", **pad)
        ttk.Button(row2, text="选择图片（可多选）", command=self.pick_files).pack(side="left", padx=4)
        ttk.Button(row2, text="选择整个文件夹", command=self.pick_folder).pack(side="left", padx=4)
        self.btn_ocr = ttk.Button(row2, text="开始识别", command=self.start_ocr)
        self.btn_ocr.pack(side="left", padx=12)
        self.lbl_files = ttk.Label(row2, text="还没选图片", style="Hint.TLabel")
        self.lbl_files.pack(side="left", padx=8)

        self.progress = ttk.Progressbar(box2, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=(0, 8))
        self.lbl_status = ttk.Label(box2, text="提示：第一次识别会稍慢，之后会快一些。", style="Hint.TLabel")
        self.lbl_status.pack(anchor="w", padx=12, pady=(0, 8))

        box3 = ttk.LabelFrame(self, text="③ 核对结果（双击格子可以改）")
        box3.pack(fill="both", expand=True, padx=12, pady=4)

        tree_frame = ttk.Frame(box3)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=8)
        cols = [c[0] for c in DISPLAY_COLS]
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        for key, title, width in DISPLAY_COLS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_double_click)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=12, pady=10)
        self.btn_export = ttk.Button(bottom, text="导出 Excel", command=self.export_excel)
        self.btn_export.pack(side="left")
        self.lbl_tpl = ttk.Label(bottom, style="Hint.TLabel")
        self.lbl_tpl.pack(side="left", padx=12)
        ttk.Button(bottom, text="更换模板", command=self.pick_template).pack(side="left")
        self._refresh_template_label()

    def _refresh_template_label(self):
        if self.template_path:
            self.lbl_tpl.config(text=f"将写入模板：{self.template_path.name}")
        else:
            self.lbl_tpl.config(text="还没找到模板，请点「更换模板」选 雅达物料清单模板.xlsx")

    def _load_defaults(self):
        s = load_settings()
        self.var_po.set(s.get("PONo", ""))
        self.var_line.set(s.get("POLine", ""))
        self.var_invoice.set(s.get("InvoiceDN", ""))
        self.var_unit.set(s.get("Unit", "EA"))
        self.var_pkg.set(s.get("NoOfPackage", "1"))
        self.var_supplier.set(s.get("supplier", "深圳市海纳宏业科技有限公司"))
        self.var_scode.set(s.get("supplier_code", "101913"))

    def _save_defaults(self):
        save_settings(
            {
                "PONo": self.var_po.get().strip(),
                "POLine": self.var_line.get().strip(),
                "InvoiceDN": self.var_invoice.get().strip(),
                "Unit": self.var_unit.get().strip(),
                "NoOfPackage": self.var_pkg.get().strip(),
                "supplier": self.var_supplier.get().strip(),
                "supplier_code": self.var_scode.get().strip(),
            }
        )

    def _manual(self) -> dict:
        def maybe_num(text: str):
            text = text.strip()
            if text.isdigit():
                return int(text)
            return text

        return {
            "PONo": maybe_num(self.var_po.get()),
            "POLine": maybe_num(self.var_line.get()),
            "InvoiceDN": self.var_invoice.get().strip(),
            "Unit": self.var_unit.get().strip() or "EA",
            "NoOfPackage": maybe_num(self.var_pkg.get() or "1"),
            "supplier": self.var_supplier.get().strip(),
            "supplier_code": maybe_num(self.var_scode.get()),
        }

    def pick_files(self):
        files = filedialog.askopenfilenames(
            title="选择标签图片",
            filetypes=[("图片", "*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp"), ("全部", "*.*")],
        )
        if files:
            self.image_paths = [Path(f) for f in files]
            self.lbl_files.config(text=f"已选 {len(self.image_paths)} 张")

    def pick_folder(self):
        folder = filedialog.askdirectory(title="选择放标签图片的文件夹")
        if not folder:
            return
        paths = [p for p in Path(folder).iterdir() if p.is_file() and is_label_image(p)]
        paths.sort()
        self.image_paths = paths
        self.lbl_files.config(text=f"文件夹内找到 {len(paths)} 张图片")

    def pick_template(self):
        path = filedialog.askopenfilename(
            title="选择物料清单模板",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.template_path = Path(path)
            self._refresh_template_label()

    def start_ocr(self):
        if not self.image_paths:
            messagebox.showinfo("还没选图片", "请先点「选择图片」或「选择整个文件夹」。")
            return
        self._save_defaults()
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.image_paths)
        self.lbl_status.config(text="正在识别，请稍等…")
        self.btn_ocr.state(["disabled"])
        self.btn_export.state(["disabled"])
        threading.Thread(target=self._ocr_worker, daemon=True).start()

    def _ocr_worker(self):
        def progress(i, total, name):
            self.after(
                0,
                lambda i=i, total=total, name=name: self._set_progress(i, total, name),
            )

        try:
            results = extract_many(self.image_paths, progress=progress)
        except Exception as exc:
            self.after(0, lambda exc=exc: self._ocr_failed(exc))
            return
        self.after(0, lambda: self._show_results(results))

    def _ocr_failed(self, exc: Exception):
        self.btn_ocr.state(["!disabled"])
        self.btn_export.state(["!disabled"])
        self.lbl_status.config(text="识别失败")
        messagebox.showerror("识别失败", str(exc))

    def _set_progress(self, i, total, name):
        self.progress["value"] = i
        self.lbl_status.config(text=f"正在识别 {i}/{total}：{name}")

    def _show_results(self, results: list[dict]):
        manual = self._manual()
        self.rows = []
        skipped = 0
        for item in results:
            if not item.get("ok"):
                skipped += 1
            fields = dict(item.get("fields") or {})
            row = {**manual, **fields, "_file": item.get("file"), "name": item.get("name")}
            if item.get("error"):
                row["name"] = f"{row['name']}（失败）"
            self.rows.append(row)
        self._reload_tree()
        ok = sum(1 for r in results if r.get("ok"))
        self.lbl_status.config(text=f"完成：成功 {ok} 张" + (f"，有 {skipped} 张没读全，请双击核对" if skipped else "，请扫一眼再导出"))
        self.btn_ocr.state(["!disabled"])
        self.btn_export.state(["!disabled"])

    def _reload_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows):
            values = [fmt_cell(row.get(k)) for k, *_ in DISPLAY_COLS]
            self.tree.insert("", "end", iid=str(i), values=values)

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        key = DISPLAY_COLS[col_idx][0]
        if key == "name":
            return
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return
        x, y, w, h = bbox
        row_i = int(item)
        current = fmt_cell(self.rows[row_i].get(key))
        entry = tk.Entry(self.tree, font=("Microsoft YaHei UI", 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current)
        entry.focus()

        committed = {"done": False}

        def commit(_=None):
            if committed["done"]:
                return
            committed["done"] = True
            text = entry.get().strip()
            try:
                entry.destroy()
            except tk.TclError:
                pass
            self.rows[row_i][key] = self._coerce(key, text)
            self._reload_tree()
            self.tree.selection_set(str(row_i))

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def _coerce(self, key, text):
        if not text:
            return ""
        if key in ("DateCode", "ManufactureDate", "ExpDate"):
            t = text.replace("-", "/").replace(".", "/")
            parts = t.split("/")
            if len(parts) == 3 and all(p.isdigit() for p in parts):
                return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            return text
        if key in ("UnitQty", "LotNo", "PONo", "POLine", "NoOfPackage", "supplier_code") and text.isdigit():
            return int(text)
        if key in ("ArtesynPN", "ManufacturerPN") and text.isdigit():
            return int(text)
        return text

    def export_excel(self):
        if not self.rows:
            messagebox.showinfo("还没有数据", "请先识别标签。")
            return
        if not self.template_path or not self.template_path.exists():
            messagebox.showinfo("缺少模板", "请先选择「雅达物料清单模板.xlsx」。")
            self.pick_template()
            if not self.template_path:
                return
        path = filedialog.asksaveasfilename(
            title="保存物料清单",
            defaultextension=".xlsx",
            initialfile=default_output_name(),
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        export_rows = []
        skipped = 0
        for row in self.rows:
            item = dict(row)
            if not item.get("CLID") and not item.get("ArtesynPN"):
                skipped += 1
                continue
            if item.get("DateCode"):
                item["ManufactureDate"] = item["DateCode"]
            export_rows.append(item)
        if not export_rows:
            messagebox.showinfo("没有可导出的行", "没有识别到条码或料号，请检查图片是不是标签。")
            return
        if skipped:
            if not messagebox.askyesno("部分跳过", f"有 {skipped} 张不像标签（没有条码/料号），将不写入表格。继续导出其余行吗？"):
                return
        try:
            out = write_excel(self.template_path, path, export_rows)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self._save_defaults()
        if messagebox.askyesno("已保存", f"已保存到：\n{out}\n\n要打开这个文件吗？"):
            try:
                import os

                os.startfile(out)  # type: ignore[attr-defined]
            except Exception:
                pass


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
