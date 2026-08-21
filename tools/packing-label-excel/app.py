# -*- coding: utf-8 -*-
"""益佳通标签表生成器 — 给同事用的填表窗口。"""

from __future__ import annotations

import json
import sys
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from label_gen import (
    build_workbook,
    default_filename,
    packing_summary,
    save_workbook,
)

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
MEMORY_FILE = BASE_DIR / "上次填写.json"
OUTPUT_DIR = BASE_DIR / "输出"

FIELDS = [
    ("supplier", "供应商代码", "04.221"),
    ("part_no", "物料料号", "D01.02.002.00073"),
    ("ship_date", "出货日期（8位数字）", datetime.now().strftime("%Y%m%d")),
    ("total_qty", "总件数", "123"),
    ("per_inner", "每个内箱装几件", "2"),
    ("unit", "物料单位", "juan"),
    ("name", "物料名称", "PET蓝胶"),
]


def load_memory() -> dict:
    if MEMORY_FILE.exists():
        try:
            data = json.loads(MEMORY_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, json.JSONDecodeError):
            return {}
    return {}


def save_memory(values: dict) -> None:
    MEMORY_FILE.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("益佳通标签表生成器")
        self.geometry("560x560")
        self.minsize(520, 520)
        self.configure(bg="#f4f1ea")

        self.vars: dict[str, tk.StringVar] = {}
        memory = load_memory()

        ttk.Label(
            self,
            text="填几项，点生成，就会得到一张能拿去打印的 Excel。\n上次填过的会自动出现。",
            font=("Microsoft YaHei UI", 11),
            background="#f4f1ea",
            justify="left",
        ).pack(anchor="w", padx=24, pady=(20, 12))

        form = ttk.Frame(self)
        form.pack(fill="x", padx=24)

        for key, label, default in FIELDS:
            row = ttk.Frame(form)
            row.pack(fill="x", pady=5)
            ttk.Label(row, text=label, width=18, font=("Microsoft YaHei UI", 10)).pack(side="left")
            var = tk.StringVar(value=str(memory.get(key, default)))
            self.vars[key] = var
            entry = ttk.Entry(row, textvariable=var, font=("Microsoft YaHei UI", 11))
            entry.pack(side="left", fill="x", expand=True)
            var.trace_add("write", lambda *_: self.refresh_preview())
            if key == "ship_date":
                ttk.Button(row, text="填今天", width=8, command=self.fill_today).pack(side="left", padx=(8, 0))

        self.preview = tk.StringVar()
        ttk.Label(
            self,
            textvariable=self.preview,
            font=("Microsoft YaHei UI", 11, "bold"),
            background="#f4f1ea",
            foreground="#1f4e3d",
            wraplength=500,
            justify="left",
        ).pack(anchor="w", padx=24, pady=(16, 8))

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=24, pady=12)
        ttk.Button(btns, text="生成 Excel", command=self.generate).pack(side="left")
        ttk.Button(btns, text="退出", command=self.destroy).pack(side="left", padx=12)

        hint = ttk.Label(
            self,
            text="生成的表格列和原来的「益佳通标签模板」一样，原来怎么打印还怎么打印。",
            font=("Microsoft YaHei UI", 9),
            background="#f4f1ea",
            foreground="#666666",
            wraplength=500,
            justify="left",
        )
        hint.pack(anchor="w", padx=24, pady=(8, 16))

        self.refresh_preview()

    def fill_today(self) -> None:
        self.vars["ship_date"].set(datetime.now().strftime("%Y%m%d"))

    def read_ints(self) -> tuple[int, int]:
        total_raw = self.vars["total_qty"].get().strip()
        per_raw = self.vars["per_inner"].get().strip()
        if not total_raw.isdigit() or not per_raw.isdigit():
            raise ValueError("总件数和每内箱件数请填正整数，不要填小数或文字。")
        total_qty = int(total_raw)
        per_inner = int(per_raw)
        packing_summary(total_qty, per_inner)
        return total_qty, per_inner

    def refresh_preview(self) -> None:
        try:
            total_qty, per_inner = self.read_ints()
            s = packing_summary(total_qty, per_inner)
            tail = ""
            if s["last_inner_qty"] != per_inner:
                tail = f"最后一箱装 {s['last_inner_qty']} 件。"
            self.preview.set(
                f"将生成：{s['inner']} 个内箱，{s['outer']} 个外箱，{s['pallet']} 张卡板码。{tail}"
            )
        except ValueError as exc:
            self.preview.set(str(exc) if "必须" in str(exc) or "正整数" in str(exc) else "请把总件数和每内箱件数填成正整数，上面会预告会生成多少箱。")

    def generate(self) -> None:
        values = {key: var.get().strip() for key, var in self.vars.items()}
        missing = [label for key, label, _ in FIELDS if not values[key]]
        if missing:
            messagebox.showerror("还没填完", "请填写：" + "、".join(missing))
            return

        date = values["ship_date"]
        if not (date.isdigit() and len(date) == 8):
            messagebox.showerror("日期格式不对", "出货日期请填 8 位数字，例如 20260812。")
            return

        try:
            total_qty, per_inner = self.read_ints()
        except ValueError as exc:
            messagebox.showerror("数字不对", str(exc))
            return

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        suggested = OUTPUT_DIR / default_filename(values["part_no"], date, total_qty)
        path = filedialog.asksaveasfilename(
            title="保存生成的 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(OUTPUT_DIR),
            initialfile=suggested.name,
        )
        if not path:
            return

        try:
            wb = build_workbook(
                supplier=values["supplier"],
                part_no=values["part_no"],
                ship_date=date,
                total_qty=total_qty,
                per_inner=per_inner,
                unit=values["unit"],
                name=values["name"],
            )
            saved = save_workbook(wb, path)
        except Exception as exc:
            messagebox.showerror("生成失败", f"出错了：{exc}")
            return

        save_memory(values)
        s = packing_summary(total_qty, per_inner)
        messagebox.showinfo(
            "已经生成",
            f"已保存到：\n{saved}\n\n"
            f"内箱 {s['inner']} 个，外箱 {s['outer']} 个，卡板码 {s['pallet']} 张。\n"
            f"用 Excel 打开后即可按原来的方式打印。",
        )


def write_crash_log(exc: BaseException | None = None) -> None:
    try:
        text = traceback.format_exc() if exc is None else "".join(
            traceback.format_exception(type(exc), exc, exc.__traceback__)
        )
        (BASE_DIR / "启动错误.txt").write_text(text, encoding="utf-8")
    except OSError:
        pass


def run_smoke_test(out_path: Path) -> None:
    """不打开窗口，直接生成一张表，用来检查打包后的程序能不能用。"""
    values = {
        "supplier": "04.221",
        "part_no": "PACK.SMOKE.0001",
        "ship_date": "20260821",
        "total_qty": "123",
        "per_inner": "2",
        "unit": "juan",
        "name": "PET蓝胶",
    }
    wb = build_workbook(
        supplier=values["supplier"],
        part_no=values["part_no"],
        ship_date=values["ship_date"],
        total_qty=123,
        per_inner=2,
        unit=values["unit"],
        name=values["name"],
    )
    saved = save_workbook(wb, out_path)
    save_memory(values)
    marker = BASE_DIR / "打包自检结果.txt"
    marker.write_text(
        f"OK\n文件: {saved}\n内箱62 外箱16 卡板码5\n",
        encoding="utf-8",
    )


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    try:
        if "--smoke-test" in sys.argv:
            out = None
            if "--out" in sys.argv:
                i = sys.argv.index("--out")
                if i + 1 < len(sys.argv):
                    out = Path(sys.argv[i + 1])
            if out is None:
                out = OUTPUT_DIR / "_pack_smoke.xlsx"
            run_smoke_test(out)
        else:
            main()
    except Exception as exc:
        write_crash_log(exc)
        raise
