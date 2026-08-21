# -*- coding: utf-8 -*-
"""怡富万物料标示单批量生成工具：按你选的模板生成，只改序列号和二维码。"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from generator import (
    GenerateRequest,
    TemplateInfo,
    generate_to_file,
    inspect_qr_in_file,
    list_label_sheets,
    load_template,
    parse_serial_spec,
    planned_qr_list,
    summarize_serials,
    template_summary,
)


def _windows_desktop() -> Path:
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        ) as key:
            return Path(winreg.QueryValueEx(key, "Desktop")[0])
    except Exception:
        for name in ("Desktop", "桌面"):
            p = Path.home() / name
            if p.exists():
                return p
        return Path.home() / "Desktop"


DEFAULT_OUT = _windows_desktop() / "怡富万物料标签.xlsx"
DEFAULT_TEMPLATE = Path(
    r"c:\Users\fr189\Documents\WXWork\1688855726876583\Cache\File\2026-08\怡富万新标签.xlsx"
)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("怡富万物料标示单生成工具")
        self.geometry("780x720")
        self.minsize(700, 620)

        self.template_path = tk.StringVar(
            value=str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE.exists() else ""
        )
        self.sheet_name = tk.StringVar(value="")
        self.serial_spec = tk.StringVar(value="1-12")
        self.output = tk.StringVar(value=str(DEFAULT_OUT))

        self._template: TemplateInfo | None = None
        self._sheet_combo: ttk.Combobox | None = None

        self._build()
        if self.template_path.get():
            self.after(200, self._load_template_silent)

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 4}
        root = ttk.Frame(self)
        root.pack(fill=tk.BOTH, expand=True, **pad)

        ttk.Label(
            root,
            text="每次先选模板 → 程序读取料号品名等 → 你自己填写要生成哪些序列号。旧料号固定留空。",
            wraplength=740,
        ).pack(anchor="w", pady=(0, 8))

        tpl = ttk.LabelFrame(root, text="① 选择模板（换一批货就换一份模板）")
        tpl.pack(fill=tk.X, pady=4)
        tpl.columnconfigure(1, weight=1)

        ttk.Label(tpl, text="模板文件").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(tpl, textvariable=self.template_path).grid(
            row=0, column=1, sticky="ew", padx=6, pady=4
        )
        ttk.Button(tpl, text="浏览…", command=self._browse_template).grid(
            row=0, column=2, padx=6, pady=4
        )

        ttk.Label(tpl, text="工作表").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self._sheet_combo = ttk.Combobox(tpl, textvariable=self.sheet_name, state="readonly")
        self._sheet_combo.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        self._sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._reload_sheet())
        ttk.Button(tpl, text="重新读取", command=self._reload_sheet).grid(
            row=1, column=2, padx=6, pady=4
        )

        serial_box = ttk.LabelFrame(root, text="② 自己决定要哪些序列号（自由填写）")
        serial_box.pack(fill=tk.X, pady=8)
        serial_box.columnconfigure(0, weight=1)

        ttk.Label(
            serial_box,
            text="用逗号分隔多段。可以写区间，也可以写单个号，也可以重复同一段。",
            wraplength=740,
        ).grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))

        e = ttk.Entry(serial_box, textvariable=self.serial_spec)
        e.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        e.bind("<KeyRelease>", lambda _e: self._refresh_qr_preview())

        ttk.Label(
            serial_box,
            text="例子：1-24,1-24,1-24　　或　　1-12,20-30,5　　或　　只写 1-24",
            wraplength=740,
        ).grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))

        out_box = ttk.LabelFrame(root, text="③ 保存位置")
        out_box.pack(fill=tk.X, pady=4)
        out_box.columnconfigure(0, weight=1)
        ttk.Entry(out_box, textvariable=self.output).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Button(out_box, text="浏览…", command=self._browse_output).grid(
            row=0, column=1, padx=6, pady=6
        )

        mid = ttk.Panedwindow(root, orient=tk.VERTICAL)
        mid.pack(fill=tk.BOTH, expand=True, pady=4)

        info_frame = ttk.LabelFrame(mid, text="从模板读到的内容（只读，换模板会变）")
        qr_frame = ttk.LabelFrame(mid, text="二维码内容预览（看这里就能确认每张是否在变）")
        mid.add(info_frame, weight=1)
        mid.add(qr_frame, weight=1)

        self.info_text = tk.Text(info_frame, height=10, wrap="word", font=("Microsoft YaHei", 9))
        self.info_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.info_text.insert("1.0", "请先选择模板文件。")
        self.info_text.configure(state="disabled")

        qr_wrap = ttk.Frame(qr_frame)
        qr_wrap.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.qr_text = tk.Text(qr_wrap, height=10, wrap="none", font=("Consolas", 9))
        yscroll = ttk.Scrollbar(qr_wrap, orient="vertical", command=self.qr_text.yview)
        self.qr_text.configure(yscrollcommand=yscroll.set)
        self.qr_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        btns = ttk.Frame(root)
        btns.pack(fill=tk.X, pady=8)
        ttk.Button(btns, text="生成 Excel", command=self._generate).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="检查已生成文件的二维码", command=self._inspect_file).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btns, text="退出", command=self.destroy).pack(side=tk.RIGHT, padx=4)

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择物料标签模板 Excel",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(Path(self.template_path.get()).parent)
            if self.template_path.get()
            else str(_windows_desktop()),
        )
        if not path:
            return
        self.template_path.set(path)
        self._load_template()

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存生成的标签",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=Path(self.output.get()).name or "怡富万物料标签.xlsx",
            initialdir=str(Path(self.output.get()).parent),
        )
        if path:
            self.output.set(path)

    def _load_template_silent(self) -> None:
        try:
            self._load_template()
        except Exception:
            pass

    def _load_template(self) -> None:
        path = self.template_path.get().strip()
        if not path:
            raise ValueError("请先选择模板文件")
        sheets = list_label_sheets(path)
        if not sheets:
            raise ValueError("这个 Excel 里没有可用的工作表")
        self._sheet_combo["values"] = sheets
        # 尽量保留当前选择，否则用第一张标签表
        current = self.sheet_name.get()
        chosen = current if current in sheets else sheets[0]
        self.sheet_name.set(chosen)
        self._reload_sheet()

    def _reload_sheet(self) -> None:
        try:
            path = self.template_path.get().strip()
            if not path:
                return
            info = load_template(path, self.sheet_name.get() or None)
            self._template = info
            self._set_text(self.info_text, template_summary(info))
            # 输出文件名跟料号走，方便辨认
            if info.material_no:
                self.output.set(str(_windows_desktop() / f"怡富万标签-{info.material_no}.xlsx"))
            self._refresh_qr_preview()
        except Exception as exc:
            self._template = None
            self._set_text(self.info_text, f"读取失败：{exc}")
            self._set_text(self.qr_text, "")
            messagebox.showerror("读取模板失败", str(exc))

    def _set_text(self, widget: tk.Text, content: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", content)
        widget.configure(state="disabled")

    def _parse_serials(self) -> list[int]:
        return parse_serial_spec(self.serial_spec.get())

    def _refresh_qr_preview(self) -> None:
        if not self._template:
            self._set_text(self.qr_text, "还没有读到模板。")
            return
        try:
            serials = self._parse_serials()
        except Exception as exc:
            self._set_text(self.qr_text, str(exc))
            return

        all_items = planned_qr_list(self._template, serials)
        show_n = min(len(serials), 30)
        lines = [
            summarize_serials(serials),
            "下面是「表格序列号 → 二维码全文」。相同序号出现多次是正常的（你故意重复写了）。",
            "",
        ]
        for disp, payload in all_items[:show_n]:
            lines.append(f"{disp}  →  {payload}")
        if len(serials) > show_n:
            lines.append("...")
            for disp, payload in all_items[-1:]:
                lines.append(f"{disp}  →  {payload}")
        self._set_text(self.qr_text, "\n".join(lines))

    def _generate(self) -> None:
        try:
            if not self._template:
                self._load_template()
            if not self._template:
                raise ValueError("请先成功读取模板")
            serials = self._parse_serials()
            total = len(serials)
            out = Path(self.output.get().strip())
            if not out.name.lower().endswith(".xlsx"):
                out = out.with_suffix(".xlsx")
            req = GenerateRequest(
                template=self._template,
                serials=serials,
                sheet_name=self._template.material_no or self._template.sheet_name,
            )
            generate_to_file(req, out)
            self._refresh_qr_preview()

            checked = inspect_qr_in_file(out, limit=min(total, 12))
            check_lines = "\n".join(f"· {a}\n  {b}" for a, b in checked[:5])
            messagebox.showinfo(
                "完成",
                f"已生成：\n{out}\n\n"
                f"共 {total} 张。\n{summarize_serials(serials)}\n\n"
                f"抽查二维码（前几张）：\n{check_lines}\n\n"
                f"想看全部：点「检查已生成文件的二维码」。",
            )
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("失败", str(exc))

    def _inspect_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要检查二维码的 Excel",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=Path(self.output.get()).name,
            initialdir=str(Path(self.output.get()).parent)
            if self.output.get()
            else str(_windows_desktop()),
        )
        if not path:
            return
        try:
            rows = inspect_qr_in_file(path, limit=100)
            if not rows:
                messagebox.showwarning("结果", "没有读到二维码（可能文件里没有图，或图无法识别）。")
                return
            # 检查末尾序列是否有重复
            tails = [t.split(";")[-1] for _, t in rows if ";" in t]
            uniq = len(set(tails))
            text = [
                f"文件：{path}",
                f"读到 {len(rows)} 个二维码；末尾序列共 {uniq} 种"
                + ("（都不同，正常）" if uniq == len(tails) else "（有重复，请检查）"),
                "",
            ]
            for loc, payload in rows:
                text.append(f"{loc}")
                text.append(f"  {payload}")
            self._set_text(self.qr_text, "\n".join(text))
            messagebox.showinfo(
                "检查完成",
                f"读到 {len(rows)} 个二维码，末尾序列 {uniq} 种。\n详细内容已显示在下方预览框。",
            )
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("检查失败", str(exc))


def main() -> None:
    here = Path(__file__).resolve().parent
    if str(here) not in sys.path:
        sys.path.insert(0, str(here))
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
