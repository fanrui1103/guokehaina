"""送货单出货报告生成工具 — 网页入口（自选 Excel + 模板）。"""

from __future__ import annotations

import atexit
import shutil
import sys
import tempfile
import threading
import webbrowser
from datetime import date
from pathlib import Path

from flask import Flask, jsonify, render_template, request


def _bundle_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = _app_dir()
BUNDLE = _bundle_dir()

if str(BUNDLE) not in sys.path:
    sys.path.insert(0, str(BUNDLE))
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_parser import parse_delivery_excel  # noqa: E402
from report_gen import GenerateRequest, generate_report  # noqa: E402
from template_match import match_templates  # noqa: E402

app = Flask(__name__, template_folder=str(BUNDLE / "templates"))

_lock = threading.Lock()
_session: dict = {
    "tmpdir": None,
    "excel_path": None,
    "excel_name": None,
    "items": [],
    "templates": {},
    "uploaded_templates": [],
    "generating": False,
}


def _ensure_tmpdir() -> Path:
    if _session["tmpdir"] is None:
        _session["tmpdir"] = Path(tempfile.mkdtemp(prefix="dn_oqc_"))
    return Path(_session["tmpdir"])


def _reset_session_files() -> None:
    old = _session.get("tmpdir")
    _session["excel_path"] = None
    _session["excel_name"] = None
    _session["items"] = []
    _session["templates"] = {}
    _session["uploaded_templates"] = []
    _session["tmpdir"] = None
    if old and Path(old).exists():
        shutil.rmtree(old, ignore_errors=True)


def _cleanup_on_exit() -> None:
    with _lock:
        if not _session.get("generating"):
            _reset_session_files()


atexit.register(_cleanup_on_exit)


def _unique_output(folder: Path, part_no: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    base = folder / f"{part_no}.pdf"
    if not base.exists():
        return base
    stamp = date.today().strftime("%Y%m%d")
    candidate = folder / f"{part_no}_{stamp}.pdf"
    if not candidate.exists():
        return candidate
    n = 2
    while True:
        candidate = folder / f"{part_no}_{stamp}_{n}.pdf"
        if not candidate.exists():
            return candidate
        n += 1


def _pick_output_dir() -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        root.wm_attributes("-topmost", 1)
    except Exception:
        pass
    path = filedialog.askdirectory(title="选择出货报告保存位置")
    root.destroy()
    return Path(path) if path else None


def _refresh_item_match() -> None:
    items = _session.get("items") or []
    tpl_map = _session.get("templates") or {}
    for it in items:
        p = tpl_map.get(it["part_no"])
        it["has_template"] = p is not None
        it["template"] = Path(p).name if p else None


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/session")
def api_session():
    with _lock:
        return jsonify(
            {
                "excel_name": _session.get("excel_name"),
                "items": _session.get("items") or [],
                "template_count": len(_session.get("templates") or {}),
                "generating": bool(_session.get("generating")),
            }
        )


@app.post("/api/upload-excel")
def api_upload_excel():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "请选择送货单明细 Excel"}), 400
    name = f.filename.lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls")):
        return jsonify({"ok": False, "error": "请选择 Excel 文件（.xlsx）"}), 400

    with _lock:
        if _session.get("generating"):
            return jsonify({"ok": False, "error": "正在生成中，请稍候再换表格"}), 409

        tmp = _ensure_tmpdir()
        uploaded_raw = [Path(p) for p in (_session.get("uploaded_templates") or []) if Path(p).exists()]

        dest = tmp / ("excel_" + Path(f.filename).name)
        f.save(dest)
        try:
            groups = parse_delivery_excel(dest)
        except Exception as e:
            return jsonify({"ok": False, "error": f"无法读取 Excel：{e}"}), 400
        if not groups:
            return jsonify({"ok": False, "error": "表格里没有识别到物料编码"}), 400

        _session["excel_path"] = str(dest)
        _session["excel_name"] = Path(f.filename).name
        _session["items"] = [
            {
                "part_no": g.part_no,
                "name": g.name,
                "qty": g.qty,
                "po_no": g.po_no,
                "row_count": g.row_count,
                "has_template": False,
                "template": None,
            }
            for g in groups
        ]

        if uploaded_raw:
            matched = match_templates(uploaded_raw, [g.part_no for g in groups])
            _session["templates"] = {k: str(v) for k, v in matched.items()}
            _session["uploaded_templates"] = [str(p) for p in uploaded_raw]
            _refresh_item_match()
        else:
            _session["templates"] = {}

        return jsonify(
            {
                "ok": True,
                "excel_name": _session["excel_name"],
                "item_count": len(_session["items"]),
                "items": _session["items"],
            }
        )


@app.post("/api/upload-templates")
def api_upload_templates():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "error": "请选择至少一个出货报告模板 PDF"}), 400

    with _lock:
        if _session.get("generating"):
            return jsonify({"ok": False, "error": "正在生成中，请稍候再上传模板"}), 409
        if not _session.get("items"):
            return jsonify({"ok": False, "error": "请先选择送货单 Excel"}), 400

        tmp = _ensure_tmpdir()
        saved: list[Path] = []
        for f in files:
            if not f.filename or not f.filename.lower().endswith(".pdf"):
                continue
            dest = tmp / ("tpl_" + Path(f.filename).name)
            if dest.exists():
                dest = tmp / f"tpl_{len(saved)}_{Path(f.filename).name}"
            f.save(dest)
            saved.append(dest)

        if not saved:
            return jsonify({"ok": False, "error": "没有有效的 PDF 模板"}), 400

        prev = [Path(p) for p in (_session.get("uploaded_templates") or []) if Path(p).exists()]
        all_paths = prev + saved
        by_name: dict[str, Path] = {}
        for p in all_paths:
            by_name[p.name] = p
        all_paths = list(by_name.values())

        part_nos = [it["part_no"] for it in _session["items"]]
        matched = match_templates(all_paths, part_nos)
        _session["uploaded_templates"] = [str(p) for p in all_paths]
        _session["templates"] = {k: str(v) for k, v in matched.items()}
        _refresh_item_match()

        used = {Path(v).resolve() for v in _session["templates"].values()}
        unmatched = [p.name for p in all_paths if p.resolve() not in used]
        missing = [it["part_no"] for it in _session["items"] if not it.get("has_template")]

        return jsonify(
            {
                "ok": True,
                "matched": len(_session["templates"]),
                "uploaded": len(all_paths),
                "unmatched_files": unmatched,
                "missing_parts": missing,
                "items": _session["items"],
            }
        )


@app.post("/api/generate")
def api_generate():
    data = request.get_json(force=True) or {}
    rows = data.get("items") or []
    if not rows:
        return jsonify({"ok": False, "error": "请至少选择一种物料"}), 400

    snap_dir: Path | None = None
    try:
        with _lock:
            if _session.get("generating"):
                return jsonify({"ok": False, "error": "正在生成中，请稍候"}), 409
            if not _session.get("templates"):
                return jsonify({"ok": False, "error": "请先选择出货报告模板"}), 400

            item_by_part = {it["part_no"]: it for it in (_session.get("items") or [])}
            snap_dir = Path(tempfile.mkdtemp(prefix="dn_gen_"))
            snap_map: dict[str, Path] = {}
            for part_no, tpl in (_session.get("templates") or {}).items():
                src = Path(tpl)
                if not src.exists():
                    continue
                dest = snap_dir / src.name
                shutil.copy2(src, dest)
                snap_map[part_no] = dest
            if not snap_map:
                shutil.rmtree(snap_dir, ignore_errors=True)
                snap_dir = None
                return jsonify({"ok": False, "error": "模板文件已丢失，请重新选择"}), 400
            _session["generating"] = True

        out_dir = _pick_output_dir()
        if not out_dir:
            return jsonify({"ok": False, "error": "已取消：未选择保存位置"}), 400

        results = []
        errors = []
        for row in rows:
            part_no = (row.get("part_no") or "").strip()
            src_item = item_by_part.get(part_no) or {}
            try:
                qty = int(row.get("qty") if row.get("qty") is not None else src_item.get("qty") or 0)
            except (TypeError, ValueError):
                errors.append({"part_no": part_no, "error": "数量必须是整数"})
                continue
            if qty < 1:
                errors.append({"part_no": part_no, "error": "发货数量合计至少为 1"})
                continue

            po_no = str(row.get("po_no") or src_item.get("po_no") or "").strip()
            tpl = snap_map.get(part_no)
            if not tpl or not tpl.exists():
                errors.append({"part_no": part_no, "error": "没有匹配到该料号的模板"})
                continue

            out = _unique_output(out_dir, part_no)
            try:
                info = generate_report(
                    GenerateRequest(
                        template_path=tpl,
                        output_path=out,
                        po_no=po_no,
                        order_qty=qty,
                        deliver_qty=qty,
                    )
                )
                results.append(
                    {
                        "part_no": part_no,
                        "output": info["output"],
                        "filename": Path(info["output"]).name,
                        "ship_date": info["ship_date"],
                        "qty": qty,
                        "po_no": po_no,
                        "notes": info["notes"],
                    }
                )
            except Exception as e:
                if out.exists():
                    try:
                        out.unlink()
                    except OSError:
                        pass
                errors.append({"part_no": part_no, "error": str(e)})

        return jsonify(
            {
                "ok": len(results) > 0,
                "results": results,
                "errors": errors,
                "output_dir": str(out_dir),
            }
        )
    finally:
        with _lock:
            _session["generating"] = False
        if snap_dir and snap_dir.exists():
            shutil.rmtree(snap_dir, ignore_errors=True)


@app.post("/api/reset")
def api_reset():
    with _lock:
        if _session.get("generating"):
            return jsonify({"ok": False, "error": "正在生成中，请稍后再清空"}), 409
        _reset_session_files()
    return jsonify({"ok": True})


def main():
    url = "http://127.0.0.1:8766"
    print("=" * 50)
    print("送货单出货报告工具已启动")
    print("请用浏览器操作：选择 Excel 送货单 + 各料号 PDF 模板")
    print(f"打开: {url}")
    print("不要关闭本窗口（关闭即退出工具）")
    print("=" * 50)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    app.run(host="127.0.0.1", port=8766, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
