"""出货报告生成工具 — 网页入口（自选文件版）。"""

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

from aql import lookup_aql  # noqa: E402
from po_parser import match_templates, parse_purchase_order  # noqa: E402
from report_gen import GenerateRequest, generate_report  # noqa: E402

app = Flask(__name__, template_folder=str(BUNDLE / "templates"))

_lock = threading.Lock()
_session: dict = {
    "tmpdir": None,
    "po_path": None,
    "po_no": None,
    "items": [],
    "templates": {},
    "uploaded_templates": [],
    "generating": False,
}


def _ensure_tmpdir() -> Path:
    if _session["tmpdir"] is None:
        _session["tmpdir"] = Path(tempfile.mkdtemp(prefix="oqc_tool_"))
    return Path(_session["tmpdir"])


def _reset_session_files() -> None:
    old = _session.get("tmpdir")
    _session["po_path"] = None
    _session["po_no"] = None
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
    """弹出系统「选择文件夹」对话框。"""
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
                "po_no": _session.get("po_no"),
                "po_name": Path(_session["po_path"]).name if _session.get("po_path") else None,
                "items": _session.get("items") or [],
                "template_count": len(_session.get("templates") or {}),
                "generating": bool(_session.get("generating")),
            }
        )


@app.post("/api/upload-order")
def api_upload_order():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "请选择采购订单 PDF"}), 400
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "订单必须是 PDF 文件"}), 400

    with _lock:
        if _session.get("generating"):
            return jsonify({"ok": False, "error": "正在生成中，请稍候再换订单"}), 409

        tmp = _ensure_tmpdir()
        uploaded_raw = [Path(p) for p in (_session.get("uploaded_templates") or []) if Path(p).exists()]

        dest = tmp / ("order_" + Path(f.filename).name)
        f.save(dest)
        try:
            po = parse_purchase_order(dest)
        except Exception as e:
            return jsonify({"ok": False, "error": f"无法解析采购订单：{e}"}), 400
        if not po.items:
            return jsonify({"ok": False, "error": "订单里没有识别到物料，请确认文件是否正确"}), 400

        _session["po_path"] = str(dest)
        _session["po_no"] = po.po_no
        _session["items"] = [
            {
                "part_no": it.part_no,
                "name": it.name,
                "qty": it.qty,
                "has_template": False,
                "template": None,
            }
            for it in po.items
        ]

        if uploaded_raw:
            matched = match_templates(uploaded_raw, [it.part_no for it in po.items])
            _session["templates"] = {k: str(v) for k, v in matched.items()}
            _session["uploaded_templates"] = [str(p) for p in uploaded_raw]
            _refresh_item_match()
        else:
            _session["templates"] = {}

        return jsonify(
            {
                "ok": True,
                "po_no": po.po_no,
                "po_name": Path(f.filename).name,
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
            return jsonify({"ok": False, "error": "请先选择采购订单"}), 400

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

        return jsonify(
            {
                "ok": True,
                "matched": len(_session["templates"]),
                "uploaded": len(all_paths),
                "unmatched_files": unmatched,
                "items": _session["items"],
            }
        )


@app.post("/api/preview-aql")
def api_preview_aql():
    data = request.get_json(force=True) or {}
    try:
        qty = int(data.get("deliver_qty") or 0)
        a = lookup_aql(qty)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    return jsonify(
        {
            "ok": True,
            "letter": a.letter,
            "sample_size": a.sample_size,
            "maj_ac": a.maj_ac,
            "maj_re": a.maj_re,
            "min_ac": a.min_ac,
            "min_re": a.min_re,
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
                return jsonify({"ok": False, "error": "请先上传出货报告模板"}), 400

            po_no = (_session.get("po_no") or data.get("po_no") or "").strip()
            snap_dir = Path(tempfile.mkdtemp(prefix="oqc_gen_"))
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
                return jsonify({"ok": False, "error": "模板文件已丢失，请重新上传"}), 400
            _session["generating"] = True

        out_dir = _pick_output_dir()
        if not out_dir:
            return jsonify({"ok": False, "error": "已取消：未选择保存位置"}), 400

        results = []
        errors = []
        for row in rows:
            part_no = (row.get("part_no") or "").strip()
            try:
                order_qty = int(row.get("order_qty"))
                deliver_qty = int(row.get("deliver_qty"))
            except (TypeError, ValueError):
                errors.append({"part_no": part_no, "error": "数量必须是整数"})
                continue
            if deliver_qty < 2:
                errors.append({"part_no": part_no, "error": "发货数量至少为 2"})
                continue

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
                        order_qty=order_qty,
                        deliver_qty=deliver_qty,
                    )
                )
                results.append(
                    {
                        "part_no": part_no,
                        "output": info["output"],
                        "filename": Path(info["output"]).name,
                        "aql": info["aql"],
                        "ship_date": info["ship_date"],
                        "test_date": info["test_date"],
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
    url = "http://127.0.0.1:8765"
    print("=" * 50)
    print("出货报告生成工具已启动")
    print("请用浏览器操作：自己选择订单和模板文件")
    print(f"打开: {url}")
    print("不要关闭本窗口（关闭即退出工具）")
    print("=" * 50)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    app.run(host="127.0.0.1", port=8765, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
