# -*- coding: utf-8 -*-
"""本地网页小程序：上传欠料 / 库存 / 排产表，生成带标注的欠料对照表。"""

from __future__ import annotations

import io
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for

from compare import compare_files, to_excel_bytes

ROOT = Path(__file__).resolve().parent
app = Flask(__name__)
app.secret_key = "lu-shortage-compare-local"
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024

LAST_RESULT: dict = {}


def _save_upload(fs):
    if fs is None or not fs.filename:
        return None
    return io.BytesIO(fs.read())


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", result=None, summary=None, preview=None)


@app.route("/compare", methods=["POST"])
def do_compare():
    shortage = _save_upload(request.files.get("shortage"))
    inventory = _save_upload(request.files.get("inventory"))
    plan = _save_upload(request.files.get("plan"))

    if shortage is None or inventory is None:
        flash("请至少上传「欠料表」和「库存汇总表」。")
        return redirect(url_for("index"))

    try:
        result, summary = compare_files(shortage, inventory, plan)
    except Exception as exc:
        flash(str(exc))
        return redirect(url_for("index"))

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    LAST_RESULT["df"] = result
    LAST_RESULT["name"] = f"江西奥海欠料_对照结果_{stamp}.xlsx"

    preview_cols = [
        c
        for c in [
            "物料编码",
            "物料名称K3",
            "总欠料",
            "抵扣后的最终欠料",
            "成品库存",
            "库存判断",
            "库存缺口",
            "已排产数量",
            "还需生产",
            "还需排产数量",
            "标注说明",
        ]
        if c in result.columns
    ]
    preview = result[preview_cols].fillna("").astype(str).to_dict(orient="records")
    return render_template("index.html", result=True, summary=summary, preview=preview, columns=preview_cols)


@app.route("/download")
def download():
    df = LAST_RESULT.get("df")
    if df is None:
        flash("还没有生成结果，请先点「开始比对」。")
        return redirect(url_for("index"))
    data = to_excel_bytes(df)
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=LAST_RESULT.get("name", "欠料对照结果.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main():
    url = "http://127.0.0.1:5000"
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    print(f"小程序已启动，浏览器将打开：{url}")
    print("用完后，在这个窗口按 Ctrl+C 可以关闭。")
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
