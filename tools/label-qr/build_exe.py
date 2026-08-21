# -*- coding: utf-8 -*-
"""用 PyInstaller 打包成免安装 Python 的可执行程序（给同事用）。"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent.parent
DIST = ROOT / "dist-exe"
RELEASE = REPO / "release" / "怡富万物料标签生成工具"
NAME = "怡富万物料标签生成工具"


def main() -> None:
    subprocess.check_call(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "-q",
            "pyinstaller",
            "openpyxl",
            "qrcode[pil]",
            "Pillow",
            "pyzbar",
        ]
    )

    for p in (ROOT / "build", DIST, ROOT / f"{NAME}.spec"):
        if p.is_file():
            p.unlink()
        elif p.is_dir():
            shutil.rmtree(p, ignore_errors=True)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        NAME,
        "--distpath",
        str(DIST),
        "--workpath",
        str(ROOT / "build"),
        "--specpath",
        str(ROOT),
        "--paths",
        str(ROOT),
        "--hidden-import",
        "generator",
        "--hidden-import",
        "qrcode",
        "--hidden-import",
        "qrcode.image.pil",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "PIL",
        "--collect-all",
        "qrcode",
        "--collect-all",
        "pyzbar",
        str(ROOT / "app.py"),
    ]
    print("Running:", " ".join(cmd))
    subprocess.check_call(cmd)

    exe_dir = DIST / NAME
    exe = exe_dir / f"{NAME}.exe"
    if not exe.exists():
        raise SystemExit(f"打包失败，未找到：{exe}")

    # 同步到 release，方便你直接发给同事
    if RELEASE.exists():
        shutil.rmtree(RELEASE, ignore_errors=True)
    shutil.copytree(exe_dir, RELEASE)

    readme = RELEASE / "给同事的说明.txt"
    readme.write_text(
        "【怡富万物料标签生成工具 — 给同事】\n"
        "\n"
        "1. 必须拷贝「整个文件夹」\n"
        "   - 正确：整个「怡富万物料标签生成工具」文件夹（里面有 exe 和 _internal）\n"
        "   - 错误：只发一个 .exe（会打不开）\n"
        "\n"
        "2. 怎么用\n"
        "   - 双击「怡富万物料标签生成工具.exe」\n"
        "   - 点「浏览」选出你们做好的模板 Excel（.xlsx）\n"
        "   - 选对工作表（一份文件里可能有多种料号）\n"
        "   - 看中间「从模板读到的内容」是否正确\n"
        "   - 在「序列号」框自己填写要哪些号，用逗号分隔\n"
        "     例子：1-24,1-24,1-24\n"
        "     例子：1-12,20-30,5\n"
        "     例子：只写 1-24\n"
        "   - 点「生成 Excel」\n"
        "   - 结果默认在桌面：怡富万标签-料号.xlsx\n"
        "\n"
        "3. 怎么确认二维码有没有变\n"
        "   - 看软件下方预览：每张末尾数字应不同（0001、0002…）\n"
        "   - 或点「检查已生成文件的二维码」\n"
        "   - 或用手机微信扫一扫不同标签\n"
        "\n"
        "4. 说明\n"
        "   - 旧料号固定留空，不用填\n"
        "   - 厂商/料号/品名等全部跟你选的模板走，换模板就换内容\n"
        "   - 每张标签只有序列号和二维码不同\n"
        "\n"
        "5. 若打不开或报错\n"
        "   - 确认是拷贝了整个文件夹，不是单独一个 exe\n"
        "   - 把报错截图发回范瑞排查\n",
        encoding="utf-8",
    )

    print("DONE:", exe)
    print("给同事发这个文件夹:", RELEASE)


if __name__ == "__main__":
    main()
