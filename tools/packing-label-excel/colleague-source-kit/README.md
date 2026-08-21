# 益佳通标签生成器（源码）

电脑小工具：填少量出货信息，按装箱规则生成与《益佳通标签模板》同结构的 Excel（内箱 / 外箱 / 卡板二维码字符串）。

## 怎么改规则

用 Cursor 或 WorkBuddy **打开本文件夹**，用中文描述要改什么即可。

最常改的两个数字在 `label_gen.py`：

- `INNER_PER_OUTER`：几个内箱 = 1 个外箱（当前 4）
- `INNER_PER_PALLET_QR`：几个内箱 = 1 张卡板码（当前 12）

## 本机运行

```text
pip install -r requirements.txt
python app.py
```

或双击 `打开标签生成器.bat`。

## 文件说明

| 文件 | 说明 |
|------|------|
| `app.py` | 填表界面 |
| `label_gen.py` | 装箱与生成 Excel |
| `益佳通标签模板.xlsx` | 原模板对照 |
| `requirements.txt` | 依赖 |
| `请先读我.txt` | 给非程序员的中文说明 |
