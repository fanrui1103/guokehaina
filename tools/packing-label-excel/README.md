# packing-label-excel

装箱标签 Excel 生成器：按总件数与每内箱件数展开内箱 / 外箱 / 卡板码字符串。

当前规则摘要：
- 外箱：每 **4** 个内箱一组（尾组单独成箱）
- 卡板码：每 **12** 个内箱一张（由 `label_gen.py` 里 `INNER_PER_PALLET_QR` 控制）

- 入口：`打开标签生成器.bat` 或 `python app.py`
- 依赖：`pip install -r requirements.txt`
- `colleague-source-kit/`：给同事改参数用的源码说明包
- `portfolio-notes/`：需求澄清与协作实录
