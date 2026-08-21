# shortage-compare

上传欠料表 / 库存表 / 主生产计划，生成带「交货分类」标注的欠料对照结果（本地网页小工具）。

结果会区分：
1. 库存就够交货
2. 库存不够但已排产
3. 库存和已排产都不够

- 入口：`打开小程序.bat` 或 `python app.py`
- 依赖：`pip install -r requirements.txt`
- 同事使用说明见 `使用说明.txt`
- 重新打包可参考 `打包.bat`（生成的 exe / `_internal` 不纳入本仓库）
