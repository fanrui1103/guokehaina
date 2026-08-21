# Case: Reservation & Stock Report Enrichment

**工具目录：** [`tools/reservation-enrich`](../tools/reservation-enrich)

## Problem

每天要从客户系统导出「预约与库存报表」，再人工对照「主生产计划」，补订单数量、仓库结存、投产与入库差额，并判断欠料 / 备货过多。表格大、料号多、还要按地点分行，手工易错。

## Approach

1. 读取预约报表 + 主生产计划  
2. 用供应料号 ↔ 客户货号做匹配（含模糊匹配兜底）  
3. 在「5 周订货量」右侧插入：订单数量、仓库结存、投产减入库、提示  
4. 按料号汇总覆盖量，自动标色：  
   - 黄：欠料  
   - 红：欠料且超过阈值阈值阈值覆盖 ×1.2  
   - 蓝：库存预警（备太多）  
5. 只保留「需求」行，同料号聚合，顺序跟当天客户导出一致（越靠前越急）  

同目录另有「排产合并表 + 投产减入库」小工具，服务排产链路。

## Tech

- Python + pandas + openpyxl  
- 跨表键对齐与数值汇总  
- Tkinter 桌面选文件 / 导出  
- 条件格式化（业务可读的风险颜色）  

## Outcome

- 把「每天对表」变成选两个文件 → 一键导出  
- 预警规则产品化，方便跟单按颜色筛选  
- 体现：**数据拼接 + 业务规则引擎 + 交付给非技术用户**  

## What to look at in code

- `reservation_mps_enrich.py` — 主流程与提示规则  
- `smart_schedule.py` — 读表、匹配、投产减入库映射  
- `add_diff_column.py` — 排产表增量列  
