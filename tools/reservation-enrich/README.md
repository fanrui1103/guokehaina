# reservation-enrich

预约与库存报表 + 主生产计划：自动补列并做欠料 / 库存预警提示。

- 预约报表补数：`启动预约报表补数.bat` → `python reservation_mps_enrich.py`
- 排产加投产减入库：`启动排产加投产减入库.bat` → `python add_diff_column.py`
- 依赖：pandas、openpyxl 等（与实习项目环境一致即可）
- 案例说明：[docs/case-reservation.md](../../docs/case-reservation.md)
