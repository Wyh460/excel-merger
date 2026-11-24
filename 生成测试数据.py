# -*- coding: utf-8 -*-
"""生成测试数据并验证 excel_merger 能正常合并。"""
import os
from openpyxl import Workbook

base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试数据")
os.makedirs(base, exist_ok=True)

# 文件1
wb = Workbook(); ws = wb.active
ws.append(["姓名", "金额", "日期"])
ws.append(["张三", 100, "2026-11-01"])
ws.append(["李四", 200, "2026-11-02"])
ws.append(["张三", 100, "2026-11-01"])  # 重复行
wb.save(os.path.join(base, "表1.xlsx"))

# 文件2（表头顺序不同 + 多一列）
wb2 = Workbook(); ws2 = wb2.active
ws2.append(["日期", "姓名", "金额", "备注"])
ws2.append(["2026-11-03", "王五", 300, "新客户"])
ws2.append(["2026-11-04", "赵六", 400, "新客户"])
wb2.save(os.path.join(base, "表2.xlsx"))

print("测试数据已生成:", os.listdir(base))
