# -*- coding: utf-8 -*-
import openpyxl, sys
sys.path.insert(0, 'src')

EXCEL_PATH = "data/TO Score.xlsx"
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

print("=== 원장 전체 ===")
ws = wb['원장']
for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
    if any(c is not None for c in row):
        print(row[:9])

print("\n=== 입금현황 전체 멤버 납부 합계 ===")
ws2 = wb['입금현황']
rows = list(ws2.iter_rows(values_only=True))
header = rows[0]
month_cols = [(i, str(header[i])) for i in range(3, len(header))
              if header[i] and isinstance(header[i], str) and '년' in header[i]]
for row in rows[1:]:
    name = row[0]
    if not name or not isinstance(name, str) or name.strip() == '계':
        continue
    vals = [row[i] for i, _ in month_cols if row[i]]
    if vals:
        print(f"  {name}: {vals}")
