from openpyxl import load_workbook
wb = load_workbook("sample5_2.xlsx")
ws = wb.active

# 삽입을 하면 셀 한줄이 추가됨
# # ws.insert_rows(8)
# wb.save("sample_insert_rows.xlsx")

# ws.insert_rows(8, 5) # 5줄 삽입
# wb.save("sample_insert_rows_2.xlsx")

# ws.insert_cols(2) # B 열에 추가됨
# wb.save("sample_insert_cols.xlsx")

ws.insert_cols(2, 3) # B 열부터 3칸 추가됨
wb.save("sample_insert_cols_2.xlsx")

