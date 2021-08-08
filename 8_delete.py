from openpyxl import load_workbook
wb = load_workbook("sample5_2.xlsx")
ws = wb.active

# 삭제 특정 정보 줄을 제거

# ws.delete_rows(8) # 8번째에 있는 7번 정보 제거
# wb.save("sample_delete_rows.xlsx")

# ws.delete_rows(8, 3) # 8번째 줄부터 3줄제거
# wb.save("sample_delete_rows_2.xlsx")


# ws.delete_cols(2) # 2번째 B열 삭제
# wb.save("sample_delete_cols.xlsx")

ws.delete_cols(2, 2) # 2번째 열부터 2줄 삭제
wb.save("sample_delete_cols_2.xlsx")