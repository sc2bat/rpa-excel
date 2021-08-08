from openpyxl import Workbook
from openpyxl.drawing.image import Image
wb = Workbook()
ws = wb.active

img = Image("img.png")

# 이미지 삽입
ws.add_image(img, "C3")

wb.save("sample16.xlsx")

# ImportError : You must install Pillow to fetch image...
# pip install Pillow