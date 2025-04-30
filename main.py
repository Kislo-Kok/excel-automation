# excel_auto.py
import openpyxl
wb = openpyxl.Workbook()
sheet = wb.active
sheet["A1"] = "Товар"
sheet["B1"] = "Цена"
sheet["A2"] = "Кофе"
sheet["B2"] = 300
wb.save("price_list.xlsx")