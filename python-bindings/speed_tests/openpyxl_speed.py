import openpyxl
import os

from openpyxl.styles import Border, Side, Alignment

# Параметры стилей
thin = Side(border_style="thin", color="000000")
border = Border(top=thin, left=thin, right=thin, bottom=thin)
alignment = Alignment(horizontal="center", vertical="center")


base_dir = os.path.dirname(os.path.abspath(__file__))

wb = openpyxl.open(os.path.join(base_dir, "../../test/100mb.xlsx"))
ws = wb['Tablo3']
def generate_table(width: int, height: int):
    table = [['a' + str(i + j) * 10 for i in range(width)] for j in range(height)]
    return table

data = generate_table(5, 200)
for row in data:
    ws.append(row)
# Границы диапазона (в твоем случае: A9930:E99529)
start_row = 99930
end_row = start_row + len(data) - 1
start_col = 1  # A
end_col = 5    # E

# Применение стилей ко всем ячейкам в диапазоне
for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
    for cell in row:
        cell.border = border
        cell.alignment = alignment

wb.save(os.path.join(base_dir, "100mb_openpyxl.xlsx"))
