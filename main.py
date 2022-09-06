import xlrd
import requests

bk = xlrd.open_workbook("data.xls")  # 收集结果文件名自行更改
sheet = bk.sheet_by_index(0)

for i in range(1, sheet.nrows):
    name = sheet.cell_value(i, 2)
    pic_name = f"pic/{name}.jpg"
    url = sheet.hyperlink_map.get((i, 4)).url_or_path
    print(f"{i}/{sheet.nrows - 1} {name}")
    with requests.get(url) as res:
        with open(pic_name, "wb") as f:
            f.write(res.content)
