from PIL import ImageGrab
import win32com.client as win32
import os
from pprint import pprint

excel = win32.gencache.EnsureDispatch('Excel.Application')
curr = os.getcwd() + "/일괄등록테스트.xlsx"
workbook = excel.Workbooks.Open(curr)

# 셀마다 dict 생성
cell_dicts = []
for sheet in workbook.Worksheets:
    used_range = sheet.UsedRange
    for r in range(used_range.Row + 1, used_range.Row + used_range.Rows.Count):
        cell_dict = {"img_path": None}
        for c in range(used_range.Column, used_range.Column + used_range.Columns.Count):
            cell = sheet.Cells(r, c)
            key = sheet.Cells(1, c).Value  # 첫 번째 열을 키값으로 사용
            value = cell.Value
            cell_dict[key] = value
        cell_dicts.append(cell_dict)

# 이미디 추출 및 저장
for sheet in workbook.Worksheets:
    for i, shape in enumerate(sheet.Shapes):
        if shape.Name.startswith('Picture'):  
            workbook_name = workbook.Name.split(".")[0]
            sheet_name = sheet.Name
            cell_position = shape.TopLeftCell.Address
            shape.Copy()
            image = ImageGrab.grabclipboard()
            image=image.convert("RGB")
            image.save(f"pictures/{workbook_name}_{sheet_name}_{cell_position}.jpg", 'jpeg')
            # 이미지 경로 dict에 저장하기
            cell_dict = cell_dicts[shape.TopLeftCell.Row - 2]  # -2 because the first row is excluded and 0-based indexing
            cell_dict["img_path"] = f"pictures/{workbook_name}_{sheet_name}_{cell_position}.jpg"

pprint(cell_dicts)

workbook.Close()
