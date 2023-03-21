from PIL import ImageGrab
import win32com.client as win32
import os



excel = win32.gencache.EnsureDispatch('Excel.Application')
curr = os.getcwd() + "/일괄등록테스트.xlsx"
workbook = excel.Workbooks.Open(curr)



# 이미지 아닌 데이터들 뽑기
for sheet in workbook.Worksheets:
    used_range = sheet.UsedRange
    for r in range(used_range.Row + 1, used_range.Row + used_range.Rows.Count):
        for c in range(used_range.Column, used_range.Column + used_range.Columns.Count):
            cell = sheet.Cells(r, c)
            print(cell.Value, cell.Address)
# workbook.Close()



# 이미지 추출 및 저장
for sheet in workbook.Worksheets:
    for i, shape in enumerate(sheet.Shapes):
        if shape.Name.startswith('Picture'):  
            workbook_name = workbook.Name.split(".")[0]
            sheet_name = sheet.Name
            cell_position = shape.TopLeftCell.Address
            shape.Copy()
            image = ImageGrab.grabclipboard()
            print(image)
            image=image.convert("RGB")
            image.save(f"pictures/{workbook_name}_{sheet_name}_{cell_position}.jpg", 'jpeg')
            print(f"Image {i+1} is in cell {cell_position}")
    workbook.Close()



# 참고 코드
# from PIL import ImageGrab
# import win32com.client as win32

# excel = win32.gencache.EnsureDispatch('Excel.Application')
# workbook = excel.Workbooks.Open(r'C:\Users\file.xlsx')

# for sheet in workbook.Worksheets:
#     for i, shape in enumerate(sheet.Shapes):
#         if shape.Name.startswith('Picture'):  # or try 'Image'
#             shape.Copy()
#             image = ImageGrab.grabclipboard()
#             image.save('{}.jpg'.format(i+1), 'jpeg')