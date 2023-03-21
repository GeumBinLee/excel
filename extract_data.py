from PIL import ImageGrab
import win32com.client as win32


def getExcelData(path) :
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(path)

    # 셀마다 dict 만들기
    cell_dicts = []
    for sheet in workbook.Worksheets:
        used_range = sheet.UsedRange # 사용중인 셀만 뽑아온다.
        for r in range(used_range.Row + 1, used_range.Row + used_range.Rows.Count):
            cell_dict = {"img_path": None}
            for c in range(used_range.Column, used_range.Column + used_range.Columns.Count):
                cell = sheet.Cells(r, c)
                key = sheet.Cells(1, c).Value  # 첫 번째 열을 key값으로 사용
                if "Image" not in key:
                    value = cell.Value
                    cell_dict[key] = value
            cell_dicts.append(cell_dict)

    # 이미지 추출 및 저장
    for sheet in workbook.Worksheets:
        for i, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith('Picture'):  
                workbook_name = workbook.Name.split(".")[0]
                sheet_name = sheet.Name
                cell_position = shape.TopLeftCell.Address
                shape.Copy() # 복사
                image = ImageGrab.grabclipboard() # 붙여넣기
                image=image.convert("RGB")
                image.save(f"pictures/{workbook_name}_{sheet_name}_{cell_position}.jpg", 'jpeg') # 이미지 저장
                # 이미지 경로 삽입
                cell_dict = cell_dicts[shape.TopLeftCell.Row - 2]  # 인덱스는 0부터 시작하고 첫 번째 행 제외해야 하니까 -2 해주기기
                cell_dict["img_path"] = f"pictures/{workbook_name}_{sheet_name}_{cell_position}.jpg"

    workbook.Close() # 엑셀 닫기
    return cell_dicts
