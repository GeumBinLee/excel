
from PIL import ImageGrab
import win32com.client as win32 # Windows COM (Component Object Model) objects 다루기
import os
excel = win32.gencache.EnsureDispatch('Excel.Application') # 엑셀의 인스턴스 생성
curr = os.getcwd() + "./일괄등록테스트.xlsx" # 현재 경로에 있는 엑셀 파일 경로
workbook = excel.Workbooks.Open(curr) # curr 경로에 있는 엑셀 파일 열기



for sheet in workbook.Worksheets: # 엑셀 파일 내의 워크시트를 돈다.
    for i, shape in enumerate(sheet.Shapes): # 워크시트 내 object를 돈다
        if shape.Name.startswith('Picture'): #shape 이름이 Picture로 시작하는지 검사
            workbook_name = workbook.Name.split(".")[0] # 엑셀파일 이름 추출
            sheet_name = sheet.Name # 엑셀시트 이름 추출
            cell_position = shape.TopLeftCell.Address # 셀 위치 추출
            shape.Copy() # 현재 쉐잎(shape)을 클립보드에 복사
            image = ImageGrab.grabclipboard() # 클립보드에 있는 이미지를 변수 image에 할당
            print(image)
            image=image.convert("RGB") # 이미지를 rgb color가진 애로 치환
            image.save(f"{workbook_name}_{sheet_name}_{cell_position}.jpg", 'jpeg')
            print("{i+1}번째 이미지는 {cell_position}에 위치합니다.")
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