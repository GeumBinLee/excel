import pandas as pd


readData = pd.read_excel("./일괄등록테스트.xlsx")
print(readData)
print()

for i in readData.index:
    print(i, readData['PROD_COST'][i],
          readData['TAX_CLF'][i],
          readData['PUPPLE_CATE_L'][i],
          readData['PUPPLE_CATE_M'][i],
          readData['PUPPLE_CATE_S'][i],
          readData['PROD_IMG_URL'][i],
          readData['PROD_INFO_ID'][i],
          readData['AS_PHONE_NUM'][i],
          readData['DELIV_POLICY_NM'][i],)