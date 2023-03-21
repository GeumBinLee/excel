import pymysql, os
from . import extract_data

db = pymysql.connect(
    user='root',
    password='dltjdud9811!',
    host='localhost',
    db='test')

cursor = db.cursor(pymysql.cursors.Cursor)


# df = pd.read_excel(f"{os.getcwd()}\일괄등록테스트.xlsx")
# Data = []
# for i in range(1, len(df.columns)-3) :
#     Data.append(df.iloc[i].to_json())



sql = "insert into new_table values(NULL, '데모 브랜드', 'sku_101882090001', 'kingsman', 'sku_101882090001', '부제', '골든서클', 29000, 4387, '과세', '20세기스튜디오', '영화', '액션', 'https://pupple-images.s3.ap-northeast-2.amazonaws.com/product/_demo-202201061953511641498831-스크린샷 2022-01-07 오전 4.53.22.png', 33, '010-0000-0000', 'BULK_MIG', NULL)"



# sql = "select * from new_table"

# cursor.execute(query = sql)

# result = cursor.fetchall()
# print(result)