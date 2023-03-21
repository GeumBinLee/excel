import pymysql
import os
import extract_data

curr = os.getcwd() + "/일괄등록테스트.xlsx"

db = pymysql.connect(
    user='root',
    password='dltjdud9811!',
    host='localhost',
    db='test')

cursor = db.cursor(pymysql.cursors.Cursor)

results = extract_data.getExcelData(curr)
print(results)

for result in results:
    sql = f"INSERT INTO new_table VALUES (NULL, '{result['BRAND_NM']}', '{result['PRIMARY_PROD_ID']}', '{result['PROD_ID']}', '{result['OPT_ID']}', '{result['OPT_NM_1']}', '{result['OPT_VALUE_1']}', '{result['PROD_PRICE']}', '{result['PROD_COST']}', '{result['TAX_CLF']}', '{result['PUPPLE_CATE_L']}', '{result['PUPPLE_CATE_M']}', '{result['PUPPLE_CATE_S']}', '{result['PROD_IMG_URL']}', '{result['PROD_INFO_ID']}', '{result['AS_PHONE_NUM']}', '{result['DELIV_POLICY_NM']}', '{result['img_path']}' )"
    cursor.execute(query=sql)

db.commit()
db.close()


pang = cursor.fetchall()
print(pang)