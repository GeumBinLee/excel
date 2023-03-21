import pymysql

db = pymysql.connect(
    user='root',
    password='dltjdud9811!',
    host='localhost',
    db='test')

cursor = db.cursor(pymysql.cursors.Cursor)

sql = "select * from new_table"

cursor.execute(query = sql)

result = cursor.fetchall()
print(result)