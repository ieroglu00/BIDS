import pyodbc

driver="{SQL Server}"
server="ben-rds-ss-bidsdb-nonprod-east.ckc8jx9o1aev.us-east-1.rds.amazonaws.com"
database="BIDS_DEV"
username="etluser"
password="4xbEyPFqfNjjTaTW6C"

con = pyodbc.connect("DRIVER="+driver+";SERVER="+server+";DATABASE="+database+";UID="+username+";PWD="+password)
#                      )
# con = (
#     r'DRIVER={SQL Server};'
#     r'SERVER=NUCHPH\\SQLEXPRESS;'
#     r'DATABASE=BOOKSTORE;'
#     r'UID=sa;'
#     r'PWD=Test123;'
#     )
#con = pyodbc.connect(driver='{SQL Server}', Server='ben-rds-ss-bidsdb-nonprod-east.ckc8jx9o1aev.us-east-1.rds.amazonaws.com;', database='BIDS_DEV;', user='etluser;', password='4xbEyPFqfNjjTaTW6C;')
# con = pyodbc.connect(
#     r'Driver={SQL Server};Server=ben-rds-ss-bidsdb-nonprod-east.ckc8jx9o1aev.us-east-1.rds.amazonaws.com;Database=BIDS_DEV;Trusted_Connection=yes;user=etluser;password=4xbEyPFqfNjjTaTW6C',
#     autocommit=True)
cnxn = pyodbc.connect(con)
cursor = con.cursor()
sql_query =  'SELECT * FROM Students'
cursor.execute(sql_query)

for row in cursor:
    print(row)

# def read(conn):
#     print("Read")
#     cursor = conn.cursor()
#     cursor.execute("select * from dummy")
#     for row in cursor:
#         print(f'row = {row}')
#     print()
#
# def create(conn):
#     print("Create")
#     cursor = conn.cursor()
#     cursor.execute(
#         'insert into dummy(a,b) values(?,?);',
#         (3232, 'catzzz')
#     )
#     conn.commit()
#     read(conn)
#
# def update(conn):
#     print("Update")
#     cursor = conn.cursor()
#     cursor.execute(
#         'update dummy set b = ? where a = ?;',
#         ('dogzzz', 3232)
#     )
#     conn.commit()
#     read(conn)
#
# def delete(conn):
#     print("Delete")
#     cursor = conn.cursor()
#     cursor.execute(
#         'delete from dummy where a > 5'
#     )
#     conn.commit()
#     read(conn)
#
# conn = pyodbc.connect(
#     'Driver={SQL Server};'
#                   'Server=LAPTOP-PPDS6BPG;'
#                   'Database=training;'
#                   'Trusted_Connection=yes;'
# )
#
# read(conn)
# create(conn)
# update(conn)
# delete(conn)
#
# conn.close()