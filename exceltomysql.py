import pymysql
pymysql.install_as_MySQLdb()
import MySQLdb
import xlrd
from time import sleep
from datetime import datetime

start_time = datetime.now()
FILE_NAME = input('File name only, no need file extention, \nExample: customer_data: ')
database = MySQLdb.connect(host="localhost",db="test")
# database = MySQLdb.connect(host="localhost" , user="root" , passwd="root" ,db="mysql")

cursor = database.cursor()
table_name = 'product_and_details'
#Creating the table
product_details_table = (f"CREATE TABLE IF NOT EXISTS {table_name}(\
                            id int,product_id varchar(255) NOT NULL,\
                                product_name text,product_price varchar(255),\
                                    product_rating BLOB,product_star_rating float,\
                                        product_url LONGTEXT, PRIMARY KEY (product_id))")

cursor.execute(product_details_table)
sleep(1)
print(f'\nPreparing {table_name}...')
sleep(1)
print(f'{table_name} is Prepared.')
REMOVE_EXTENTION = FILE_NAME.split('.')[0]
excel_sheet = xlrd.open_workbook(f'{REMOVE_EXTENTION}.xlsx')
print(f'Inserting Data into {table_name}')
sheet_name = excel_sheet.sheet_names()
# Creating insert Query
insert_query = "INSERT INTO product_and_details (\
    product_id,product_name,\
        product_price,product_rating,\
        product_star_rating,product_url) \
            VALUES (%s,%s,%s,%s,%s,%s)"
count = 0
for sh in range(0,len(sheet_name)):
    sheet= excel_sheet.sheet_by_index(sh)
    # Inserting Data
    for row in range(1,sheet.nrows):
        product_id = sheet.cell(row,0).value
        product_name = sheet.cell(row,1).value
        product_price = sheet.cell(row,2).value
        product_rating = sheet.cell(row,3).value
        product_star_rating = sheet.cell(row,4).value
        product_url = sheet.cell(row,5).value
        product_details_value = (product_id,product_name,product_price,product_rating,product_star_rating,product_url)
        cursor.execute(insert_query,product_details_value)
        database.commit()
        count += 1
sleep(.5)
end_time = datetime.now()
duration = end_time-start_time
print(f'{count} of Data inserterd successfully in {table_name} and took {duration}')