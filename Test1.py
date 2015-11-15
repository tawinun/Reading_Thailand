import xlrd

path = 'C:\Python34\Project\Test1.xlsx'

workbook = xlrd.open_workbook(path) #เปิดไฟล์
worksheet = workbook.sheet_by_index(0) # ดูเวิคชีท
print(worksheet.cell_value(0, 0)) # ปริ้นค่าแถวที่ ,คอลัมน์ที่

for col in range(worksheet.ncols):
    print(worksheet.cell_value(2, col)) # ปริ้นค่าชื่อคอลัมน์

data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
            row in range(worksheet.nrows)] #เอาข้อมูลยัดใส่ list
print(data_set) # ปริ้นค่าทั้งหมดในตารางออกมา data_set[0][0] เอาแค่บางค่า row col
