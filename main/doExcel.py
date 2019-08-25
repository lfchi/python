import xlrd
import xlwt
from xlutils.copy import copy

file_path = 'mydoc.xlsx'
xlrd.book.ensure_unicode = 'utf-8'  # 设置编码
data = xlrd.open_workbook(file_path)  # 打开文件
new_book = copy(data)
sheet_names = data.sheet_names()  # 获取文件中所包含的Sheet名
for n in range(len(sheet_names)):
    m_t = data.sheet_by_index(n)
    m_cols_count = m_t.ncols  # 取总行数
    data1 = ""
    if sheet_names[n] == 'Sheet1':
        for n_r in range(0, m_cols_count + 1):
            data1 = m_t.cell(n_r, 1).value
            data2 = m_t.cell(n_r, 2).value
            print("test data1= " + data1 + ",  m_cols_count = " + str(m_cols_count))
            for n_r_p in range(0, data.sheet_by_index(1).ncols + 1):
                if data1 == data.sheet_by_index(1).cell(n_r_p, 1).value:
                    sheet = new_book.get_sheet(1)  # 获取第一个表格的数据
                    sheet.write(n_r_p, 2, data2)  # 修改0行1列的数据为'Haha'
                    print("test data2= " + data2)
                    new_book.save('secondsheet.xls')  # 保存新的excel

for n in range(len(sheet_names)):
    m_table = data.sheet_by_index(n)
    m_rows_count = m_table.nrows  # 取总行数
    m_cols_count = m_table.ncols  # 取总列数
    m_row_data = m_table.row_values(0)
    m_col_data = m_table.col_values(0)
    print(sheet_names[n])
    print(m_row_data, m_col_data)
    for row in range(0, m_rows_count):
        for col in range(0, m_cols_count):
            data1 = m_table.cell(row, col).value
            print(data1, end=' ')
        print('\n')
table = data.sheet_by_index(0)
rows_count = table.nrows  # 取总行数
cols_count = table.ncols  # 取总列数
#print(rows_count, cols_count)  # 4行4列

row_data = table.row_values(0)  # 取第一行的数据
col_data = table.col_values(0)  # 取第一列的数据
# print(row_data, col_data)  # ['张三', '仙剑奇侠传', 'aaa', 'Beautiful', 20180806.0] ['张三', '李四', '王五', '雷六']

cell_data = row_data[0]  # 获取第0行第0列的值
cell_data_A1 = table.cell(1, 1).value  # 获取第一行第一列的值
# print(cell_data, cell_data_A1)  # 张三 西游记  注意下标从0开始

for row in range(0, rows_count):
    for col in range(0, cols_count):
        data1 = table.cell(row, col).value
        # print(data1, end='')
    # print('\n')
#