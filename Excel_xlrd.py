# encoding: utf-8
"""
Author: 沙振宇
CreateTime: 2019-12-6
UpdateTime: 2019-12-6
Info: 读取Excel文件,可以读取xls，也可以读xlsx
"""
import xlrd

def read_excel():
    excel_tables = []
    file_name = 'file/模板.xls'
    workbook = xlrd.open_workbook(file_name) # 打开文件
    sheet = workbook.sheet_by_index(0) # 根据sheet索引或者名称获取sheet内容 sheet索引从0开始
    print(file_name, sheet.name, sheet.nrows, sheet.ncols) # sheet的名称，行数，列数
    print("获取第2行内容:", sheet.row_values(1))
    print("获取第3列内容:", sheet.col_values(2))

    for rown in range(sheet.nrows):
        array = {'L1': '', 'L2': '', 'Question': '', 'Answer': '', 'Similar':''}
        array['L1'] = sheet.cell_value(rown, 0)
        array['L2'] = sheet.cell_value(rown, 1)
        array['Question'] = sheet.cell_value(rown, 2)
        array['Answer'] = sheet.cell_value(rown, 4)
        array['Similar'] = sheet.cell_value(rown, 6)
        excel_tables.append(array)
    print("一共%d条数据"%len(excel_tables))
    return excel_tables

if __name__ == '__main__':
    # 读取Excel
    read_excel()