# encoding: utf-8
"""
Author: 沙振宇
CreateTime: 2019-10-30
UpdateTime: 2019-12-6
Info: 读取Txt文件，并写Excel文件
"""
import xlrd
import xlsxwriter

# 读取（模板.xls）Excel文件
def read_old_excel():
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
    print("read_old_excel 一共%d条数据"%len(excel_tables))
    return excel_tables

# 写（生成的Excel.xlsx）Excel文件
def write_old_excel():
    workbook = xlsxwriter.Workbook('file/生成的Excel_main.xlsx')  # 创建一个excel文件
    worksheet = workbook.add_worksheet()  # 在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    worksheet.write('A1', '1级分类')
    worksheet.write('B1', '2级分类')
    worksheet.write('C1', 'FAQ问题')
    worksheet.write('D1', '渠道')
    worksheet.write('E1', 'FAQ回答')
    worksheet.write('F1', '关联FAQ问题')
    worksheet.write('G1', 'FAQ相似问句')

    excel_tables = read_old_excel()
    print("write_old_excel 一共%d条数据"%len(excel_tables))
    cowNum = 1
    for list in excel_tables:
        L1List = list['L1']
        L2List = list['L2']
        QuestionList = list['Question']
        AnswerList = list['Answer']
        Similar = list['Similar']
        SimilarList = Similar.split('||')
        if L1List == "1级分类":
            continue
        # print("SimilarList",SimilarList)
        worksheet.write(cowNum, 0, L1List)
        worksheet.write(cowNum, 1, L2List)
        worksheet.write(cowNum, 2, QuestionList)
        worksheet.write(cowNum, 4, AnswerList)
        for item in SimilarList:
            worksheet.write(cowNum, 6, item)
            cowNum += 1
    workbook.close()

# 读取数据源文件
def read_txt():
    with open('file/data.txt',  "r",encoding="utf-8") as f:  # 设置文件对象
        txtList = f.readlines()  # 可以是随便对文件的操作
    return txtList

# 写（对比.xlsx）Excel文件
def write_excel(txtList):
    workbook = xlsxwriter.Workbook('file/对比.xlsx')  # 创建一个excel文件
    worksheet = workbook.add_worksheet()  # 在文件中创建一个名为TEST的sheet,不加名字默认为sheet1
    worksheet.write('A1', '工单头')
    worksheet.write('B1', 'words')
    worksheet.write('C1', 'cor')
    worksheet.write('D1', 'ins')
    worksheet.write('E1', 'del')
    worksheet.write('F1', 'sub')
    worksheet.write('G1', 'corr')
    worksheet.write('H1', 'cer')
    worksheet.write('I1', 'ref')
    worksheet.write('J1', 'res')

    print("write_excel 一共%d条"%len(txtList))
    value_297 = []
    for index,value in enumerate (txtList):
        if index < 297 :
            curIndex = int(index/3)+1
            if index % 3 == 1:
                str1 = str(value).split("ref:	")[1]
                worksheet.write(curIndex, 8, str1)
            elif index %3 == 2:
                str2 = str(value).split("res:	")[1]
                worksheet.write(curIndex, 9, str2)
            else:
                str_t = str(value)[0:19]
                worksheet.write(curIndex, 0, str_t)
                str_w = str(value).split(str_t)[1].split(") ")
                str_w1 = str_w[0][8:]
                str_w2 = str_w1.split(",")

                # 中间
                str_words = str_w2[0]
                worksheet.write(curIndex, 1, str_words)
                str_cor = str_w2[1].split("=")[1]
                worksheet.write(curIndex, 2, str_cor)
                str_ins = str_w2[2].split("=")[1]
                worksheet.write(curIndex, 3, str_ins)
                str_del = str_w2[3].split("=")[1]
                worksheet.write(curIndex, 4, str_del)
                str_sub = str_w2[4].split("=")[1]
                worksheet.write(curIndex, 5, str_sub)

                str_l = str_w[1].split(",")
                str_corr=str_l[0].split("=")[1]
                worksheet.write(curIndex, 6, str_corr)
                str_cer=str_l[1].split("=")[1]
                worksheet.write(curIndex, 7, str_cer)
        else:
            value_297.append(index)

    print(">= 297: ", value_297)

    workbook.close()

if __name__ == '__main__':
    write_old_excel()

    txtList = read_txt()
    write_excel(txtList)