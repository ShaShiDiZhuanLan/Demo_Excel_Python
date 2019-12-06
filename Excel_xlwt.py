# encoding: utf-8
"""
Author: 沙振宇
CreateTime: 2019-12-6
UpdateTime: 2019-12-6
Info: 写Excel文件 xlwt 中生成的xls文件最多能支持 65536 行数据。
"""
import xlwt

def write_excel():
    myWorkbook = xlwt.Workbook() # 创建Excel工作薄
    mySheet = myWorkbook.add_sheet('A Test Sheet') # 添加Excel工作表
    myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')   # 写入数据 数据格式
    mySheet.write(1, 2, 1234.56, myStyle)
    mySheet.write(2, 0, 1)                          #写入A3，数值等于1
    mySheet.write(2, 1, 1)                          #写入B3，数值等于1
    mySheet.write(2, 2, xlwt.Formula("A3+B3"))      #写入C3，数值等于2（A3+B3）

    new_path = 'file/生成的Excel_XLWT.xls'
    myWorkbook.save(new_path) # 保存
    print(new_path)

if __name__ == '__main__':
    # 写入Excel
    write_excel()