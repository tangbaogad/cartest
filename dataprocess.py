import xlrd
import xlwt
import re

def read_data():
    # 打开文件

    workbook = xlrd.open_workbook(r'X3data.xlsx')
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

    # sheet的名称，行数，列数
    print (sheet1.name, sheet1.nrows, sheet1.ncols)

    # 获取整行和整列的值（数组）
    #rows = sheet1.row_values(3)  # 获取第四行内容
    cols = sheet1.col_values(4)  # 获取第三列内容

    #print(cols)
    print(type(cols[0]))
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('test', cell_overwrite_ok=True)
    f2 = open('X3.txt', 'w',encoding='utf-8')

    j=0
    for i in cols:
        i = re.sub('<!--athm--><!--@athm_BASE64@-->',' ',i)
        i = re.sub('<!--@athm_js@--><br>', ' ', i)
        i = re.sub('<!--@athm_js@-->', ' ', i)
        i = re.sub('<!--@athm_js@-->', ' ', i)
        i = re.sub('<br>', ' ', i)
        pattern0=r'<div class=".*</div>'
        dotall = re.compile(pattern0, re.DOTALL)
        i = re.sub(dotall, ' ', i)
        pattern1 = r'<div class=".*<br><br>'
        dotall = re.compile(pattern1, re.DOTALL)
        i = re.sub(dotall, ' ', i)
        i = re.sub('【', '\n【 ', i)
        sheet.write(j, 0,i)
        f2.write(i +'\n')
        j+=1
        # print(i)
    book.save(r'X3.xls')
    f2.close()
        # # 获取单元格内容
    # print
    # sheet1.cell(1, 0).value.encode('utf-8')
    # print
    # sheet1.cell_value(1, 0).encode('utf-8')
    # print
    # sheet1.row(1)[0].value.encode('utf-8')
    #
    # # 获取单元格内容的数据类型
    # print
    # sheet1.cell(1, 0).ctype
    #
read_data()
