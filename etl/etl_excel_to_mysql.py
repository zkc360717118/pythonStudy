# -*- coding: utf-8 -*-
import xlrd
from datetime import datetime


def read_excel():
    #获取文件
    # ExcelFile = xlrd.open_workbook(r'E:\root\pythonStudy\etl\kevin.xlsx')
    ExcelFile = xlrd.open_workbook(r'E:\root\pythonStudy\etl\Sample-File-utf.xlsx')

    # 获取目标EXCEL文件sheet名
    # print(ExcelFile.sheet_names())

    # 若有多个sheet，则需要指定读取目标sheet例如读取sheet2

    # sheet2_name=ExcelFile.sheet_names()[1]
    # print(sheet2_name);

    # 获取sheet内容【1.根据sheet索引2.根据sheet名称】

    sheet=ExcelFile.sheet_by_index(0)
    # sheet = ExcelFile.sheet_by_name('TestCase002')

    # 打印sheet的名称，行数，列数
    # print(sheet.name, sheet.nrows, sheet.ncols)

    insert_sql = "insert into test.seaHorse (id,PublicationTitle,PublicationLink,Abstract,Authors,JournalName,PublicationDate,ResearchArea,CellLine,Product,Part,CellTypes,Assay,Species,IsolationMethod,CellSeedingDensity,PlateCoating,Medium,Concentration,NormalizationMethods,Language) values \n"
    with open('lara_mysql.sql', 'a', encoding='utf-8') as w:
        w.write(insert_sql)

    print(sheet.nrows)

    for row in range(0,sheet.nrows):
        eachRow=""
        for point in range(0,21):
            eachcell = sheet.row_values(row)[point]

            if point == 6: #这一行有日期
                eachcell= (xlrd.xldate_as_datetime(eachcell,0).strftime("%Y/%d/%m"))
            elif point == 0:
                eachcell = (str(eachcell)[0:-2])


            # 如果字符串里面有引号和括号 要专一掉
            eachcell = eachcell.strip().replace('"','\\"')
            # eachcell = eachcell.replace('(','\\(')
            # eachcell = eachcell.replace(')','\\)')
            eachRow+='"'+eachcell+'",'
        #去掉拼接多余的逗号
        eachRow = eachRow.rstrip(",")
        #在前后加上括号
        eachRow= "("+eachRow+")"
        # eachRow += "\n"

        # if row == 14:
        #     print(eachRow)

        with open('lara_mysql.sql', 'a', encoding='utf-8') as w:
            if row <= sheet.nrows - 2:
                eachRow = eachRow + ","
            w.write(eachRow)
            w.write("\n")
    # print(insert_sql)


read_excel()
