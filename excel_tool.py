#!/usr/bin/env python
# -*- encoding: utf-8 -*-

'''
excel 文件操作函数
包括excel文件得序列化/反序列化,成绩计算和排序
'''

import xlrd
import xlwt
import pandas as pd

_type_dict = { 0 : int, 1 : str, 2 : int }
title_row = 0
data_row = 4
header_row_level1 = 1
credit_row = 2
credit_col = 4
header_row_level2 = 3

# '221综测表.xlsx'
def serialize_excel(is_write, input_file, output_file):
    #到处excel文件内容至txt
    def export_excel_txt(excel_file, txt_file):
        #将读出的excel数据写入到txt文件中
        def write_excel_txt(sheets_data, txt_file):
            #覆盖写权限，没有文件时创建文件
            file = open(txt_file, '+w', encoding='utf-8')
            for k, table_data in sheets_data.items():
                #将每一个sheet对应的表数据对应sheet的名字写入到字典中 sheet_name : { table_data }
                file.write('%s:{\n' % (k))
                for row_index in range(len(table_data)):
                    for data in table_data[row_index]:
                        #每行数据以空格间隔
                        file.write('%s ' % (data))
                    #单行数据写入结束 \n 换行
                    file.write('\n')
                #以 ‘}’ 作为一页表格数据的结束
                file.write('}\n')
            #文件操作结束
            file.close()

        #打开表格
        impexcel = xlrd.open_workbook(excel_file)
        #创建字典以sheet名称作为键值保存每页的表格数据
        table_row_data = {}
        for sn in impexcel.sheet_names():
            table = impexcel.sheet_by_name(sn)
            sheet_data = []
            #读出单页的所有数据
            for i in range(table.nrows):
                row_data = []
                #读出单行数据
                for data_index in range(len(table.row_values(i))):
                    if table.row_values(i)[data_index] == '':
                        #如果单元格内没有内容提供，替换为字符串None写入至文件中
                        row_data.append('None')
                    else:
                        if i >= data_row and data_index in _type_dict.keys():
                            #保证数据格式
                            row_data.append(_type_dict[data_index](table.row_values(i)[data_index]))
                        else:
                            row_data.append(table.row_values(i)[data_index])
                sheet_data.append(row_data)
                table_row_data[k] = sheet_data
        #按照 格式：每行数据以空格间隔，表头一样 写入到文件中
        write_excel_txt(table_row_data, output_file)
    #将txt文件内容导入至excel
    def import_excel_txt(txt_file, excel_file):
        #计算智测成绩
        def calculate(table_datas):
            for n, table in table_datas.items():
                for row_index in range(len(table)):
                    #非数据行，尾部插入新列
                    if row_index == title_row \
                        or row_index == credit_row \
                        or row_index == header_row_level2:
                        #插入空单元格
                        table[row_index].append('')
                        continue
                    elif row_index == header_row_level1:
                        #插入标题
                        table[row_index].append('智测成绩')
                        continue
                    #替换原行数据的新行数据，计算并插入智测
                    row_sum = 0
                    row_total = 0
                    for col_index in range(len(table[row_index])):
                        #从成绩列开始计算智测
                        if col_index >= credit_col and not isinstance(table[row_index][col_index],str):
                            row_sum += table[row_index][col_index]
                            row_total += table[credit_row][col_index] * table[row_index][col_index]
                    table[row_index].append(round(row_total/row_sum,3) if row_sum != 0 else 0)
            return table_datas
        def _sort(table_datas):
            #排序
            for n,table in table_datas.items():
                for i in range(len(table)):
                    if i < data_row:
                        continue
                    #降序排列
                    for j in range(i, len(table)):
                        if table[i][len(table[i])-1] < table[j][len(table[j])-1]:
                            #如果成绩高则交换位置
                            table[j],table[i] = table[i],table[j]
                #重新设置序号
                for row_index in range(len(table)):
                    if row_index >= data_row:
                        table[row_index][0] = str(row_index - data_row + 1)
                        #转换学号格式
                        table[row_index][2] = str(table[row_index][2])
            return table_datas
        new_table_row_data = {}
        file = open(txt_file, 'r', encoding='utf-8')
        file_date = file.readlines()
        sheet_name = 'None'
        parse_index = 0 #标识解析的行号 注意:这个行号是单个表的行号
        for row_text in file_date:
            #除去每行的换行符
            row_text = row_text.rstrip('\n')
            if ':{' in row_text:
                #判断如果是一个页的开始,则将一个空的链表存入到excel文件数据的字典内
                key_line = row_text.split(':')
                #给页名赋值，表示找到有效格式，可以开始解析
                sheet_name = key_line[0]
                new_table_row_data[sheet_name] = []
                continue
            if '}' in row_text:
                #某页数据导出结束
                sheet_name = 'None'
                parse_index = 0
                continue
            if sheet_name == 'None':
                #没有解析到有效表格数据跳过
                continue
            #运行至该处表示有效解析到页内容，开始解析页内容
            #根据空格取出每个单元格的内容
            datas = row_text.split(' ')
            #出去末尾无效数据
            del datas[-1]
            #是否需要切换行的解析方式
            text_line_jump = False
            row_data = []
            for row_index in range(len(datas)):
                if parse_index < data_row:
                    #如果为非学生的行则按照表头解析
                    text_line_jump = True
                    if parse_index == credit_row and row_index >= credit_col:
                        #学分格式将学分转换为int格式
                        row_data.append(int(float(datas[row_index])))
                    else:
                        #字符串内容，直接存入
                        row_data.append('' if datas[row_index] == 'None' else datas[row_index])
                else:
                    if row_index in _type_dict.keys():
                        #需要特殊处理的单元格内容
                        row_data.append(_type_dict[row_index](datas[row_index]))
                        continue
                    if datas[row_index] != 'None' and datas[row_index] != '':
                        #学生成绩单元格，为有效内容时，转为float类型
                        row_data.append(datas[row_index] if row_index == 3 else float(datas[row_index]))
                    else:
                        row_data.append('')
            if text_line_jump:
                parse_index += 1
            new_table_row_data[sheet_name].append(row_data)
        #计算并插入智测
        new_table_row_data = calculate(new_table_row_data)
        new_table_row_data = _sort(new_table_row_data)

        for n, table in new_table_row_data.items():
            for row in table:
                print(row)

        df_list = []
        for n, table in new_table_row_data.items():
            df_list.append(pd.DataFrame(table))

        for df in df_list:
            df = df.reset_index(drop=True)
            df.to_excel(excel_file,header=None,index=False)
            #for i in df.values()
    #序列化
    if is_write:
        export_excel_txt(input_file, output_file)
    else:
        import_excel_txt(input_file, output_file)



#serialize_excel(False,'temp_data.txt','test.xlsx')