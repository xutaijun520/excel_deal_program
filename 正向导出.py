import xlrd
import xlwt
import datetime
from pypinyin import pinyin, lazy_pinyin, Style
import time
from xlutils.copy import copy

while True:

    # 打开已存在的Excel文件进行读取
    bookname = input("请输入源文件名:")
    workbook = xlrd.open_workbook(f'{bookname}.xlsx')
    sheet = workbook.sheet_by_index(0) # 假设我们只处理第一个sheet

    # 创建一个新的Excel文件来写入数据
    new_workbook = xlwt.Workbook()
    new_sheet = new_workbook.add_sheet('Sheet 1')

    header = sheet.row_values(0)
    header.append("分组")
    header.append("单位#编码")
    header.append("转换类型")
    for col_index, header_cell in enumerate(header):
        new_sheet.write(0, col_index, header_cell)

    for row_index in range(1,sheet.nrows):
        # 读取当前行的内容
        original_row = sheet.row_values(row_index)
        cut = 2000
        original_row.append( row_index%cut+1 if row_index>=cut else row_index%cut)
        pinyin_without_tone = lazy_pinyin(original_row[8], style=Style.NORMAL)
        original_row.append(pinyin_without_tone)
        original_row.append('转换前')
        #original_row.append(datetime.datetime.now().strftime('%Y-%m-%d'))
        # 将原始行写入新的Excel文件
        for col_index, cell_value in enumerate(original_row):
            new_sheet.write(row_index * 2 - 1, col_index, cell_value) # 写入原始行
        
        # 对原始行的数据进行修改        
        modified_row = list(original_row)
        # 假设这里我们修改第一个单元格的值，您可以根据需要进行修改
        
        modified_row[4] = 'YJ2024'
        if original_row[5]!="":
            modified_row[5] = '2024/1/1'
        if original_row[6] !="":
            modified_row[6] = '9999/12/31'
        modified_row[len(modified_row)-1]='转换后'
        modified_row.append(datetime.datetime.now().strftime('%Y-%m-%d'))
        
        # 将修改后的行写入新的Excel文件
        for col_index, cell_value in enumerate(modified_row):
            if row_index >= 1:
                new_sheet.write(row_index * 2 , col_index, cell_value) # 写入修改后的行

    # 保存新的Excel文件
    new_workbook.save('modified_file.xls')

    time.sleep(1)


    # 打开现有的Excel文件以及工作表
    workbook_reader = xlrd.open_workbook('modified_file.xls')
    sheet_reader = workbook_reader.sheet_by_index(0) # 这个是读取第一个工作表，index从0开始

    # 创建一个新的工作簿和工作表
    old_workbook = xlrd.open_workbook('批号调整单_test.xls')
    old_sheet = old_workbook.sheet_by_index(0)

    # 使用xlutils.copy复制旧的workbook对象
    new_workbook = copy(old_workbook)
    new_sheet = new_workbook.get_sheet(0)


    start_col = 20 # 设置目标工作表的起始列（从0开始索引）
    for col_index in range(sheet_reader.ncols):
        if col_index == 18:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 20, cell_value)
        if col_index == 20:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 21, cell_value)
        if col_index == 0:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 22, cell_value)
        if col_index == 4:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 25, cell_value)
        if col_index == 19:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 28, cell_value)
        if col_index == 8:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 29, cell_value)
        if col_index == 7:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 30, cell_value)
        if col_index == 3:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 31, cell_value)
        if col_index == 10:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 33, cell_value)
        if col_index == 5:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 44, cell_value)
        if col_index == 6:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 45, cell_value)
        if col_index == 21:
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                new_sheet.write(row_index+1, 37, cell_value)
        # if col_index == 2:
        #     for row_index in range(1,sheet_reader.nrows):
        #         # 读取单元格数据
        #         #cell_value = sheet_reader.cell_value(row_index, col_index)
        #         # 将数据写入新工作表的指定列
        #         new_sheet.write(row_index+1, 33, "KCZT01_SYS")
        if col_index == 2:
            mingxi_num = int(input("请输入明细序号:"))
            for row_index in range(1,sheet_reader.nrows):
                # 读取单元格数据
                #cell_value = sheet_reader.cell_value(row_index, col_index)
                # 将数据写入新工作表的指定列
                
                new_sheet.write(row_index+1, 19, str(mingxi_num))
                mingxi_num+=1
        if col_index == 2:
            cell_value1 = sheet_reader.cell_value(col_index, 12)
            cell_value2 = sheet_reader.cell_value(col_index, 13)
            cell_value3 = sheet_reader.cell_value(col_index, 14)
            new_sheet.write(2, 13, cell_value1)
            new_sheet.write(2, 14, cell_value2)
            new_sheet.write(2, 15, cell_value3)

        # if col_index == 2:
        #     cangku = input("请输入仓库编码：")
        #     cangku.strip(" ")
        #     for row_index in range(1,sheet_reader.nrows):
        #         # 读取单元格数据
        #         #cell_value = sheet_reader.cell_value(row_index, col_index)
        #         # 将数据写入新工作表的指定列
                
        #         new_sheet.write(row_index+1, 31, str(cangku))
    # 保存这个新的工作簿
    new_workbook.save(f'modified{bookname}.xlsx')