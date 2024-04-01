import xlrd
import xlwt
from xlutils.copy import copy

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
            new_sheet.write(row_index+1, 32, cell_value)
    if col_index == 10:
        for row_index in range(1,sheet_reader.nrows):
            # 读取单元格数据
            cell_value = sheet_reader.cell_value(row_index, col_index)
            # 将数据写入新工作表的指定列
            new_sheet.write(row_index+1, 34, cell_value)
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
# 保存这个新的工作簿
new_workbook.save('modified_existing_file.xls')