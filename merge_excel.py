# coding:utf-8
# !/usr/bin/env python
# Author: weiweicai@qudian.com
# Purpose:
# Create: 29/03/2020 08:21


import xlrd
import xlwt
import os
import copy

DATA_SHEET_NAME = '手工数据表'
CITY_LIST = ['福州', '厦门', '宁德', '莆田', '泉州', '漳州', '龙岩', '三明', '南平']


def get_file_name_list(file_path, suffix_list_str='xls,xlsx'):
    suffix_list = suffix_list_str.split(',')
    res = os.listdir(file_path)
    res_file_list = []
    for file_name in res:
        for suffix in suffix_list:
            if file_name.endswith(suffix):
                res_file_list.append(file_name)
    return tuple([os.path.join(file_path, name) for name in res_file_list])


def formate_data(data, start_merge_row_idx, start_merge_col_idx):
    res_data = {}
    # 相同部分
    table_rows = 0
    table_cols = 0
    same_data = []
    for file_name, table in data.items():
        table_rows = table.nrows
        table_cols = table.ncols
        for row_num in range(table_rows):
            same_data.append([])
            for col_num in range(start_merge_col_idx):
                same_data[row_num].append(table.cell(row_num, col_num).value)
        break

    # 不同部分
    for col_num in range(start_merge_col_idx, table_cols):
        file_idx = 0
        for city_name in CITY_LIST:
            if city_name not in data:
                print(f'没有{city_name}的数据')
                continue
            table = data[city_name]
            sheet_name = table.cell(start_merge_row_idx, col_num).value
            if file_idx == 0:
                res_data[sheet_name] = copy.deepcopy(same_data)
            res_data[sheet_name][start_merge_row_idx].append(city_name)
            for row_num in range(start_merge_row_idx + 1, table_rows):
                res_data[sheet_name][row_num].append(
                    table.cell(row_num, col_num).value)
            file_idx += 1
    return res_data


def read_one_file(file_path):
    data = xlrd.open_workbook(file_path)
    data.sheet_names()
    data_sheet_idx = 0
    table = data.sheet_by_name(DATA_SHEET_NAME)
    return table


def read_all_file(file_path_list):
    all_data = {}
    for file_path in file_path_list:
        print(f'开始读取文件:{file_path}')
        file_name = os.path.basename(file_path)
        file_name = file_name.split('.')[0]
        data = read_one_file(file_path)
        city_name = file_name[:2]
        all_data[city_name] = data
    return all_data


def write_data(file_name, data):
    workbook = xlwt.Workbook()
    for sheet_name in data:
        sheet_data = data[sheet_name]
        worksheet = workbook.add_sheet(sheet_name)
        ncols = 0
        for row_num in range(len(sheet_data)):
            for col_num in range(len(sheet_data[row_num])):
                ncols = col_num
                worksheet.write(row_num, col_num, sheet_data[row_num][col_num])
        for i in range(ncols + 1):
            # Set the column width 256为一个字符宽度
            if i == 0:
                width = 256 * 5
            elif i == 1:
                width = 256 * 50
            else:
                width = 256 * 17
            worksheet.col(i).width = width

    workbook.save(file_name)


def print_heart(s):
    data_list = []
    for y in range(15, -15, -1):
        tmp_list = []
        for x in range(-30, 30):
            tmp_cal_value = ((x * 0.05) ** 2 + (y * 0.1) ** 2 - 1) ** 3 - (
                    x * 0.05) ** 2 * (y * 0.1) ** 3
            if tmp_cal_value <= 0:
                tmp_str = s[(x - y) % 4]
            else:
                tmp_str = ' '
            tmp_list.append(tmp_str)
        data_list.append(''.join(tmp_list))

    data_str = '\n'.join(data_list)
    print(data_str)


def main():
    # file_path = './'
    result_file_name = 'result.xls'
    file_path = os.path.abspath(os.path.dirname(__file__))
    # 删除生成的结果文件
    try:
        os.remove(result_file_name)
        print('删除了上次生成的文件')
    except Exception as e:
        print('不需要删除上次生成的文件')
    file_path_list = get_file_name_list(file_path)
    # print(file_path_list)
    if not file_path_list:
        print(f'请把excel文件放到目录（文件夹）：{file_path}')
    data = read_all_file(file_path_list)
    reserved_rows = 4
    reserved_cols = 2

    data = formate_data(data, reserved_rows, reserved_cols)
    write_data(result_file_name, data)


if __name__ == '__main__':
    # print_heart('oooo')
    main()
    print_heart('love')
    input('press Enter')

