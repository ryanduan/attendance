# -*- coding: utf-8 -*-

"""
Create at 16/6/16
"""

__author__ = 'TT'

import xlrd
import xlwt
import datetime
import os

# tar_row = [u'姓名',  u'部门', u'日期', u'上班', u'下班', u'次数', u'时长', u'状态']
# tar_content = xlwt.Workbook()
# tar_sheet = tar_content.add_sheet('sheet1', cell_overwrite_ok=True)
#
# for i in range(0, len(tar_row)):
#     tar_sheet.write(0, i, tar_row[i])
#
#
# get_file = '160704.xlsx'
# data = xlrd.open_workbook(get_file)
# table = data.sheets()[0]
#
# tar_file = '160704out.xls'
#
# row_num = table.nrows
# col_num = table.ncols
#
# get_title = ['time', 'name', 'depart', 'start', 'stop', 'count', 'long', 'status']
#
# # row_data = dict()
#
# no = u'异常'
# co = [u'未打卡', '未打卡', '--', u'--']
#
# row_count = 0
# for i in range(1, row_num):
#
#     row_data = table.row_values(i)
#     status = u'True'
#     if row_data[0] < '2016-06-16':
#         continue
#     row_count += 1
#     if row_data[3] in co or row_data[4] in co:
#         status = no
#     elif row_data[3] > '10:00':
#         status = no
#     elif row_data[4] < '18:00':
#         status = no
#     elif row_data[6] in co:
#         status = no
#     elif row_data[6][0] < 9:
#         status = no
#     else:
#         st_h, st_m = tuple(map(int, row_data[3].split(':')))
#         sp_h, sp_m = tuple(map(int, row_data[4].split(':')))
#         if sp_h - st_h < 9:
#             status = no
#         elif sp_h - st_h == 9:
#             if sp_m <= st_m:
#                 status = no
#     # if row_data[1] in row_data.keys():
#     #     row_data[row_data[1]].append(
#     #         (row_data[0], row_data[2], row_data[3], row_data[4], row_data[5], row_data[6], status))
#     # else:
#     #     row_data[row_data[1]] = [
#     #         (row_data[0], row_data[2], row_data[3], row_data[4], row_data[5], row_data[6], status)]
#     tar_sheet.write(row_count, 0, row_data[1])
#     tar_sheet.write(row_count, 1, row_data[2])
#     tar_sheet.write(row_count, 2, row_data[0])
#     tar_sheet.write(row_count, 3, row_data[3])
#     tar_sheet.write(row_count, 4, row_data[4])
#     tar_sheet.write(row_count, 5, row_data[5])
#     tar_sheet.write(row_count, 6, row_data[6])
#     tar_sheet.write(row_count, 7, status)
#
# tar_content.save(tar_file)


titles = ['name', 'thing', 'sick', 'year', 'paid', 'other', 'weekend', 'off',
          '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
          '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
          '21', '22', '23', '24', '25', '26', '27', '28', '29', '30']

leave_reason = dict(paternity=u'陪', thing=u'事', sick=u'病', year=u'年', paid=u'调',
                    off=u'旷', marriage=u'婚', maternity=u'产', other=u'其他')

t1 = ['name', 'reason', 'days' 'start', 'stop']

leave = '6leave.xlsx'
total = '6total.xlsx'

data = xlrd.open_workbook(leave)
table = data.sheets()[0]

tar_xls = xlrd.open_workbook(total)
tar_sheet = tar_xls.sheets()[0]

row_num = table.nrows
col_num = table.ncols

name_list = []

result = dict()

for i in range(row_num):
    res = ['', '', '', '', '', '', '', '',
           '', '', '', '', '', '', '', '', '', '',
           '', '', '', '', '', '', '', '', '', '',
           '', '', '', '', '', '', '', '', '', '']

    row_data = table.row_values(i)
    # print row_data
    name = row_data[0]

    reason = row_data[1]
    reason_tag = leave_reason.get(reason, None)
    if reason in ['paternity', 'marriage', 'maternity']:
        reason = 'other'
    # print row_data[2]
    long = float(row_data[2])
    start = row_data[3]
    stop = row_data[4]
    res[0] = name
    # print start
    # print  name
    start_date = start.split(' ')
    stop_date = stop.split(' ')
    # print start_date
    start_day = str(int(start_date[0].split('/')[-1]))
    stop_day = str(int(stop_date[0].split('/')[-1]))

    leave_range = [titles.index(start_day), titles.index(stop_day) + 1]
    leave_date = start.split(' ')
    if len(leave_date) > 1:
        tag_date, tag_noon = leave_date
    else:
        tag_date = leave_date
        tag_noon = None
    if long < 1:
        if reason_tag:
            if tag_noon and tag_noon == 'AM':
                reason_tag = u'{}/'.format(reason_tag)
            elif tag_noon:
                reason_tag = u'/{}'.format(reason_tag)

    if name not in name_list:
        old_reason_long = 0
        name_list.append(name)
    else:
        res = result.get(name)
        old_reason_long = res[titles.index(reason)]
        if old_reason_long == '':
            old_reason_long = 0
    res[titles.index(reason)] = long + old_reason_long

    for i in range(*leave_range):
        res[i] = reason_tag
    result[name] = res

tar_row = [u'姓名', u'事假', u'病假', u'年假', u'调休', u'其他', u'加班', u'旷工',
           '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
           '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
           '21', '22', '23', '24', '25', '26', '27', '28', '29', '30']
tar_content = xlwt.Workbook()
tar_sheet_2 = tar_content.add_sheet('sheet1', cell_overwrite_ok=True)

for i in range(0, len(tar_row)):
    tar_sheet_2.write(0, i, tar_row[i])

tar_file = '160705total.xls'

tar_row_num = tar_sheet.nrows
tar_col_num = tar_sheet.ncols

new_result = dict()

name_order = []
for i in range(tar_row_num):
    tar_row_data = tar_sheet.row_values(i)
    name_order.append(tar_row_data[0])
    if tar_row_data[0] in result.keys():
        v = result[tar_row_data[0]]
    # for k, v in result.items():
    #     if tar_row_data[0] == k:
        if len(tar_row_data) < len(v):
            for i in range(len(v) - len(tar_row_data)):
                tar_row_data.append('')
        for index, value in enumerate(v):
            if value:
                try:
                    tar_row_data[index] = value
                except:
                    tar_row_data.append(value)
    new_result[tar_row_data[0]] = tar_row_data

for idx, i in enumerate(name_order):
    v = new_result[i]
    for index, value in enumerate(v):
        tar_sheet_2.write(idx + 1, index, value)

tar_content.save(tar_file)

