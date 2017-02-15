# -*- coding: utf-8 -*-

"""
Description

update money_flow set is_delete=1 where room_id=24931;
update contract_segment set is_delete=1 where room_contract_id=12905;
update room_contract set is_delete=1 where id=12905;
update room set rent_status='empty' where id=24931;
"""

__author__ = 'TT'

import xlrd
import xlwt
import datetime
import os


def getwt(wh, wm):
    if wm > 0:
        return wh, wm
    wh = wh - 1
    wm = wm + 60
    return getwt(wh, wm)


tar_file = '1512_output.xls'
get_file = 'a.xlsx'

# tar_content = xlwt.Workbook()
# tar_sheet = tar_content.add_sheet('sheet1', cell_overwrite_ok=True)
# tar_row = [u'员工编号', u'姓名', u'日期', u'上班', u'下班', u'是否加班', u'是否需人工审核']
# for i in range(0, len(tar_row)):
#     tar_sheet.write(0, i, tar_row[i])


data = xlrd.open_workbook(get_file)
print data.sheet_names()
table = data.sheets()[0]

# d = table.row_value(4)
# print d

row_num = table.nrows
col_num = table.ncols
# 1, 2, 4, 5, 6
print(row_num)
d = table.row_values(4)
print d
# for i in range(1, row_num):
#     print(i)
#     row_data = table.row_values(i)
#     # for d in range(0, col_num):
#     is_plus = u''
#     is_check = u''
#     try:
#         e_no = row_data[0]
#     except:
#         e_no = ''
#     try:
#         e_name = row_data[1]
#     except:
#         e_name = ''
#     try:
#         e_date = row_data[3]
#     except:
#         e_date = ''
#     try:
#         e_st = row_data[4]
#     except:
#         e_st = ''
#     try:
#         e_en = row_data[5]
#     except:
#         e_en = ''
#     if e_st != u'休息':
#         print(u'休息')
#         if e_st != '':
#             if e_en != '':
#                 try:
#                     # eh, em = [int(i) for i in e_en.split(':')]
#                     eh, em = e_en.split(':')
#                     if int(eh) > 19:
#                         print(eh, 0000)
#                         is_plus = u'是'
#                     elif int(eh) == 19:
#                         if int(em) > 55:
#                             print(eh, em, 1111)
#                             is_plus = u'是'
#                 except:
#                     is_check = u'是'
#             else:
#                 is_check = u'是'
#         else:
#             is_check = u'是'
#
#     tar_sheet.write(i, 0, e_no)
#     tar_sheet.write(i, 1, e_name)
#     tar_sheet.write(i, 2, e_date)
#     tar_sheet.write(i, 3, e_st)
#     tar_sheet.write(i, 4, e_en)
#     tar_sheet.write(i, 5, is_plus)
#     tar_sheet.write(i, 6, is_check)
#     # break
#
# tar_content.save(tar_file)



