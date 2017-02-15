# -*- coding: utf-8 -*-

"""
Create at 16/11/25

"√出勤  ○事假
△年假  ☆病假
●法定假/正常休假
■带薪假 ︻ 调休假
＊9：15迟到  ¤旷工
×长病假"

"""
import xlrd
import xlwt
import datetime
import os
from xlrd.xldate import xldate_as_datetime
from calendar import monthrange, weekday
from xlutils.copy import copy as cp
import copy

__author__ = 'TT'

att_mark = {
    u'出勤': u'√', u'事假': u'○',
    u'年假': u'△', u'病假': u'☆',
    u'法定假': u'●', u'长病假': u'×',
    u'带薪假': u'■', u'调休': u'︻',
    u'迟到': u'＊', u'旷工': u'¤',

    u'陪产假': u'■', u'产假': u'■',
    u'婚假': u'■',

}

error_name = set([])

start_day, end_day = '1.1', '1.31'

qj_sheet, kq_sheet, jb_sheet, cd_sheet, dk_sheet = None, None, None, None, None

att_file = '201701.xlsx'

data = xlrd.open_workbook(att_file)
for s in data.sheets():
    if u'请假' in s.name:
        qj_sheet = s
        continue
    if u'考勤' in s.name:
        kq_sheet = s
        continue
    if u'加班' in s.name:
        jb_sheet = s
        continue
    if u'迟到' in s.name:
        cd_sheet = s
        # if u'打卡' in s.name:
        #     dk_sheet = s

# # ----请假汇总开始----
# # 这里计算请假的数据

qj_start = 2


def get_weekday(*yd):
    y = 2017
    m = yd[-2]
    d = yd[-1]
    # print m, '---------------', int(start_day.split('.')[0][0]), m != int(start_day.split('.')[0][0])
    if m != int(start_day.split('.')[0]):
        return False
    if weekday(y, m, d) not in [5, 6]:
        return True
    return False


def split_dt(dt):

    res_list = []
    return_list = []
    if not isinstance(dt, float):
        l1 = [dt]
        l2 = []
        l3 = []
        for dt in l1:
            if u'；' in dt:
                l2.extend(dt.split(u'；'))
            else:
                l2.append(dt)
        for dt in l2:
            if ';' in dt:
                l3.extend(dt.split(';'))
            else:
                l3.append(dt)
        for dt in l3:
            if '-' in dt:
                ds, de = dt.split('-')
                if int(ds.split('.')[0]) < int(start_day.split('.')[0]):
                    ds = start_day
                if int(de.split('.')[0]) != int(end_day.split('.')[0]) and int(de.split('.')[0]) != int(ds.split('.')[0]):
                    de = end_day
                try:
                    for di in range(int(ds.split('.')[-1]), int(de.split('.')[-1]) + 1):
                        res_list.append('{}.{}'.format(ds.split('.')[0], str(di)))
                except Exception as e:
                    ds_m = None
                    de_m = None
                    if ds.endswith(u'（下午）'):
                        ds = ds.replace(u'（下午）', '')
                        ds_m = 'PM'
                    if de.endswith(u'（上午）'):
                        de = de.replace(u'（上午）', '')
                        de_m = 'AM'
                    for idx, di in enumerate(range(int(ds.split('.')[-1]), int(de.split('.')[-1]) + 1)):
                        if idx == 0 and ds_m:
                            res_list.append(u'{}.{}{}'.format(ds.split('.')[0], str(di), u'（下午）'))
                        elif de_m and idx == len(range(int(ds.split('.')[-1]), int(de.split('.')[-1]) + 1)) - 1:
                            res_list.append(u'{}.{}{}'.format(ds.split('.')[0], str(di), u'（上午）'))
                        else:
                            res_list.append('{}.{}'.format(ds.split('.')[0], str(di)))
            else:
                if dt:
                    res_list.append(dt)
        for dr in res_list:
            if dr.endswith(u'（上午）'):
                return_list.append([dr.replace(u'（上午）', ''), 'AM'])
            elif dr.endswith(u'(上午）'):
                return_list.append([dr.replace(u'(上午）', ''), 'AM'])
            elif dr.endswith(u'（下午）'):
                return_list.append([dr.replace(u'（下午）', ''), 'PM'])
            else:
                if dr:
                    return_list.append([dr, 'AM'])
                    return_list.append([dr, 'PM'])
    else:
        if dt:
            dt = str(dt)
            return_list.append([dt, 'AM'])
            return_list.append([dt, 'PM'])
    bill = []
    for i in return_list:
        yd = map(int, i[0].split('.'))
        # if yd[0] != 11:
        #     continue
        try:
            if get_weekday(*yd):
                bill.append(i)
        except Exception as e:
            import traceback
            print traceback.format_exc()
            print i[0]
            raise e
    if True:
        return bill


error_set = set([])


def ana_qj(data, col):
    """
    qj_title = [(u'序号', 0), (u'姓名', 1), (u'总计', -2), (u'备注', -1)]
    :return dict
    {
    name: TT,
    qj:[['11-1', 'A', u'☆'], ['11-2', 'P', u'︻']]
    }
    """
    try:
        qj_list = []
        for idx in range(2, col - 2, 2):
            dt_data = data[idx]
            if dt_data:
                dt_list = split_dt(dt_data)
                di = att_mark.get(data[idx + 1], None)
                for ana_dt in dt_list:
                    ana_dt.append(di)
                    qj_list.append(tuple(ana_dt))

        return {data[1]: qj_list}
    except:
        import traceback
        print traceback.format_exc()
        print data[1]
        print '===*===  这里请假异常了'
        error_name.add(data[1])
        return None


qj_data = {}
qj_count = 0

if qj_sheet is not None:
    qj_row_num = qj_sheet.nrows
    qj_col_num = qj_sheet.ncols

    for qj_row in range(2, qj_row_num):
        qj_row_data = qj_sheet.row_values(qj_row)
        ana_data = ana_qj(qj_row_data, qj_col_num)
        if ana_data is not None:
            qj_data.update(ana_data)

# # ----请假汇总结束----

# # ----加班开始----
# 这里计算加班的数据
jb_data = {}
jb_count = 0

jb_start = 2

if jb_sheet is not None:
    jb_row_num = jb_sheet.nrows
    jb_col_num = jb_sheet.ncols
    last_name = ''
    for jb_row in range(jb_start, jb_row_num):
        jb_list = []
        jb_row_data = jb_sheet.row_values(jb_row)
        name = jb_row_data[1]
        jb_date = jb_row_data[2]
        if isinstance(jb_date, float):
            jb_date = str(jb_date)
        else:
            if '-' in jb_date:
                jb_date = jb_date.split('-')[0]
            if jb_date.endswith(u'晚'):
                jb_date = jb_date.replace(u'晚', '')
            else:
                jb_date = jb_date.split(u'（')[0]
        if name:
            last_name = name
        else:
            name = last_name
        if jb_date.split('.')[0] != start_day.split('.')[0]:
            continue
        ann_jb = (jb_date, float(jb_row_data[3]))
        if name in jb_data.keys():
            jb_data[name].append(ann_jb)
        else:
            jb_data[name] = [ann_jb]
print jb_data
# # ----加班结束----

# # ----迟到开始----
cd_data = {}
cd_start = 2
cd_error = {}

if cd_sheet is not None:
    cd_rom_num = cd_sheet.nrows
    cd_col_num = cd_sheet.ncols

    for cd_row in range(cd_start, cd_rom_num):
        cd_row_data = cd_sheet.row_values(cd_row)
        name = cd_row_data[1]
        if name == u'下雨':
            continue
        if name == u'双十一不记迟到':
            continue
        cd_date = cd_row_data[2]
        if isinstance(cd_date, int) or isinstance(cd_date, float):
            cd_date = str(xldate_as_datetime(cd_date, False).date())
        mark_time = cd_row_data[3]
        if not mark_time:
            error_name.add(name)
            if name in cd_error.keys():
                cd_error[name].append(cd_date)
            else:
                cd_error[name] = [cd_date]
            error_set.add(name)
        else:
            if name in cd_data.keys():
                cd_data[name].append(cd_date)
            else:
                cd_data[name] = [cd_date]


# # ----迟到结束----

# ----考勤开始----


def ana_kq(data, m):
    """5-34
     u'出勤': u'√', u'事假': u'○',
    u'年假': u'△', u'病假': u'☆',
    u'法定假': u'●', u'长病假': u'×',
    u'带薪假': u'■', u'调休': u'︻',
    u'迟到': u'＊', u'旷工': u'¤',

    u'陪产假': u'■', u'产假': u'■',
    u'婚假': u'■',
    """
    but_days = [2, 24, 25, 26, 27, 28, 29, 30, 31]
    cd_days = []
    qj_days = []
    wo_days = []
    for idx, i in enumerate(data[5:35]):
        if idx + 1 in but_days:
            continue
        if m == 'WO':
            if i:
                wo_days.append((idx + 1, i))
        else:
            if m == 'AM':
                if i == u'＊':
                    cd_days.append(idx + 1)
                if i in [u'○', u'△', u'☆', u'■', u'︻', u'¤']:
                    qj_days.append((idx + 1, 'AM', i))
            else:
                if i in [u'○', u'△', u'☆', u'■', u'︻', u'¤']:
                    qj_days.append((idx + 1, 'PM', i))
    return cd_days, qj_days, wo_days


kq_start = 6

kq_error_data = []

kq_error_data.extend([kq_sheet.row_values(0), kq_sheet.row_values(1), kq_sheet.row_values(2),
                      kq_sheet.row_values(3), kq_sheet.row_values(4), kq_sheet.row_values(5)])

days_start = 4
print '------------------'
if kq_sheet is not None:
    kq_row_num = kq_sheet.nrows
    kq_col_num = kq_sheet.ncols
    for kq_row in range(kq_start, kq_row_num, 3):

        error_info = []
        try:
            am_data = kq_sheet.row_values(kq_row)
            pm_data = kq_sheet.row_values(kq_row + 1)
            wo_data = kq_sheet.row_values(kq_row + 2)
        except IndexError:
            break
        name = am_data[2]
        am_cd, am_qj, am_wo = ana_kq(am_data, 'AM')
        pm_cd, pm_qj, pm_wo = ana_kq(pm_data, 'PM')
        wo_cd, wo_qj, wo_wo = ana_kq(wo_data, 'WO')
        cd_error_list = []
        qj_error_list = []
        jb_error_list = []
        try:
            print 'ana chidao ----'
            print name
            if name in cd_data.keys():
                if len(am_cd) != len(cd_data[name]):
                    error_set.add(name)
                    print '迟到错误，两边长度不等1'
                    cd_error_list.append(u'迟到错误1')
                for day in cd_data[name]:
                    try:
                        dd = int(day.split('/')[-1])
                    except ValueError:
                        dd = int(day.split('-')[-1])
                    except:
                        dd = 0
                    if am_data[days_start + dd] != u'＊':
                        cd_error_list.append(u'{}迟到未标记2'.format(day))
                        error_set.add(name)
                        print day
                        print '迟到错误，未标记2'
                    if dd not in am_cd:
                        error_set.add(name)
                        cd_error_list.append(u'{}迟到标记错误3'.format(day))
                        print day
                        print '迟到标记错误3'
                for d in am_cd:
                    dd_l = []
                    for day in cd_data[name]:
                        try:
                            dd_l.append(int(day.split('/')[-1]))
                        except ValueError:
                            dd_l.append(int(day.split('-')[-1]))
                        except:
                            pass
                    if d not in dd_l:
                        error_set.add(name)
                        cd_error_list.append(u'{}迟到未标记4'.format(d))
                        print d
                        print '迟到未标记4'
            else:
                if am_cd:
                    error_set.add(name)
                    cd_error_list.append(u'迟到未标记5')
                    print '迟到未标记5'
            qj_list = am_qj + pm_qj
            print 'qingjia ----'
            if name in qj_data.keys():
                if len(qj_list) != len(qj_data[name]):
                    error_set.add(name)
                    qj_error_list.append(u'请假错误1')
                    print '请假长度不等1'
                    print qj_list
                    print qj_data[name]
                for day, t, m in qj_data[name]:
                    if t == 'AM':
                        if am_data[days_start + int(str(day).split('.')[-1])] not in [m, u'●']:
                            print day
                            print '上午错误3'
                            qj_error_list.append(u'{}上午请假错误3'.format(day))
                            error_set.add(name)
                    else:
                        if pm_data[days_start + int(day.split('.')[-1])] not in [m, u'●']:
                            print day
                            print '下午错误4'
                            qj_error_list.append(u'{}下午请假错误4'.format(day))
                            error_set.add(name)
            else:
                if qj_list:
                    if name == u'李晓敏':
                        print '----------------2'
                        print qj_list
                    error_set.add(name)
                    qj_error_list.append(u'请假错误5')
                    print '请假丢失5'
            print '加班 ----'

            if name in jb_data.keys():
                if len(wo_wo) != len(jb_data[name]):
                    error_set.add(name)
                    jb_error_list.append(u'加班错误1')
                    print '加班长度1'
                    print wo_wo
                    print jb_data[name]
                for day, h in jb_data[name]:
                    if wo_data[days_start + int(day.split('.')[-1])] != h:
                        jb_error_list.append(u'{}加班错误2'.format(day))
                        error_set.add(name)
                        print day
                        print '加班错误2'
            else:
                print name in jb_data.keys()
                if name == u'蒋燕':
                    print jb_data
                if wo_wo:
                    error_set.add(name)
                    jb_error_list.append(u'加班错误3')
                    print '加班缺失3'
        except:
            import traceback

            print traceback.format_exc()
            print name
            print '爆了异常'
            error_set.add(name)
            cd_error_list.append(u'需人工检查2')

        if name in error_set:
            cd_error_msg = ''
            if cd_error_list:
                cd_error_msg = u','.join(cd_error_list)
            qj_error_msg = ''
            if qj_error_list:
                qj_error_msg = u','.join(qj_error_list)
            jb_error_msg = ''
            if jb_error_list:
                jb_error_msg = u','.join(jb_error_list)
            if not cd_error_msg and not qj_error_msg and not jb_error_msg:
                cd_error_msg = u'需人工审核1'
            am_data.append(cd_error_msg)
            pm_data.append(qj_error_msg)
            wo_data.append(jb_error_msg)
            er_dt = [am_data, pm_data, wo_data]
            kq_error_data.extend(er_dt)
        # break

# target_file = '201611sh-out.xlsx'
target_file = att_file.replace('.xlsx', '-out-file-1.xls')

tar_content = xlwt.Workbook()
tar_sheet = tar_content.add_sheet('sheet1', cell_overwrite_ok=True)
# row_num = copy.copy(kq_start) - 1
for row_num, error_data in enumerate(kq_error_data):
    # for every_data in error_data:
    if error_data:
        for col_idx, col_data in enumerate(error_data):
            tar_sheet.write(row_num, col_idx, col_data)
    # print row_num, error_data
tar_content.save(target_file)
# tar_sheet.write(i, 0, e_no)

# rb = xlrd.open_workbook(target_file, encoding_override='utf8')
#
# wb = cp(rb)
# sheet = wb.get_sheet(0)
# row_num = copy.copy(kq_start) - 1
# error_col = 50
# print kq_error_data
# for error_data in kq_error_data:
#     print error_data
#     # for every_data in error_data:
#     if error_data:
#         row_num += 1
#         for col_idx, col_data in enumerate(error_data):
#             sheet.write(row_num, col_idx, col_data)
#
# wb.save(target_file)
