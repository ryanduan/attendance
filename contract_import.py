# -*- coding: utf-8 -*-

"""
Create at 16/9/5

select id, uuid, start_time, end_time, amount, pay_status, pay_method, rent_deposit, rent_utilities, is_delete from house_rentorder where house_id=8297 and is_delete=0;

"""

__author__ = 'TT'

import xlrd
from xlrd.xldate import xldate_as_datetime
import copy


class ContractFileImport(object):
    """
    合同导入类
    """
    START_ROW = 3  # 开始行数，前三行是抬头，从第4行开始，0， 1， 2， 3， 所以等于 3
    TITLE_NUMBER = dict(  # 每一列的抬头
        city=0,  # 城市
        district=1,  # 区域
        block=2,  # 板块，商圈
        area=3,  # 小区
        address=4,  # 街道，地址
        building=5,  # 楼栋号
        unit_num=6,  # 单元号
        floor_num=7,  # 楼层
        house_num=8,  # 门牌号
        is_whole=9,  # 出租方式, 整租-合租
        room_num=10,  # 房间数
        owner_name=11,  # 业主姓名
        owner_phone=12,  # 业主电话
        owner_id_number=13,  # 业主身份证号
        house_start_time=14,  # 业主合同开始时间
        house_end_time=15,  # 业主合同结束时间
        house_pay_method_y=16,  # 业主合同押金押几
        house_pay_method_f=17,  # 业主合同房租付几
        house_advanced_days=18,  # 业主合同提前交租天数
        house_month_rental=19,  # 业主合同房租
        house_deposit=20,  # 业主合同押金
        room_name=21,  # 房间名
        customer_name=22,  # 租客姓名
        customer_phone=23,  # 租客电话
        customer_id_number=24,  # 租客身份证号
        room_start_time=25,  # 租客开始时间
        room_end_time=26,  # 租客结束时间
        room_pay_method_y=27,  # 租客押金押几
        room_pay_method_f=28,  # 租客房租付几
        room_advanced_days=29,  # 租客提前交租天数
        room_month_rental=30,  # 租客房租
        room_deposit=31,  # 租客押金
    )

    HOUSE_INFO = dict(
        city=0,  # 城市
        district=1,  # 区域
        block=2,  # 板块，商圈
        area=3,  # 小区
        address=4,  # 街道，地址
        building=5,  # 楼栋号
        unit_num=6,  # 单元号
        floor_num=7,  # 楼层
        house_num=8,  # 门牌号
        is_whole=9,  # 出租方式, 整租-合租
        room_num=10,  # 房间数
    )

    HOUSE_CONTRACT = dict(
        owner_name=11,  # 业主姓名
        owner_phone=12,  # 业主电话
        owner_id_number=13,  # 业主身份证号
        house_start_time=14,  # 业主合同开始时间
        house_end_time=15,  # 业主合同结束时间
        house_pay_method_y=16,  # 业主合同押金押几
        house_pay_method_f=17,  # 业主合同房租付几
        house_advanced_days=18,  # 业主合同提前交租天数
        house_month_rental=19,  # 业主合同房租
        house_deposit=20,  # 业主合同押金
    )
    ROOM_CONTRACT = dict(
        room_name=21,  # 房间名
        customer_name=22,  # 租客姓名
        customer_phone=23,  # 租客电话
        customer_id_number=24,  # 租客身份证号
        room_start_time=25,  # 租客开始时间
        room_end_time=26,  # 租客结束时间
        room_pay_method_y=27,  # 租客押金押几
        room_pay_method_f=28,  # 租客房租付几
        room_advanced_days=29,  # 租客提前交租天数
        room_month_rental=30,  # 租客房租
        room_deposit=31,  # 租客押金
    )

    REQUIRED_LIST = [
        # 房间信息必填项
        ['city', 'district', 'block', 'area', 'building',
         'floor_num', 'house_num', 'is_whole', 'room_num', ],
        # 业主合同信息必填项
        ['owner_name', 'house_start_time', 'house_end_time',
         'house_pay_method_y', 'house_pay_method_f',
         'house_advanced_days', 'house_month_rental', 'house_deposit', ],
        # 租客合同必填项
        ['room_name', ],
        ['customer_name', 'room_start_time', 'room_end_time',
         'room_pay_method_y', 'room_pay_method_f', 'room_advanced_days',
         'room_month_rental', 'room_deposit', ],
    ]
    TOTAL_LIST = [
        # 房间信息必填项
        ['city', 'district', 'block', 'area', 'building',
         'floor_num', 'house_num', 'is_whole', 'room_num',
         'area', 'unit_num', ],
        # 业主合同信息必填项
        ['owner_name', 'house_start_time', 'house_end_time',
         'house_pay_method_y', 'house_pay_method_f',
         'house_advanced_days', 'house_month_rental', 'house_deposit',
         'owner_phone', 'owner_id_number', ],
        # 租客合同必填项
        ['room_name', ],
        ['customer_name', 'room_start_time', 'room_end_time',
         'room_pay_method_y', 'room_pay_method_f', 'room_advanced_days',
         'room_month_rental', 'room_deposit',
         'customer_phone', 'customer_id_number', ],
    ]

    HIR = 0
    HCR = 1
    RIR = 2
    RCR = 3

    def __init__(self, contract_file=None, contract_content=None):
        self.contract_file = contract_file
        self.contract_content = contract_content
        self.table, self.row_num, self.col_num = self.open_contract_file()

    def open_contract_file(self):
        """"""
        data = xlrd.open_workbook(filename=self.contract_file, file_contents=self.contract_content)
        table = data.sheets()[0]
        row_num = table.nrows
        col_num = table.ncols
        return table, row_num, col_num

    def analysis_row(self):
        """分析每一行，获取数据"""
        # 是否包含房源和业主合同信息，取出房间数，跟开始行数相加
        house_row_count = copy.copy(self.START_ROW)
        result_data = []
        house_room_list = []
        house_room_data = {}
        for row_count in xrange(self.START_ROW, self.row_num):
            row_data = self.table.row_values(row_count)
            if house_row_count == row_count + 1:
                house_room_data['room_list'] = house_room_list
                result_data.append(house_room_data)
            if house_row_count == row_count:
                house_info = self.get_house_info(row_data, row_count)
                house_row_count += house_info.get('room_num', 0)
                house_contract = self.get_house_contract(row_data, row_count)
                house_room_list = []
                house_room_data = {}
                house_room_data['house_info'] = house_info
                house_room_data['house_contract'] = house_contract
            if house_row_count > row_count:
                room_contract = self.get_room_contract(row_data, row_count)
                house_room_list.append(room_contract)
        return result_data

    def get_single_data(self, data, name, idx=0, row=0, required=False):
        """get data from excel row data"""
        try:
            res = data[self.TITLE_NUMBER.get(name, idx)]
            if name.endswith('time'):
                if res:
                    res = xldate_as_datetime(res, False).date()
                    print res
                    res = str(res)
                    print res

            if required:
                if not res:
                    raise
        except Exception, e:
            print name
            print e
            raise ContractFileError(code=3505, row=row + 1, col=idx + 1)
        return res

    def get_house_info(self, data, row):
        """获取房源信息
        :param data : excel row data
        :param row : execl row count number
        """
        return dict(zip(self.HOUSE_INFO.keys(), [self.get_single_data(
            data, k, v, row, required=k in self.REQUIRED_LIST[self.HIR]) for k, v in self.HOUSE_INFO.items()]))

    def check_contract_info(self, data, row, contract_type):
        """"""
        value_list = filter(None, [data[k] for k in self.REQUIRED_LIST[contract_type]])
        if len(value_list) == 0:
            map(data.pop, self.TOTAL_LIST[contract_type])
            return True
        elif len(value_list) == len(self.REQUIRED_LIST[contract_type]):
            return False
        col = 0
        if contract_type == self.RCR:
            col = '22-31'
        elif contract_type == self.HCR:
            col = '11-20'
        raise ContractFileError(code=3505, row=row, col=col)

    def get_house_contract(self, data, row):
        """获取业主合同
        :param data : excel row data
        :param row : execl row count number
        """
        contract = dict(zip(self.HOUSE_CONTRACT.keys(), [self.get_single_data(
            data, k, v, row) for k, v in self.HOUSE_CONTRACT.items()]))
        if not self.check_contract_info(contract, row, self.HCR):
            contract_type = 'house'
            try:
                format_contract = dict(
                    start_time=contract.pop('{}_start_time'.format(contract_type)),
                    end_time=contract.pop('{}_end_time'.format(contract_type)),
                    advanced_days=int(contract.pop('{}_advanced_days'.format(contract_type))),
                    pay_method_y=int(contract.pop('{}_pay_method_y'.format(contract_type))),
                    pay_method_f=int(contract.pop('{}_pay_method_f'.format(contract_type))),
                    month_rental=float(contract.pop('{}_month_rental'.format(contract_type))),
                    deposit=float(contract.pop('{}_deposit'.format(contract_type))),
                )
            except:
                raise ContractFileError(code=3505, row=row, col='11-20')
            contract.update(format_contract)
        return contract

    def get_room_contract(self, data, row):
        """获取租客合同
        :param data : excel row data
        :param row : execl row count number
        """
        contract = dict(zip(self.ROOM_CONTRACT.keys(), [self.get_single_data(
            data, k, v, row, required=k in self.REQUIRED_LIST[self.RIR]) for k, v in self.ROOM_CONTRACT.items()]))
        empty = self.check_contract_info(contract, row, self.RCR)
        contract['rent_status'] = 'empty' if empty else 'rented'
        if contract['rent_status'] != 'empty':
            contract_type = 'room'
            try:
                format_contract = dict(
                    start_time=contract.pop('{}_start_time'.format(contract_type)),
                    end_time=contract.pop('{}_end_time'.format(contract_type)),
                    advanced_days=int(contract.pop('{}_advanced_days'.format(contract_type))),
                    pay_method_y=int(contract.pop('{}_pay_method_y'.format(contract_type))),
                    pay_method_f=int(contract.pop('{}_pay_method_f'.format(contract_type))),
                    month_rental=float(contract.pop('{}_month_rental'.format(contract_type))),
                    deposit=float(contract.pop('{}_deposit'.format(contract_type))),
                )
            except Exception, e:
                print e
                raise ContractFileError(code=3505, row=row, col='22-31')
            contract.update(format_contract)
        return contract


class ContractFileError(Exception):
    """合同导入文件异常"""

    error_msg = {
        # code: msg
        3501: '====',
        3505: '数据有误，请检查表格 {row}行 {col}列'
    }

    default_msg = '合同校验失败，请填写必填字段重新提交'

    def __init__(self, code=None, msg=None, *args, **kwargs):
        self.code = code
        if self.code == 3505:  # 缺少必填项，或者必填项不正确
            msg = self.error_msg.get(self.code).format(**kwargs)
        self.msg = msg or self.error_msg.get(self.code, self.default_msg)

    def __str__(self):
        return self.msg


if __name__ == '__main__':
    fn = 'a.xlsx'
    fc = open(fn, 'rb')
    cc = fc.read()
    c = ContractFileImport(contract_content=cc)
    try:
        res = c.analysis_row()
        for i in res:
            print i.get('house_contract')
            for j in i.get('room_list'):
                print j
    except ContractFileError, e:
        print e
    fc.close()
